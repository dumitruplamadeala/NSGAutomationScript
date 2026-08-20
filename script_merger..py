#!/usr/bin/env python3
"""
Merge Azure NSG wave-1 workbook with catch-all difference workbook.

Usage:
    python revision-gen.py input.xlsx dif.xlsx merged_output.xlsx

What it does:
- Keeps the workbook/sheet format from input.xlsx.
- Processes every sheet that exists in both workbooks.
- Removes catch-all rows temporarily, merges normal rules, then appends catch-all rows back at the end.
- Updates an existing rule when destination IP + destination protocol + original destination ports match.
- Creates a new rule when that destination tuple does not exist.
- For updated rules, appends newly observed ports and source values without duplicating existing values.
- Prints logs showing each update/create and what ports/source values were added.
"""

from __future__ import annotations

import argparse
import copy
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
import copy
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import Any, Dict, List

INPUT_HEADERS = [
    "Rule Name",
    "Action",
    "Source Host Name (DNS) FQDN",
    "Source IP address/ Subnet / Range IP",
    "Destination Host Name (DNS) FQDN",
    "Destination IP address/ Subnet / Range IP",
    "Destination Protocol",
    "Destination port or service",
    "Rule Number",
    "Security Profile Group -FOR NSS USE ONLY",
    "Rule Type",
    "Temporary Rule Duration (in days)",
    "Bussiness Justification",
    "Additional Comments",
]

# Canonical names used internally.
COL_RULE_NAME = "rule_name"
COL_ACTION = "action"
COL_SRC_HOST = "source host name (dns) fqdn"
COL_SRC_IP = "source ip address/ subnet / range ip"
COL_DST_HOST = "dst_host"
COL_DST_IP = "destination ip address/ subnet / range ip"
COL_DST_PROTO = "dst_proto"
COL_DST_PORTS = "dst_ports"
COL_RULE_NUMBER = "rule_number"
COL_SECURITY_GROUP = "secruity profile group -for nss use only"
COL_RULE_TYPE = "rule_type"
COL_RULE_DURATION = "temporary rule duration (in days)"
COL_BUSINESS = "business"
COL_ADDITIONAL = "additional comments"
COL_EFFECTIVE_PORTS = "effective_ports"


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def normalize_scalar(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def split_list(value: Any) -> List[str]:
    """Split semicolon/comma/newline separated values, preserving first-seen order."""
    text = normalize_scalar(value)
    if not text:
        return []
    parts = re.split(r"[;,\n]+", text)
    out: List[str] = []
    seen = set()
    for part in parts:
        item = part.strip()
        if not item:
            continue
        key = item.lower()
        if key not in seen:
            seen.add(key)
            out.append(item)
    return out


def join_values(existing: Any, incoming: Any, separator: str = ";") -> str:
    """Append list-like values without duplicates."""
    result: List[str] = []
    seen = set()
    for value in list(split_list(existing)) + list(split_list(incoming)):
        key = value.lower()
        if key not in seen:
            seen.add(key)
            result.append(value)
    return f"{separator}".join(result)


def append_text(existing: Any, incoming: Any, separator: str = "; ") -> Any:
    """Append free text only when it is not already present."""
    old = normalize_scalar(existing)
    new = normalize_scalar(incoming)
    if not new:
        return existing
    if not old:
        return new
    if new.lower() in old.lower():
        return old
    return old.rstrip(" ;") + separator + new


def port_set(value: Any) -> Tuple[str, ...]:
    """Normalized sorted port tuple for matching."""
    ports = []
    for p in split_list(value):
        p = normalize_scalar(p)
        if p:
            ports.append(p)
    return tuple(
        sorted(
            set(ports), key=lambda x: (not x.isdigit(), int(x) if x.isdigit() else x)
        )
    )


def parse_port_ranges(value: Any) -> List[Tuple[int, int]]:
    """Return numeric ports/ranges as inclusive intervals."""
    ranges: List[Tuple[int, int]] = []
    items = value if isinstance(value, (list, tuple, set)) else split_list(value)
    for item in items:
        match = re.fullmatch(r"(\d+)\s*(?:-\s*(\d+))?", item)
        if not match:
            continue
        start = int(match.group(1))
        end = int(match.group(2) or match.group(1))
        if start > end:
            start, end = end, start
        ranges.append((start, end))
    return ranges


def _merge_intervals(intervals: Sequence[Tuple[int, int]]) -> List[Tuple[int, int]]:
    merged: List[Tuple[int, int]] = []
    for start, end in sorted(intervals):
        if merged and start <= merged[-1][1] + 1:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _missing_port_intervals(
    incoming: Sequence[Tuple[int, int]], existing: Sequence[Tuple[int, int]]
) -> List[Tuple[int, int]]:
    covered = _merge_intervals(existing)
    missing: List[Tuple[int, int]] = []
    for start, end in incoming:
        cursor = start
        for covered_start, covered_end in covered:
            if covered_end < cursor:
                continue
            if covered_start > end:
                break
            if covered_start > cursor:
                missing.append((cursor, min(end, covered_start - 1)))
            cursor = max(cursor, covered_end + 1)
            if cursor > end:
                break
        if cursor <= end:
            missing.append((cursor, end))
    return _merge_intervals(missing)


def _format_port_interval(interval: Tuple[int, int]) -> str:
    start, end = interval
    return str(start) if start == end else f"{start}-{end}"


def port_already_exists(incoming_port: str, existing_ports) -> bool:
    """Return whether an incoming numeric port/range is fully covered."""
    incoming = parse_port_ranges(incoming_port)
    if not incoming:
        return normalize_scalar(incoming_port).lower() in {
            normalize_scalar(item).lower() for item in existing_ports
        }
    return not _missing_port_intervals(incoming, parse_port_ranges(existing_ports))


def destination_key(
    src_ip: Any, dst_ip: Any, dst_proto: Any, dst_ports: Any
) -> Tuple[str, str, Tuple[str, ...]]:
    return (
        tuple(sorted(x.lower() for x in split_list(src_ip))),
        tuple(sorted(x.lower() for x in split_list(dst_ip))),
        normalize_scalar(dst_proto).upper(),
        port_set(dst_ports),
    )


def is_catch_all(row: Dict[str, Any]) -> bool:
    name = normalize_scalar(row.get(COL_RULE_NAME)).lower()
    rule_type = normalize_scalar(row.get(COL_RULE_TYPE)).lower()
    ports = normalize_scalar(row.get(COL_DST_PORTS))
    return name.startswith("catch_") or rule_type == "temporary" or ports == "*"


def is_app_rule(row: Dict[str, Any]) -> bool:
    rule_name = normalize_scalar(row.get(COL_RULE_NAME)).lower()

    return rule_name.startswith("app_")


def is_empty_row(row: Dict[str, Any]) -> bool:
    return all(normalize_scalar(v) == "" for v in row.values())


def copy_row_style(ws: Worksheet, source_row: int, target_row: int) -> None:
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy.copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.protection:
            dst.protection = copy.copy(src.protection)
        if src.alignment:
            dst.alignment = copy.copy(src.alignment)
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def clear_row_values(ws: Worksheet, row_idx: int) -> None:
    for col in range(1, ws.max_column + 1):
        ws.cell(row_idx, col).value = None


def read_headers(ws: Worksheet) -> List[str]:
    return [normalize_scalar(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]


def build_input_header_map(ws: Worksheet) -> Dict[str, int]:
    raw = [normalize_header(h) for h in read_headers(ws)]
    aliases = {
        COL_RULE_NAME: ["rule name"],
        COL_ACTION: ["action"],
        COL_SRC_HOST: ["source host name (dns) fqdn"],
        COL_SRC_IP: ["source ip address/ subnet / range ip"],
        COL_DST_HOST: ["destination host name (dns) fqdn"],
        COL_DST_IP: ["destination ip address/ subnet / range ip"],
        COL_DST_PROTO: ["destination protocol"],
        COL_DST_PORTS: ["destination port or service"],
        COL_RULE_NUMBER: ["rule number"],
        COL_SECURITY_GROUP: [
            "secruity profile group -for nss use only",
            "security profile group -for nss use only",
        ],
        COL_RULE_TYPE: ["rule type"],
        COL_RULE_DURATION: ["temporary rule duration (in days)"],
        COL_BUSINESS: ["bussiness justification", "business justification"],
        COL_ADDITIONAL: ["additional comments"],
    }
    result: Dict[str, int] = {}
    for canonical, names in aliases.items():
        for name in names:
            if name in raw:
                result[canonical] = raw.index(name) + 1
                break
    missing = [k for k in aliases if k not in result]
    if missing:
        raise ValueError(f"Missing expected columns in sheet {ws.title}: {missing}")
    return result


def build_diff_header_map(ws: Worksheet) -> Dict[str, int]:
    """
    dif.xlsx sample has duplicate/mislabeled headers:
      G = Destination Protocol
      H = Destination Protocol, but actually old destination ports
      I = Effective Destination Protocol, but actually newly observed/effective ports
    This function maps by position when those headers are duplicated.
    """
    input_like = build_input_header_map_forgiving(ws)
    result = dict(input_like)

    headers = [normalize_header(h) for h in read_headers(ws)]

    # In the diff workbook, columns G/H/I are protocol/original ports/effective ports.
    # Prefer this positional mapping because the sample has duplicate header names.
    if ws.max_column == 17:
        result[COL_DST_PROTO] = 7
        result[COL_DST_PORTS] = 8
        result[COL_EFFECTIVE_PORTS] = 9
        result[COL_ADDITIONAL] = 14
        result[COL_BUSINESS] = 11
    else:
        # Fallback: use destination port if available; effective ports default to destination ports.
        result[COL_EFFECTIVE_PORTS] = result.get(COL_DST_PORTS, 8)

    return result


def build_input_header_map_forgiving(ws: Worksheet) -> Dict[str, int]:
    """Header mapping for imperfect diff sheets."""
    headers = [normalize_header(h) for h in read_headers(ws)]
    result: Dict[str, int] = {}

    # Fixed positions from the shared format, with the diff's three special traffic columns.
    fixed = {
        COL_RULE_NAME: 1,
        COL_ACTION: 2,
        COL_SRC_HOST: 3,
        COL_SRC_IP: 4,
        COL_DST_HOST: 5,
        COL_DST_IP: 6,
        COL_DST_PROTO: 7,
        COL_DST_PORTS: 8,
        COL_RULE_NUMBER: 9,
        COL_SECURITY_GROUP: 10,
        COL_RULE_TYPE: 11,
        COL_RULE_DURATION: 12,
        COL_BUSINESS: 13,
        COL_ADDITIONAL: 14,
    }
    for key, idx in fixed.items():
        if idx <= ws.max_column:
            result[key] = idx

    # If a normal destination port header exists, use it unless the duplicate diff layout is present.
    for idx, header in enumerate(headers, start=1):
        if header == "destination port or service":
            result[COL_DST_PORTS] = idx

    return result


def row_to_dict(
    ws: Worksheet, row_idx: int, header_map: Dict[str, int]
) -> Dict[str, Any]:
    return {
        key: ws.cell(row_idx, col).value
        for key, col in header_map.items()
        if col <= ws.max_column
    }


def write_dict_to_row(
    ws: Worksheet, row_idx: int, header_map: Dict[str, int], row: Dict[str, Any]
) -> None:

    is_update = normalize_scalar(row.get(COL_ACTION)).lower() == "update"

    for key, col in header_map.items():

        if col > ws.max_column:
            continue

        # Source Host Name (DNS) FQDN
        if is_update and key == COL_SRC_HOST and row.get("_added_source_hosts"):
            ws.cell(row_idx, col).value = build_rich_text_list(
                row.get(COL_SRC_HOST), row.get("_added_source_hosts"), ", "
            )
            continue

        # Source IP address/ Subnet / Range IP
        if is_update and key == COL_SRC_IP and row.get("_added_source_ips"):
            ws.cell(row_idx, col).value = build_rich_text_list(
                row.get(COL_SRC_IP), row.get("_added_source_ips"), ", "
            )
            continue

        # Destination Port or Service
        if is_update and key == COL_DST_PORTS and row.get("_added_ports"):
            ws.cell(row_idx, col).value = build_rich_text_list(
                row.get(COL_DST_PORTS), row.get("_added_ports"), ", "
            )
            continue

        # Additional Comments
        if is_update and key == COL_ADDITIONAL and row.get("_added_comment"):

            full_comment = normalize_scalar(row.get(COL_ADDITIONAL))
            added_comment = normalize_scalar(row.get("_added_comment"))

            if added_comment and added_comment in full_comment:

                rich = CellRichText()

                pos = full_comment.rfind(added_comment)

                before = full_comment[:pos]
                after = full_comment[pos + len(added_comment) :]

                if before:
                    rich.append(before)

                rich.append(TextBlock(InlineFont(b=True), added_comment))

                if after:
                    rich.append(after)

                ws.cell(row_idx, col).value = rich
                continue

        # Default behavior (unchanged)
        ws.cell(row_idx, col).value = row.get(key)


def effective_new_ports(diff_row: Dict[str, Any]) -> str:
    """
    Return ports to apply to the target rule.

    If effective ports include old + new ports, only new ports are appended on update.
    For new rows, the effective value becomes the destination port/service.
    """
    return normalize_scalar(
        diff_row.get(COL_EFFECTIVE_PORTS) or diff_row.get(COL_DST_PORTS)
    )


def filter_text_segments_by_ports(text: Any, ports: Sequence[str]) -> Any:
    """Keep semicolon-separated description segments that mention one of the newly added ports."""
    raw = normalize_scalar(text)
    if not raw or not ports:
        return ""
    wanted = {normalize_scalar(p) for p in ports}
    segments = [seg.strip() for seg in raw.split(";") if seg.strip()]
    selected = []
    for seg in segments:
        numbers = set(re.findall(r"\b\d+\b", seg))
        if numbers & wanted:
            selected.append(seg)
    return "; ".join(selected) if selected else raw


def merge_existing_rule(
    base_row: Dict[str, Any], diff_row: Dict[str, Any]
) -> Tuple[Dict[str, Any], Dict[str, List[str]]]:

    updated = dict(base_row)

    changes: Dict[str, List[str]] = {
        "ports_added": [],
        "source_ips_added": [],
        "source_hosts_added": [],
    }

    existing_ports = split_list(updated.get(COL_DST_PORTS))
    incoming_ports = split_list(effective_new_ports(diff_row))
    new_only: List[str] = []

    numeric_incoming = parse_port_ranges(incoming_ports)
    numeric_existing = parse_port_ranges(existing_ports)
    new_only.extend(
        _format_port_interval(interval)
        for interval in _missing_port_intervals(numeric_incoming, numeric_existing)
    )

    existing_non_numeric = {
        item.lower() for item in existing_ports if not parse_port_ranges(item)
    }
    new_only.extend(
        item
        for item in incoming_ports
        if not parse_port_ranges(item)
        and item.lower() not in existing_non_numeric
    )
    if new_only:
        changes["ports_added"] = new_only
        updated[COL_DST_PORTS] = join_values(
            updated.get(COL_DST_PORTS), ";".join(new_only), ";"
        )

    # Source IPs
    old_src_ips = {x.lower() for x in split_list(updated.get(COL_SRC_IP))}
    incoming_src_ips = split_list(diff_row.get(COL_SRC_IP))

    changes["source_ips_added"] = [
        x for x in incoming_src_ips if x.lower() not in old_src_ips
    ]

    # Source Hosts
    old_src_hosts = {x.lower() for x in split_list(updated.get(COL_SRC_HOST))}
    incoming_src_hosts = split_list(diff_row.get(COL_SRC_HOST))

    if changes["source_ips_added"]:
        changes["source_hosts_added"] = [
            x for x in incoming_src_hosts if x.lower() not in old_src_hosts
        ]

        updated[COL_SRC_HOST] = join_values(
            updated.get(COL_SRC_HOST), diff_row.get(COL_SRC_HOST), ", "
        )
    else:
        changes["source_hosts_added"] = []

    # Merge source IPs
    updated[COL_SRC_IP] = join_values(
        updated.get(COL_SRC_IP), diff_row.get(COL_SRC_IP), ", "
    )

    updated["_added_source_ips"] = changes["source_ips_added"]
    updated["_added_source_hosts"] = changes["source_hosts_added"]
    updated["_added_ports"] = changes["ports_added"]

    # Business Justification
    updated[COL_BUSINESS] = append_text(
        updated.get(COL_BUSINESS),
        filter_text_segments_by_ports(diff_row.get(COL_BUSINESS), new_only),
        "; ",
    )

    # Additional Comments
    new_comment = normalize_scalar(diff_row.get(COL_ADDITIONAL))

    updated[COL_ADDITIONAL] = merge_additional_comments(
        updated.get(COL_ADDITIONAL), diff_row.get(COL_ADDITIONAL)
    )

    updated["_added_comment"] = new_comment

    # Only real traffic/source changes trigger Update
    has_real_update = (
        bool(changes["ports_added"])
        or bool(changes["source_ips_added"])
        or bool(changes["source_hosts_added"])
    )

    if has_real_update:
        updated[COL_ACTION] = "Update"
    else:
        updated[COL_ACTION] = base_row.get(COL_ACTION)

    return updated, changes


def make_new_rule(diff_row: Dict[str, Any]) -> Dict[str, Any]:
    new_row = {
        key: diff_row.get(key)
        for key in [
            COL_RULE_NAME,
            COL_ACTION,
            COL_SRC_HOST,
            COL_SRC_IP,
            COL_DST_HOST,
            COL_DST_IP,
            COL_DST_PROTO,
            COL_DST_PORTS,
            COL_RULE_NUMBER,
            COL_SECURITY_GROUP,
            COL_RULE_TYPE,
            COL_RULE_DURATION,
            COL_BUSINESS,
            COL_ADDITIONAL,
        ]
    }
    new_row[COL_ACTION] = "Create"
    new_row[COL_DST_PORTS] = effective_new_ports(diff_row)
    return new_row


def _normalized_value_set(value: Any) -> set[str]:
    return {item.lower() for item in split_list(value)}


def _ports_intersect(left: Any, right: Any) -> bool:
    left_ranges = parse_port_ranges(left)
    right_ranges = parse_port_ranges(right)
    for left_start, left_end in left_ranges:
        for right_start, right_end in right_ranges:
            if max(left_start, right_start) <= min(left_end, right_end):
                return True

    left_text = _normalized_value_set(left)
    right_text = _normalized_value_set(right)
    return bool(left_text & right_text)


def find_row_to_update(
    base_rows: Sequence[Dict[str, Any]],
    diff_row: Dict[str, Any],
    allow_destination_only: bool = False,
) -> Optional[int]:
    """Find the best existing rule for a same-destination merge."""
    new_sources = _normalized_value_set(diff_row.get(COL_SRC_IP))
    new_destinations = _normalized_value_set(diff_row.get(COL_DST_IP))
    new_protocol = normalize_scalar(diff_row.get(COL_DST_PROTO)).upper()
    new_ports = effective_new_ports(diff_row)

    candidates = []
    for row_idx, row in enumerate(base_rows):
        if is_catch_all(row):
            continue
        same_destination = (
            _normalized_value_set(row.get(COL_DST_IP)) == new_destinations
        )
        same_protocol = normalize_scalar(row.get(COL_DST_PROTO)).upper() == new_protocol
        if not (same_destination and same_protocol):
            continue

        existing_sources = _normalized_value_set(row.get(COL_SRC_IP))
        same_sources = existing_sources == new_sources
        ports_overlap = _ports_intersect(row.get(COL_DST_PORTS), new_ports)

        # Normal sheets require one stable traffic dimension. CoreRules may
        # intentionally consolidate all same-destination traffic.
        if not (allow_destination_only or same_sources or ports_overlap):
            continue

        score = (
            int(same_sources),
            int(ports_overlap),
            len(existing_sources & new_sources),
            -row_idx,
        )
        candidates.append((score, row_idx))

    return max(candidates)[1] if candidates else None


def process_sheet(base_ws: Worksheet, diff_ws: Worksheet) -> Dict[str, int]:
    base_map = build_input_header_map(base_ws)
    diff_map = build_diff_header_map(diff_ws)
    stats = {"updated": 0, "created": 0, "unchanged": 0}
    print(f"\n=== Processing sheet: {base_ws.title} ===")

    is_core_rules = base_ws.title.strip().lower() == "corerules"

    normal_style_row = 2 if base_ws.max_row >= 2 else 1

    base_rows: List[Dict[str, Any]] = []
    catch_rows: List[Dict[str, Any]] = []

    for r in range(2, base_ws.max_row + 1):
        row = row_to_dict(base_ws, r, base_map)
        if is_empty_row(row):
            continue
        if is_catch_all(row):
            catch_rows.append(row)
        else:
            base_rows.append(row)

    index: Dict[Tuple[str, str, Tuple[str, ...]], int] = {}
    app_index: Dict[str, int] = {}
    for i, row in enumerate(base_rows):
        index[
            destination_key(
                row.get(COL_SRC_IP),
                row.get(COL_DST_IP),
                row.get(COL_DST_PROTO),
                row.get(COL_DST_PORTS),
            )
        ] = i
        if is_app_rule(row):
            app_index[normalize_scalar(row.get(COL_RULE_NAME)).lower()] = i
    max_rule_number = max(
        [item["rule_number"] for item in base_rows if item["rule_number"] < 1500]
    )

    for r in range(2, diff_ws.max_row + 1):
        diff_row = row_to_dict(diff_ws, r, diff_map)
        if is_empty_row(diff_row) or is_catch_all(diff_row):
            continue
        #
        # APP rules:
        # Match by rule name only
        #
        if is_app_rule(diff_row):

            app_name = normalize_scalar(diff_row.get(COL_RULE_NAME)).lower()
            if app_name in app_index:
                idx = app_index[app_name]
                old_row = base_rows[idx]
                old_ports_value = normalize_scalar(old_row.get(COL_DST_PORTS))
                old_src_ip_value = normalize_scalar(old_row.get(COL_SRC_IP))
                old_src_host_value = normalize_scalar(old_row.get(COL_SRC_HOST))
                updated_row, changes = merge_existing_rule(old_row, diff_row)
                base_rows[idx] = updated_row

                if (
                    changes["ports_added"]
                    or changes["source_ips_added"]
                    or changes["source_hosts_added"]
                ):
                    stats["updated"] += 1
                    print(
                        f"[UPDATE APP] "
                        f"Sheet={base_ws.title} "
                        f"Rule={normalize_scalar(updated_row.get(COL_RULE_NAME))}"
                    )

                else:
                    stats["unchanged"] += 1
                    print(
                        f"[NO CHANGE APP] "
                        f"Sheet={base_ws.title} "
                        f"Rule={normalize_scalar(updated_row.get(COL_RULE_NAME))}"
                    )

                continue

            else:

                new_row = make_new_rule(diff_row)
                max_rule_number += 1
                new_row["rule_number"] = max_rule_number
                base_rows.append(new_row)
                app_index[app_name] = len(base_rows) - 1
                stats["created"] += 1
                print(
                    f"[CREATE APP] "
                    f"Sheet={base_ws.title} "
                    f"Rule={normalize_scalar(new_row.get(COL_RULE_NAME))}"
                )
                continue
        idx = find_row_to_update(base_rows, diff_row, is_core_rules)

        # import pprint
        # print("#" * 80)
        # pprint.pprint(index)
        # print("#" * 80)
        # # exit()

        if idx is not None:
            old_row = base_rows[idx]

            old_ports_value = normalize_scalar(old_row.get(COL_DST_PORTS))
            old_src_ip_value = normalize_scalar(old_row.get(COL_SRC_IP))
            old_src_host_value = normalize_scalar(old_row.get(COL_SRC_HOST))

            updated_row, changes = merge_existing_rule(old_row, diff_row)
            base_rows[idx] = updated_row

            if (
                changes["ports_added"]
                or changes["source_ips_added"]
                or changes["source_hosts_added"]
            ):
                stats["updated"] += 1
                print(
                    f"[UPDATE] Sheet={base_ws.title} Rule={normalize_scalar(updated_row.get(COL_RULE_NAME)) or '<no rule name>'}"
                )
                print(
                    f"         Destination IP={normalize_scalar(updated_row.get(COL_DST_IP))} Protocol={normalize_scalar(updated_row.get(COL_DST_PROTO))}"
                )
                if changes["ports_added"]:
                    print(f"         Added ports: {';'.join(changes['ports_added'])}")
                    print(
                        f"         Ports: {old_ports_value} -> {normalize_scalar(updated_row.get(COL_DST_PORTS))}"
                    )
                if changes["source_ips_added"]:
                    print(
                        f"         Added source IPs: {', '.join(changes['source_ips_added'])}"
                    )
                    formatted_source_ips = format_updated_source_ips(
                        updated_row.get(COL_SRC_IP), changes["source_ips_added"]
                    )
                    print(
                        f"         Source IPs: "
                        f"{old_src_ip_value} -> {formatted_source_ips}"
                    )
                if changes["source_hosts_added"]:
                    print(
                        f"         Added source hosts: {', '.join(changes['source_hosts_added'])}"
                    )
                    print(
                        f"         Source hosts: {old_src_host_value} -> {normalize_scalar(updated_row.get(COL_SRC_HOST))}"
                    )
            else:
                stats["unchanged"] += 1
                print(
                    f"[NO CHANGE] Sheet={base_ws.title} Rule={normalize_scalar(updated_row.get(COL_RULE_NAME)) or '<no rule name>'} already contains requested source/ports"
                )
                print(changes)

        else:
            new_row = make_new_rule(diff_row)
            max_rule_number += 1
            new_row["rule_number"] = max_rule_number
            base_rows.append(new_row)
            index[
                destination_key(
                    new_row.get(COL_SRC_IP),
                    new_row.get(COL_DST_IP),
                    new_row.get(COL_DST_PROTO),
                    new_row.get(COL_DST_PORTS),
                )
            ] = (
                len(base_rows) - 1
            )
            stats["created"] += 1
            print(
                f"[CREATE] Sheet={base_ws.title} Rule={normalize_scalar(new_row.get(COL_RULE_NAME)) or '<no rule name>'}"
            )
            print(
                f"         Destination IP={normalize_scalar(new_row.get(COL_DST_IP))} Protocol={normalize_scalar(new_row.get(COL_DST_PROTO))} Ports={normalize_scalar(new_row.get(COL_DST_PORTS))}"
            )


    final_rows = base_rows + catch_rows

    for r in range(2, base_ws.max_row + 1):
        clear_row_values(base_ws, r)

    required_rows = len(final_rows) + 1
    while base_ws.max_row < required_rows:
        base_ws.append([None] * base_ws.max_column)

    if base_ws.max_row > required_rows:
        base_ws.delete_rows(required_rows + 1, base_ws.max_row - required_rows)

    for idx, row in enumerate(final_rows, start=2):
        copy_row_style(base_ws, normal_style_row, idx)
        write_dict_to_row(base_ws, idx, base_map, row)

    print(
        f"--- Sheet summary {base_ws.title}: updated={stats['updated']}, created={stats['created']}, unchanged={stats['unchanged']} ---"
    )
    return stats


def merge_workbooks(input_path: Path, diff_path: Path, output_path: Path) -> None:
    wb = load_workbook(input_path)
    diff_wb = load_workbook(diff_path, data_only=True)

    total = {"updated": 0, "created": 0, "unchanged": 0}
    for sheet_name in wb.sheetnames:
        if sheet_name not in diff_wb.sheetnames:
            print(
                f"\n=== Skipping sheet: {sheet_name} - not found in diff workbook ==="
            )
            continue
        stats = process_sheet(wb[sheet_name], diff_wb[sheet_name])
        for key in total:
            total[key] += stats[key]

    wb.save(output_path)
    print("\n============================================================")
    print("SUMMARY")
    print("============================================================")
    print(f"Updated rules : {total['updated']}")
    print(f"Created rules : {total['created']}")
    print(f"Unchanged hits: {total['unchanged']}")
    print(f"Output file   : {output_path}")


def format_updated_source_ips(all_source_ips: Any, newly_added: Sequence[str]) -> str:
    """
    Format source IP list and make only newly added source IPs bold.
    """
    added_set = {ip.lower() for ip in newly_added}

    formatted = []
    for ip in split_list(all_source_ips):
        if ip.lower() in added_set:
            formatted.append(f"**{ip}**")
        else:
            formatted.append(ip)

    return ", ".join(formatted)


def build_rich_text_list(
    full_value: Any, added_values: Sequence[str], separator: str = ", "
):
    """
    Build rich text where only newly added values are bold.
    """
    added = {x.lower() for x in added_values}

    rich = CellRichText()
    first = True

    for item in split_list(full_value):
        if not first:
            rich.append(separator)

        if item.lower() in added:
            rich.append(TextBlock(InlineFont(b=True), item))
        else:
            rich.append(item)

        first = False

    return rich


def merge_additional_comments(existing: Any, incoming: Any) -> str:
    """
    Merge Additional Comments.

    Duplicate CI/IPR/IP entries are collapsed.
    If one version contains ASPIED information, keep that one.
    Otherwise keep the longest version.
    """

    def split_entries(value: Any) -> List[str]:
        text = normalize_scalar(value)
        if not text:
            return []

        # Comments may use newlines, commas, or semicolons. Split only when
        # the next segment starts a recognizable entry, so punctuation inside
        # ordinary prose remains intact.
        entry_start = (
            r"(?:CI\d+\b|IPR\d+\b|"
            r"(?:\d{1,3}\.){3}\d{1,3}(?:/\d+)?\b|"
            r"[A-Za-z0-9][A-Za-z0-9.-]*\.[A-Za-z]{2,}\b)"
        )
        parts = re.split(rf"(?:[\r\n]+|[,;]\s*)(?={entry_start})", text, flags=re.I)
        return [part.strip(" ,;") for part in parts if part.strip(" ,;")]

    def get_key(entry: str) -> str:

        m = re.match(r"^\s*(CI\d+)\b", entry, re.I)
        if m:
            return m.group(1).upper()

        m = re.match(r"^\s*(IPR\d+)\b", entry, re.I)
        if m:
            return m.group(1).upper()

        m = re.match(r"^\s*((?:\d{1,3}\.){3}\d{1,3}(?:/\d+)?)", entry)
        if m:
            return m.group(1)

        m = re.match(r"^\s*([A-Za-z0-9][A-Za-z0-9.-]*\.[A-Za-z]{2,})\b", entry)
        if m:
            return m.group(1).lower()

        m = re.match(r"^\s*(OneCloud Subnet.*)$", entry, re.I)
        if m:
            return m.group(1).lower()

        return entry.lower()

    def score(entry: str) -> int:

        score = len(entry)

        if "ASPIED" in entry.upper():
            score += 10000

        if "RITM" in entry.upper():
            score += 5000

        return score

    merged = {}

    for entry in split_entries(existing):
        key = get_key(entry)

        if key not in merged:
            merged[key] = entry
        elif score(entry) > score(merged[key]):
            merged[key] = entry

    for entry in split_entries(incoming):
        key = get_key(entry)

        if key not in merged:
            merged[key] = entry
        elif score(entry) > score(merged[key]):
            merged[key] = entry

    return "\n".join(merged.values())


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge Azure NSG input workbook with diff workbook for second-wave NSG implementation."
    )
    parser.add_argument(
        "input", type=Path, help="First-wave NSG workbook, e.g. input.xlsx"
    )
    parser.add_argument(
        "diff",
        type=Path,
        help="Difference workbook from catch-all traffic, e.g. dif.xlsx",
    )
    parser.add_argument(
        "output", type=Path, help="Output workbook path, e.g. second_wave.xlsx"
    )
    args = parser.parse_args()

    merge_workbooks(args.input, args.diff, args.output)


if __name__ == "__main__":
    main()
