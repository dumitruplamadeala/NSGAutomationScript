#!/usr/bin/env python3
"""Aggregate NSG flow-log TSV chunks without loading all rows into memory."""

from __future__ import annotations

import argparse
import csv
import json
import sqlite3
from pathlib import Path
from typing import Dict, Optional


EXPECTED_COLUMNS = [
    "SrcEff",
    "DestEff",
    "DestPort",
    "L4Protocol",
    "Server",
    "SrcPorts",
    "Requests",
]


def open_tsv_text(path: Path):
    with path.open("rb") as handle:
        prefix = handle.read(2)

    encoding = "utf-16" if prefix in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
    return path.open("r", encoding=encoding, newline="")


def aggregate(
    chunk_directory: Path,
    output_csv: Path,
    database_path: Path,
    summary_json: Optional[Path],
) -> Dict[str, int]:
    chunk_files = sorted(chunk_directory.glob("*.tsv"))
    if not chunk_files:
        raise RuntimeError(f"No chunk CSV files found in: {chunk_directory}")

    database_path.parent.mkdir(parents=True, exist_ok=True)
    output_csv.parent.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(database_path) as connection:
        connection.execute("PRAGMA journal_mode = WAL")
        connection.execute("PRAGMA synchronous = NORMAL")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS flows (
                SrcEff TEXT NOT NULL,
                DestEff TEXT NOT NULL,
                DestPort TEXT NOT NULL,
                L4Protocol TEXT NOT NULL,
                Server TEXT NOT NULL,
                SrcPorts TEXT NOT NULL,
                Requests INTEGER NOT NULL,
                PRIMARY KEY (
                    SrcEff, DestEff, DestPort, L4Protocol, Server, SrcPorts
                )
            )
            """
        )

        processed = 0
        for chunk_file in chunk_files:
            with open_tsv_text(chunk_file) as handle:
                reader = csv.reader(handle, delimiter="\t")

                batch = []
                for line_number, row in enumerate(reader, start=1):
                    if not row or (len(row) == 1 and not row[0].strip()):
                        continue

                    if row == ["PrimaryResult"]:
                        continue

                    if len(row) == len(EXPECTED_COLUMNS) + 1 and row[-1] == "PrimaryResult":
                        row = row[:-1]

                    if len(row) != len(EXPECTED_COLUMNS):
                        raise RuntimeError(
                            f"Unexpected column count in {chunk_file.name} line "
                            f"{line_number}: {len(row)}, expected {len(EXPECTED_COLUMNS)}"
                        )

                    batch.append(
                        (
                            row[0] or "",
                            row[1] or "",
                            row[2] or "",
                            row[3] or "",
                            row[4] or "",
                            row[5] or "",
                            int(row[6] or 0),
                        )
                    )
                    processed += 1

                    if len(batch) >= 10_000:
                        connection.executemany(
                            """
                            INSERT INTO flows VALUES (?, ?, ?, ?, ?, ?, ?)
                            ON CONFLICT(
                                SrcEff, DestEff, DestPort,
                                L4Protocol, Server, SrcPorts
                            ) DO UPDATE SET
                                Requests = Requests + excluded.Requests
                            """,
                            batch,
                        )
                        connection.commit()
                        batch.clear()

                if batch:
                    connection.executemany(
                        """
                        INSERT INTO flows VALUES (?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(
                            SrcEff, DestEff, DestPort,
                            L4Protocol, Server, SrcPorts
                        ) DO UPDATE SET
                            Requests = Requests + excluded.Requests
                        """,
                        batch,
                    )
                    connection.commit()

        with output_csv.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow([
                "SrcEff", "DestEff", "DestPort", "L4Protocol",
                "Server", "SrcPorts", "Requests",
            ])

            cursor = connection.execute(
                """
                SELECT SrcEff, DestEff, DestPort, L4Protocol,
                       Server, SrcPorts, Requests
                FROM flows
                ORDER BY Requests DESC
                """
            )
            exported = 0
            for result_row in cursor:
                writer.writerow(result_row)
                exported += 1

        final_count = connection.execute("SELECT COUNT(*) FROM flows").fetchone()[0]

    summary = {
        "processed_chunk_rows": int(processed),
        "final_aggregated_rows": int(final_count),
        "final_exported_rows": int(exported),
    }

    if summary_json is not None:
        summary_json.parent.mkdir(parents=True, exist_ok=True)
        summary_json.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    print(f"Processed chunk rows: {processed}")
    print(f"Final aggregated rows: {final_count}")
    print(f"Final exported rows: {exported}")
    print(f"Final CSV: {output_csv}")
    return summary


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--chunk-directory", required=True, type=Path)
    parser.add_argument("--output-csv", required=True, type=Path)
    parser.add_argument("--database", required=True, type=Path)
    parser.add_argument("--summary-json", type=Path)
    args = parser.parse_args()

    aggregate(
        args.chunk_directory,
        args.output_csv,
        args.database,
        args.summary_json,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
