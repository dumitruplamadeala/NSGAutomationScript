#!/usr/bin/env python3
"""Aggregate NSG flow-log CSV chunks without loading all rows into memory."""

from __future__ import annotations

import argparse
import csv
import json
import sqlite3
from pathlib import Path
from typing import Dict, Optional


REQUIRED_COLUMNS = {
    "SrcEff",
    "DestEff",
    "DestPort",
    "L4Protocol",
    "Server",
    "SrcPorts",
    "Requests",
}


def aggregate(
    chunk_directory: Path,
    output_csv: Path,
    database_path: Path,
    row_limit: int,
    summary_json: Optional[Path],
) -> Dict[str, int]:
    chunk_files = sorted(chunk_directory.glob("*.csv"))
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
            with chunk_file.open("r", encoding="utf-8-sig", newline="") as handle:
                reader = csv.DictReader(handle)
                columns = set(reader.fieldnames or [])
                missing = REQUIRED_COLUMNS - columns
                if missing:
                    raise RuntimeError(
                        f"Missing columns in {chunk_file.name}: {', '.join(sorted(missing))}"
                    )

                batch = []
                for row in reader:
                    batch.append(
                        (
                            row["SrcEff"] or "",
                            row["DestEff"] or "",
                            row["DestPort"] or "",
                            row["L4Protocol"] or "",
                            row["Server"] or "",
                            row["SrcPorts"] or "",
                            int(row["Requests"] or 0),
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

            limit_clause = ""
            query_parameters: tuple[int, ...] = ()
            if row_limit > 0:
                limit_clause = "LIMIT ?"
                query_parameters = (row_limit,)

            cursor = connection.execute(
                f"""
                SELECT SrcEff, DestEff, DestPort, L4Protocol,
                       Server, SrcPorts, Requests
                FROM flows
                ORDER BY Requests DESC
                {limit_clause}
                """,
                query_parameters,
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
    parser.add_argument("--row-limit", type=int, default=0)
    parser.add_argument("--summary-json", type=Path)
    args = parser.parse_args()

    if args.row_limit < 0:
        raise RuntimeError("--row-limit must be greater than or equal to zero")

    aggregate(
        args.chunk_directory,
        args.output_csv,
        args.database,
        args.row_limit,
        args.summary_json,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
