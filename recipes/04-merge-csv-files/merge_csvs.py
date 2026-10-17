#!/usr/bin/env python3
"""Merge every CSV in a folder into one file, matching columns by header name.

Usage:  python3 merge_csvs.py ./exports merged.csv
Standard library only, no installs needed.
"""
import csv
import sys
from pathlib import Path


def merge(folder: Path, out_path: Path) -> int:
    files = sorted(p for p in folder.glob("*.csv") if p.resolve() != out_path.resolve())
    headers, rows = [], []
    for path in files:
        with path.open(newline="", encoding="utf-8-sig") as f:  # utf-8-sig strips Excel's BOM
            for row in csv.DictReader(f):
                row = {(k or "").strip(): v for k, v in row.items() if k}
                if not any((v or "").strip() for v in row.values()):
                    continue  # skip blank lines
                for k in row:
                    if k not in headers:
                        headers.append(k)
                row["Source file"] = path.name
                rows.append(row)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=headers + ["Source file"], restval="")
        writer.writeheader()
        writer.writerows(rows)
    return len(rows)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("usage: merge_csvs.py <folder> <output.csv>")
    n = merge(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"wrote {n} rows to {sys.argv[2]}")
