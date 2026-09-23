#!/usr/bin/env python3
from __future__ import annotations

import csv
import shutil
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "database.xlsx"
PENDING = ROOT / "updates" / "pending"
APPLIED = ROOT / "updates" / "applied" / date.today().isoformat()

SHEET_MAP = {
    "Venues": "Venues",
    "Booking_Agents": "Booking Agents",
    "Peer_Bands": "Peer Bands",
}

def norm(v: object) -> str:
    if v is None:
        return ""
    return str(v).strip().casefold()

def row_sig(row: list[object]) -> tuple[str, ...]:
    return tuple(norm(v) for v in row)

def find_header_row(ws, expected: list[str]) -> int:
    expected_sig = tuple(norm(x) for x in expected)
    for row_no in range(1, min(ws.max_row, 10) + 1):
        values = [ws.cell(row_no, col).value for col in range(1, len(expected) + 1)]
        if tuple(norm(v) for v in values) == expected_sig:
            return row_no
    raise RuntimeError(f"Could not find expected header in sheet {ws.title!r}")

def load_csv(path: Path) -> tuple[list[str], list[list[str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))
    if not rows:
        raise RuntimeError(f"Empty CSV: {path}")
    return rows[0], rows[1:]

def main() -> None:
    if not DB.exists():
        raise RuntimeError("database.xlsx is missing")

    wb = load_workbook(DB)
    changed = False
    applied: list[tuple[Path, int]] = []

    for csv_path in sorted(PENDING.glob("*.csv")):
        prefix = csv_path.name.split("__", 1)[0]
        sheet = SHEET_MAP.get(prefix)
        if not sheet:
            continue
        if sheet not in wb.sheetnames:
            raise RuntimeError(f"Missing sheet {sheet!r}")

        header, data = load_csv(csv_path)
        ws = wb[sheet]
        header_row = find_header_row(ws, header)
        start_data_row = header_row + 1

        existing = set()
        for row in ws.iter_rows(min_row=start_data_row, max_row=ws.max_row, max_col=len(header), values_only=True):
            existing.add(row_sig(list(row)))

        added = 0
        for raw in data:
            row = raw[: len(header)] + [""] * max(0, len(header) - len(raw))
            sig = row_sig(row)
            if sig in existing:
                continue
            ws.append(row)
            existing.add(sig)
            added += 1
            changed = True

        applied.append((csv_path, added))

    if changed:
        wb.save(DB)

    if applied:
        APPLIED.mkdir(parents=True, exist_ok=True)
        for path, _ in applied:
            shutil.move(str(path), str(APPLIED / path.name))

    print(f"pending_files={len(applied)} changed={changed}")
    for path, added in applied:
        print(f"{path.name}: added={added}")

if __name__ == "__main__":
    main()
