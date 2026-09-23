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
    "Beacons": "Beacons",
}

REMOVAL_HEADER = ["Sheet", "Key_Type", "Key", "Reason", "Source_URL", "Verified_Date"]

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
    applied: list[tuple[Path, int, str]] = []

    for csv_path in sorted(PENDING.glob("*.csv")):
        if csv_path.name.startswith("Removals__"):
            header, data = load_csv(csv_path)
            if header != REMOVAL_HEADER:
                raise RuntimeError(f"Invalid removals header in {csv_path}: {header}")
            removed = 0
            for raw in data:
                if len(raw) < len(REMOVAL_HEADER):
                    raise RuntimeError(f"Malformed removal row in {csv_path}: {raw}")
                target_sheet, key_type, key = raw[0], norm(raw[1]), norm(raw[2])
                if target_sheet not in wb.sheetnames:
                    raise RuntimeError(f"Removal references missing sheet {target_sheet!r}")
                ws = wb[target_sheet]
                headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
                idx = {norm(v): i + 1 for i, v in enumerate(headers) if v not in (None, "")}
                if key_type == "name":
                    cols = [idx.get("name")]
                    vals = [key]
                elif key_type == "email":
                    cols = [idx.get("email")]
                    vals = [key]
                elif key_type == "website":
                    cols = [idx.get("website")]
                    vals = [key]
                elif key_type == "name+city":
                    cols = [idx.get("name"), idx.get("city")]
                    parts = [norm(x) for x in raw[2].split("||", 1)]
                    vals = parts if len(parts) == 2 else []
                elif key_type == "name+country":
                    cols = [idx.get("name"), idx.get("country")]
                    parts = [norm(x) for x in raw[2].split("||", 1)]
                    vals = parts if len(parts) == 2 else []
                else:
                    raise RuntimeError(f"Unsupported removal key type {raw[1]!r}")
                if not cols or any(c is None for c in cols) or len(vals) != len(cols):
                    raise RuntimeError(f"Removal key cannot be resolved in {target_sheet!r}: {raw}")
                for row_no in range(ws.max_row, 1, -1):
                    if all(norm(ws.cell(row_no, col).value) == val for col, val in zip(cols, vals)):
                        ws.delete_rows(row_no, 1)
                        removed += 1
                        changed = True
            applied.append((csv_path, removed, "removed"))
            continue
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

        applied.append((csv_path, added, "added"))

    if changed:
        wb.save(DB)

    if applied:
        APPLIED.mkdir(parents=True, exist_ok=True)
        for path, _, _ in applied:
        shutil.move(str(path), str(APPLIED / path.name))

    print(f"pending_files={len(applied)} changed={changed}")
    for path, count, action in applied:
        print(f"{path.name}: {action}={count}")

if __name__ == "__main__":
    main()
