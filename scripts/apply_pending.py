#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import re
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
CITY_PASS_PENDING_RE = re.compile(r"^.+__CityPass__(\d{4}-\d{2}-\d{2})__(\d{4})__.+\.csv$")
REJECTED_ROOT = ROOT / "updates" / "rejected"


def city_pass_pending_allowed(csv_path: Path) -> bool:
    match = CITY_PASS_PENDING_RE.match(csv_path.name)
    if not match:
        return True
    pass_file = ROOT / "city_passes" / f"{match.group(1)}__{match.group(2)}.json"
    if not pass_file.exists():
        return False
    try:
        payload = json.loads(pass_file.read_text(encoding="utf-8"))
        return int(payload.get("research_version", 0) or 0) >= 4 and not bool(payload.get("invalidated", False))
    except Exception:
        return False



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


def find_existing_header_row(ws) -> int:
    for row_no in range(1, min(ws.max_row, 10) + 1):
        values = [norm(ws.cell(row_no, col).value) for col in range(1, min(ws.max_column, 20) + 1)]
        if "name" in values or "email" in values:
            return row_no
    raise RuntimeError(f"Could not locate header row in sheet {ws.title!r}")


def load_csv(path: Path) -> tuple[list[str], list[list[str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))
    if not rows:
        raise RuntimeError(f"Empty CSV: {path}")
    return rows[0], rows[1:]


def apply_removals(wb, csv_path: Path) -> int:
    header, data = load_csv(csv_path)
    if [norm(x) for x in header] != [norm(x) for x in REMOVAL_HEADER]:
        raise RuntimeError(f"Invalid removals header in {csv_path}: {header}")

    removed = 0
    for raw in data:
        if not raw or not any(norm(x) for x in raw):
            continue
        if len(raw) < len(REMOVAL_HEADER):
            raise RuntimeError(f"Malformed removal row in {csv_path}: {raw}")

        target_sheet = raw[0].strip()
        key_type = norm(raw[1])
        key_raw = raw[2]
        key = norm(key_raw)

        if target_sheet not in wb.sheetnames:
            raise RuntimeError(f"Removal references missing sheet {target_sheet!r}")

        ws = wb[target_sheet]
        header_row = find_existing_header_row(ws)
        headers = [
            norm(ws.cell(header_row, c).value)
            for c in range(1, ws.max_column + 1)
        ]
        idx = {value: i + 1 for i, value in enumerate(headers) if value}

        if key_type == "name":
            required_cols = ["name"]
            values = [key]
        elif key_type == "email":
            required_cols = ["email"]
            values = [key]
        elif key_type == "website":
            required_cols = ["website"]
            values = [key]
        elif key_type == "name+city":
            required_cols = ["name", "city"]
            values = [norm(x) for x in key_raw.split("||", 1)]
        elif key_type == "name+country":
            required_cols = ["name", "country"]
            values = [norm(x) for x in key_raw.split("||", 1)]
        elif key_type == "name+city+url":
            parts = key_raw.split("||", 2)
            if len(parts) != 3:
                raise RuntimeError(f"Malformed name+city+url removal key: {raw}")
            if "destination_url" in idx:
                url_col = "destination_url"
            elif "source_url" in idx:
                url_col = "source_url"
            elif "website" in idx:
                url_col = "website"
            else:
                raise RuntimeError(f"No URL column available for exact removal in {target_sheet!r}")
            required_cols = ["name", "city", url_col]
            values = [norm(parts[0]), norm(parts[1]), norm(parts[2])]
        else:
            raise RuntimeError(f"Unsupported removal key type {raw[1]!r}")

        if len(values) != len(required_cols) or any(col not in idx for col in required_cols):
            raise RuntimeError(f"Removal key cannot be resolved in {target_sheet!r}: {raw}")

        for row_no in range(ws.max_row, header_row, -1):
            if all(
                norm(ws.cell(row_no, idx[col]).value) == value
                for col, value in zip(required_cols, values)
            ):
                ws.delete_rows(row_no, 1)
                removed += 1

    return removed


def main() -> None:
    if not DB.exists():
        raise RuntimeError("database.xlsx is missing")

    wb = load_workbook(DB)
    changed = False
    applied: list[tuple[Path, int, str]] = []

    rejected_dir = REJECTED_ROOT / date.today().isoformat()

    for csv_path in sorted(PENDING.glob("*.csv")):
        if not city_pass_pending_allowed(csv_path):
            rejected_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(csv_path), str(rejected_dir / csv_path.name))
            applied.append((csv_path, 0, "rejected-legacy-city-pass"))
            print(f"Rejected legacy city-pass pending file: {csv_path.name}")
            continue

        if csv_path.name.startswith("Removals__"):
            removed = apply_removals(wb, csv_path)
            if removed:
                changed = True
            applied.append((csv_path, removed, "removed"))
            continue

        # Audit / research Peer Bands deltas can use either the legacy 17-column
        # audit header or the canonical 15-column city-pass header. Map by header names
        # instead of requiring an exact-width workbook header.
        if csv_path.name.startswith("Audit_Peer_Bands__") or csv_path.name.startswith("Peer_Bands__"):
            header, data = load_csv(csv_path)
            canonical = [
                "Name","Country","City","Genre","Email","Social","Website","Source_URL",
                "Activity","Research_Date","Confidence","Contact_Type","Contact_Source",
                "Outreach_Readiness","Notes"
            ]
            legacy = [
                "Name","Country","City","Genre","Email","Social","Website","Links",
                "Source_URL","Activity","Research_Date","Status","Confidence",
                "Contact_Type","Contact_Source","Outreach_Readiness","Notes"
            ]
            header_norm = [norm(x) for x in header]
            if header_norm not in ([norm(x) for x in canonical], [norm(x) for x in legacy]):
                raise RuntimeError(f"Invalid Peer Bands delta header in {csv_path}: {header}")

            ws = wb["Peer Bands"]
            header_row = find_existing_header_row(ws)
            workbook_headers = [
                norm(ws.cell(header_row, c).value)
                for c in range(1, ws.max_column + 1)
            ]
            col_by_name = {name: i + 1 for i, name in enumerate(workbook_headers) if name}
            if any(norm(h) not in col_by_name for h in header):
                missing = [h for h in header if norm(h) not in col_by_name]
                raise RuntimeError(f"Peer Bands workbook missing columns for {csv_path}: {missing}")

            name_col = col_by_name["name"]
            row_by_name = {}
            for row_no in range(header_row + 1, ws.max_row + 1):
                key = norm(ws.cell(row_no, name_col).value)
                if key:
                    if key in row_by_name:
                        raise RuntimeError(
                            f"Duplicate Peer Band name in canonical DB: "
                            f"{ws.cell(row_no, name_col).value!r}"
                        )
                    row_by_name[key] = row_no

            updated = 0
            for raw in data:
                if not raw or not any(norm(x) for x in raw):
                    continue
                row = raw[:len(header)] + [""] * max(0, len(header) - len(raw))
                key = norm(row[0])
                if not key:
                    raise RuntimeError(f"Peer Bands row has empty Name: {raw}")
                row_no = row_by_name.get(key)
                if row_no is None:
                    ws.append([""] * ws.max_column)
                    row_no = ws.max_row
                    row_by_name[key] = row_no
                for csv_col, value in zip(header, row):
                    ws.cell(row_no, col_by_name[norm(csv_col)]).value = value
                updated += 1
                changed = True

            applied.append((csv_path, updated, "upserted"))
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
        for row in ws.iter_rows(
            min_row=start_data_row,
            max_row=ws.max_row,
            max_col=len(header),
            values_only=True,
        ):
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
        for path, _, action in applied:
            if action == "rejected-legacy-city-pass":
                continue
            shutil.move(str(path), str(APPLIED / path.name))

    print(f"pending_files={len(applied)} changed={changed}")
    for path, count, action in applied:
        print(f"{path.name}: {action}={count}")


if __name__ == "__main__":
    main()
