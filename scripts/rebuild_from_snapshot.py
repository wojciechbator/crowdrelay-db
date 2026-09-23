#!/usr/bin/env python3
from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
SNAPSHOT = ROOT / "snapshot"
DB = ROOT / "database.xlsx"

EXPECTED = {
    "Venues": [
        "Name","City","Country","Address","Website","Audience_Genre","Capacity",
        "Booking_Contact","Email","Public_Financial_Info","Status","Source_URL",
        "Research_Date","Target_Fit","Contact_Quality","Outreach_Angle","Notes"
    ],
    "Peer_Bands": [
        "Name","Country","City","Genre","Email","Social","Website","Links","Source_URL",
        "Activity","Research_Date","Status","Confidence","Contact_Type",
        "Contact_Source","Outreach_Readiness","Notes"
    ],
    "Beacons": [
        "Name","Kind","City","Email","Destination_URL","Source_URL","Active","Verified",
        "Accepts_Outreach","Do_Not_Contact","Relationship_Score","Relevance_Pct","Confidence_Pct"
    ],
    "Booking_Agents": [
        "Name","Agency","Email","Roster_URL","Genres","Source_URL","Research_Date","Notes"
    ],
    "Contacts": [
        "Email","Name","Organization","City","Kind","Phone","Notes","Last_Seen","Disappeared"
    ],
}

FILES = {
    "Venues": SNAPSHOT / "Venues.raw.csv",
    "Peer Bands": SNAPSHOT / "Peer_Bands.raw.csv",
    "Beacons": SNAPSHOT / "Beacons.raw.csv",
    "Booking Agents": SNAPSHOT / "Booking_Agents.raw.csv",
    "Contacts": SNAPSHOT / "Contacts.raw.csv",
}

def norm(v: object) -> str:
    return "" if v is None else str(v).strip().casefold()

def clean_row(row: list[str], width: int) -> list[str]:
    row = list(row)
    if len(row) == width + 1 and row[0].strip().isdigit():
        row = row[1:]
    if len(row) < width:
        row += [""] * (width - len(row))
    return row[:width]

def read_snapshot(path: Path, expected: list[str]) -> tuple[list[str], list[list[str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))

    header_idx = None
    for i, row in enumerate(rows):
        cleaned = clean_row(row, len(expected))
        if [norm(x) for x in cleaned] == [norm(x) for x in expected]:
            header_idx = i
            break

    if header_idx is None:
        raise RuntimeError(f"Expected header not found in {path}")

    out: list[list[str]] = []
    seen = set()
    for row in rows[header_idx + 1:]:
        cleaned = clean_row(row, len(expected))
        if not any(norm(x) for x in cleaned):
            continue
        sig = tuple(norm(x) for x in cleaned)
        if sig in seen:
            continue
        seen.add(sig)
        out.append(cleaned)
    return expected, out

def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)

    for sheet, path in FILES.items():
        expected = EXPECTED[
            "Peer_Bands" if sheet == "Peer Bands"
            else "Booking_Agents" if sheet == "Booking Agents"
            else sheet
        ]
        if not path.exists():
            raise RuntimeError(f"Missing snapshot file: {path}")
        header, rows = read_snapshot(path, expected)
        ws = wb.create_sheet(sheet)
        ws.append(header)
        for row in rows:
            ws.append(row)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        print(f"{sheet}: {len(rows)} rows")

    wb.save(DB)
    print(f"Wrote {DB} ({DB.stat().st_size} bytes)")

if __name__ == "__main__":
    main()
