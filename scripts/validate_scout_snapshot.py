#!/usr/bin/env python3
from __future__ import annotations

from datetime import date
from pathlib import Path
from zipfile import ZipFile
import posixpath
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SNAPSHOT = ROOT / "database_festivals.xlsx"

NS_MAIN = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

REQUIRED_HEADERS = {
    "RUN_INFO": [
        "Run_Date", "Scope", "New_Organizer_Targets",
        "New_or_Updated_Opportunities", "Current_Deadline",
        "Outreach_This_Run", "MASTER_Writeback", "SCOUT_Writeback",
        "Continuation",
    ],
    "OPPORTUNITIES": [
        "Name", "Organizer", "Current_Status", "Source_URL", "Evidence",
    ],
    "ORGANIZERS": [
        "Organization", "City", "Website", "Email", "Verification",
    ],
    "CONTACTS": [
        "Email", "Name", "Organization/Event", "City",
        "Kind", "Source", "History/Notes",
    ],
    "CANONICAL_APPEND": [
        "Target_Sheet", "Action", "Email", "Name", "Organization",
        "City", "Kind", "Source_URL", "Reason",
    ],
}

OPTIONAL_SHEETS = {"SUPPORT_TARGETS"}
ALLOWED_CANONICAL_SHEETS = {
    "Venues", "Peer Bands", "Beacons", "Booking Agents", "Contacts",
}
ALLOWED_ACTIONS = {"NEW_PENDING", "UPDATE_PENDING"}


def fail(message: str) -> None:
    raise SystemExit(f"SCOUT_SNAPSHOT_QA_FAIL: {message}")


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(".//x:t", NS_MAIN))
    v = cell.find("x:v", NS_MAIN)
    raw = "" if v is None or v.text is None else v.text
    if cell_type == "s" and raw:
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError):
            fail(f"invalid shared-string index: {raw}")
    return raw


def col_number(ref: str) -> int:
    m = re.match(r"([A-Z]+)", ref.upper())
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n


def zip_read(path: Path, member: str) -> bytes:
    try:
        with ZipFile(path) as z:
            return z.read(member)
    except Exception as exc:
        fail(f"cannot read XLSX member {member!r}: {exc}")


def read_workbook(path: Path) -> dict[str, list[list[str]]]:
    if not path.exists():
        fail("database_festivals.xlsx is missing")
    if path.stat().st_size == 0:
        fail("database_festivals.xlsx is empty")

    try:
        with ZipFile(path) as z:
            names = set(z.namelist())
    except Exception as exc:
        fail(f"not a readable XLSX/ZIP: {exc}")

    required_files = {"xl/workbook.xml", "xl/_rels/workbook.xml.rels"}
    if not required_files.issubset(names):
        fail("missing workbook XML parts")

    shared_strings: list[str] = []
    if "xl/sharedStrings.xml" in names:
        root = ET.fromstring(zip_read(path, "xl/sharedStrings.xml"))
        for si in root.findall("x:si", NS_MAIN):
            shared_strings.append(
                "".join(t.text or "" for t in si.findall(".//x:t", NS_MAIN))
            )

    wb_root = ET.fromstring(zip_read(path, "xl/workbook.xml"))
    rel_root = ET.fromstring(zip_read(path, "xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rel_root.findall("r:Relationship", NS_REL)
        }

        sheets: dict[str, list[list[str]]] = {}
        for sheet in wb_root.findall("x:sheets/x:sheet", NS_MAIN):
            name = sheet.attrib["name"]
            rid = sheet.attrib[f"{{{REL_NS}}}id"]
            target = rel_targets.get(rid)
            if not target:
                fail(f"sheet {name!r} has no relationship target")
            if target.startswith("/"):
                sheet_path = target.lstrip("/")
            else:
                sheet_path = posixpath.normpath(posixpath.join("xl", target))
            if sheet_path not in names:
                fail(f"sheet {name!r} target missing: {sheet_path}")

            root = ET.fromstring(zip_read(path, sheet_path))
            rows: list[list[str]] = []
            max_col = 0
            parsed = []
            for row in root.findall(".//x:sheetData/x:row", NS_MAIN):
                cells = {}
                for c in row.findall("x:c", NS_MAIN):
                    idx = col_number(c.attrib.get("r", ""))
                    if idx:
                        max_col = max(max_col, idx)
                        cells[idx] = cell_value(c, shared_strings).strip()
                parsed.append(cells)

            for cells in parsed:
                rows.append([cells.get(i, "") for i in range(1, max_col + 1)])

            sheets[name] = rows

    return sheets


def header_map(rows: list[list[str]], sheet: str) -> tuple[list[str], dict[str, int]]:
    if not rows:
        fail(f"{sheet}: sheet is empty")
    header = rows[0]
    normalized = [re.sub(r"\s+", " ", x.strip()).casefold() for x in header]
    index = {x: i for i, x in enumerate(normalized) if x}
    expected = [re.sub(r"\s+", " ", x.strip()).casefold() for x in REQUIRED_HEADERS[sheet]]
    missing = [h for h in expected if h not in index]
    if missing:
        fail(f"{sheet}: missing required columns {missing}")
    return header, index


def value(row: list[str], index: dict[str, int], column: str) -> str:
    i = index.get(column.casefold())
    return row[i].strip() if i is not None and i < len(row) else ""


def nonempty_rows(rows: list[list[str]]) -> list[list[str]]:
    return [r for r in rows[1:] if any(x.strip() for x in r)]


def dedupe_check(sheet: str, rows: list[list[str]], index: dict[str, int], key_columns: list[str]) -> None:
    seen = set()
    for n, row in enumerate(nonempty_rows(rows), start=2):
        key = tuple(value(row, index, c).casefold() for c in key_columns)
        if not any(key):
            fail(f"{sheet}: row {n} has no identity fields")
        if key in seen:
            fail(f"{sheet}: duplicate identity at row {n}: {key}")
        seen.add(key)


def url_check(sheet: str, rows: list[list[str]], index: dict[str, int], column: str) -> None:
    if column.casefold() not in index:
        return
    for n, row in enumerate(nonempty_rows(rows), start=2):
        url = value(row, index, column)
        if url and not re.match(r"^https?://", url, re.I):
            fail(f"{sheet}: invalid {column} at row {n}: {url!r}")


def main() -> None:
    sheets = read_workbook(SNAPSHOT)
    names = set(sheets)
    required = set(REQUIRED_HEADERS)
    missing_sheets = required - names
    if missing_sheets:
        fail(f"missing required sheets: {sorted(missing_sheets)}")

    unexpected = names - required - OPTIONAL_SHEETS
    if unexpected:
        fail(f"unexpected sheets: {sorted(unexpected)}")

    for sheet in REQUIRED_HEADERS:
        rows = sheets[sheet]
        _, index = header_map(rows, sheet)
        if sheet == "RUN_INFO":
            data = nonempty_rows(rows)
            if len(data) != 1:
                fail(f"RUN_INFO must contain exactly one data row, got {len(data)}")
            run_date = value(data[0], index, "Run_Date")
            try:
                date.fromisoformat(run_date)
            except ValueError:
                fail(f"RUN_INFO Run_Date is not ISO date: {run_date!r}")
        elif sheet == "OPPORTUNITIES":
            dedupe_check(sheet, rows, index, ["Name", "City"])
            url_check(sheet, rows, index, "Source_URL")
            for n, row in enumerate(nonempty_rows(rows), start=2):
                if not value(row, index, "Name"):
                    fail(f"{sheet}: row {n} missing Name")
                if not value(row, index, "Evidence"):
                    fail(f"{sheet}: row {n} missing Evidence")
        elif sheet == "ORGANIZERS":
            dedupe_check(sheet, rows, index, ["Organization", "City"])
            url_check(sheet, rows, index, "Website")
            for n, row in enumerate(nonempty_rows(rows), start=2):
                if not value(row, index, "Organization"):
                    fail(f"{sheet}: row {n} missing Organization")
        elif sheet == "CONTACTS":
            data = nonempty_rows(rows)
            seen = set()
            for n, row in enumerate(data, start=2):
                email = value(row, index, "Email")
                city = value(row, index, "City")
                name = value(row, index, "Name")
                identity = (email.casefold(), city.casefold(), name.casefold())
                if not any(identity):
                    fail(f"{sheet}: row {n} has no identity")
                if identity in seen:
                    fail(f"{sheet}: duplicate identity at row {n}: {identity}")
                seen.add(identity)
            url_check(sheet, rows, index, "Source")
        elif sheet == "CANONICAL_APPEND":
            dedupe_check(sheet, rows, index, ["Target_Sheet", "Email", "Name", "City"])
            url_check(sheet, rows, index, "Source_URL")
            for n, row in enumerate(nonempty_rows(rows), start=2):
                target = value(row, index, "Target_Sheet")
                action = value(row, index, "Action")
                if target not in ALLOWED_CANONICAL_SHEETS:
                    fail(f"{sheet}: row {n} invalid Target_Sheet {target!r}")
                if action not in ALLOWED_ACTIONS:
                    fail(f"{sheet}: row {n} invalid Action {action!r}")
                if not value(row, index, "Reason"):
                    fail(f"{sheet}: row {n} missing Reason")
                if not value(row, index, "Source_URL"):
                    fail(f"{sheet}: row {n} missing Source_URL")

    print(
        "SCOUT_SNAPSHOT_QA_OK",
        {name: len(nonempty_rows(rows)) for name, rows in sorted(sheets.items())},
    )


if __name__ == "__main__":
    main()
