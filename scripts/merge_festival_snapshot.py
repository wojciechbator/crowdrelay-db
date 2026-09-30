#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "database.xlsx"
PAYLOAD = Path(os.environ.get("SCOUT_PL_PAYLOAD", ROOT / "scout_runs" / "current.json"))

SHEET_MAP = {
    "RUN_INFO": "Festival Run Info",
    "OPPORTUNITIES": "Festival Opportunities",
    "ORGANIZERS": "Festival Organizers",
    "CONTACTS": "Festival Contacts",
    "CANONICAL_APPEND": "Festival Canonical Append",
    "SUPPORT_TARGETS": "Festival Support Targets",
}

REQUIRED = {"RUN_INFO", "OPPORTUNITIES", "ORGANIZERS", "CONTACTS", "CANONICAL_APPEND"}


def norm(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def canon_url(value: object) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    try:
        p = urlparse(raw if "://" in raw else "//" + raw)
        host = p.netloc.casefold().removeprefix("www.")
        path = re.sub(r"/+$", "", p.path or "").casefold()
        return host + path
    except Exception:
        return norm(raw)


def canonical_contact_keys(wb) -> set[tuple[str, ...]]:
    if "Contacts" not in wb.sheetnames:
        return set()
    ws = wb["Contacts"]
    header_row = None
    headers = {}
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = [norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if "email" in vals and "city" in vals:
            header_row = r
            headers = {v: i + 1 for i, v in enumerate(vals) if v}
            break
    if header_row is None:
        return set()

    out: set[tuple[str, ...]] = set()
    for r in range(header_row + 1, ws.max_row + 1):
        email = norm(ws.cell(r, headers.get("email", 1)).value)
        city = norm(ws.cell(r, headers.get("city", 1)).value)
        name = norm(ws.cell(r, headers.get("name", 1)).value)
        org = norm(ws.cell(r, headers.get("organization", 1)).value)
        if email and city:
            out.add(("email_city", email, city))
        if email:
            out.add(("email", email))
        if name and org:
            out.add(("name_org", name, org))
    return out


def row_identity(sheet: str, headers: list[str], row: list[object]) -> tuple[str, ...]:
    idx = {norm(h): i for i, h in enumerate(headers)}
    def get(name: str) -> str:
        i = idx.get(norm(name))
        return norm(row[i]) if i is not None and i < len(row) else ""
    def get_url(name: str) -> str:
        i = idx.get(norm(name))
        return canon_url(row[i]) if i is not None and i < len(row) else ""

    if sheet == "RUN_INFO":
        return ("run", get("Run_Date"))
    if sheet == "OPPORTUNITIES":
        source = get_url("Source_URL")
        if source:
            return ("url", source)
        return ("name_org", get("Name"), get("Organizer"))
    if sheet == "ORGANIZERS":
        website = get_url("Website")
        email = get("Email")
        if website:
            return ("url", website)
        if email:
            return ("email", email)
        return ("org_city", get("Organization"), get("City"))
    if sheet == "CONTACTS":
        email = get("Email")
        city = get("City")
        if email and city:
            return ("email_city", email, city)
        if email:
            return ("email", email)
        return ("name_org_city", get("Name"), get("Organization/Event"), city)
    if sheet == "CANONICAL_APPEND":
        target = get("Target_Sheet")
        email = get("Email")
        city = get("City")
        if email and city:
            return ("target_email_city", target, email, city)
        if email:
            return ("target_email", target, email)
        source = get_url("Source_URL")
        if source:
            return ("target_url", target, source)
        return ("target_name_city", target, get("Name"), city)
    if sheet == "SUPPORT_TARGETS":
        source = get_url("Public URL")
        if source:
            return ("url", source)
        return ("artist_region", get("Band / Artist"), get("Region"))
    return tuple(norm(x) for x in row)


def dedupe_rows(sheet: str, headers: list[str], rows: list[list[object]]) -> list[list[object]]:
    deduped: dict[tuple[str, ...], list[object]] = {}
    order: list[tuple[str, ...]] = []
    for raw in rows:
        row = list(raw[: len(headers)]) + [""] * max(0, len(headers) - len(raw))
        if not any(norm(x) for x in row):
            continue
        key = row_identity(sheet, headers, row)
        if not any(key[1:]) and sheet != "RUN_INFO":
            continue
        if key not in deduped:
            deduped[key] = row
            order.append(key)
            continue
        current = deduped[key]
        for i, value in enumerate(row):
            if norm(value):
                current[i] = value
    return [deduped[k] for k in order]


def normalize_canonical_actions(headers: list[str], rows: list[list[object]], canonical_contacts: set[tuple[str, ...]]) -> list[list[object]]:
    idx = {norm(h): i for i, h in enumerate(headers)}
    action_i = idx.get("action")
    target_i = idx.get("target_sheet")
    email_i = idx.get("email")
    city_i = idx.get("city")
    if None in (action_i, target_i, email_i, city_i):
        return rows

    out = []
    for row in rows:
        row = list(row)
        if norm(row[target_i]) == "contacts":
            email = norm(row[email_i])
            city = norm(row[city_i])
            exists = (email and city and ("email_city", email, city) in canonical_contacts) or (
                email and ("email", email) in canonical_contacts
            )
            if exists and norm(row[action_i]) == "new_pending":
                row[action_i] = "UPDATE_PENDING"
        out.append(row)
    return out


def replace_sheet(wb, title: str, headers: list[str], rows: list[list[object]]) -> None:
    if title in wb.sheetnames:
        old = wb[title]
        idx = wb.sheetnames.index(title)
        wb.remove(old)
        ws = wb.create_sheet(title, idx)
    else:
        ws = wb.create_sheet(title)

    ws.append(headers)
    for row in rows:
        ws.append(row)

    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in range(1, ws.max_column + 1):
        max_len = max((len(str(ws.cell(r, col).value or "")) for r in range(1, min(ws.max_row, 200) + 1)), default=8)
        ws.column_dimensions[get_column_letter(col)].width = min(max(max_len + 2, 10), 48)


def main() -> None:
    if not DB.exists():
        raise SystemExit("database.xlsx missing")
    payload = json.loads(PAYLOAD.read_text(encoding="utf-8"))
    if payload.get("snapshot_mode") != "current_run_only":
        raise SystemExit("snapshot_mode must be current_run_only")
    sheets = payload.get("sheets") or {}
    missing = REQUIRED - set(sheets)
    if missing:
        raise SystemExit(f"missing payload sheets: {sorted(missing)}")

    wb = load_workbook(DB)
    canonical_contacts = canonical_contact_keys(wb)
    stats = {}

    for source_name, target_name in SHEET_MAP.items():
        if source_name not in sheets:
            continue
        source = sheets[source_name]
        headers = list(source.get("headers") or [])
        rows = [list(r) for r in (source.get("rows") or [])]
        rows = dedupe_rows(source_name, headers, rows)
        if source_name == "CANONICAL_APPEND":
            rows = normalize_canonical_actions(headers, rows, canonical_contacts)
        replace_sheet(wb, target_name, headers, rows)
        stats[target_name] = len(rows)

    tmp = DB.with_suffix(".tmp.xlsx")
    wb.save(tmp)
    os.replace(tmp, DB)
    print("FESTIVAL_MERGE_OK", stats)


if __name__ == "__main__":
    main()
