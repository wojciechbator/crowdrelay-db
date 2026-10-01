# SCOUT daily persistence

The canonical GitHub workbook is `database.xlsx`.

`scout_runs/current.json` is a **current-run delta**, not the whole database. The merge workflow upserts that delta into the persistent `Festival ...` sheets in `database.xlsx`; it must never erase older still-relevant festival/support rows merely because they were absent from today's run.

The shared producer may publish `SCOUT_PL`, `SCOUT_EUROPE`, or `SCOUT_SHARED` data, but it must be built from the live canonical Google SCOUT after Drive readback.

`database_festivals.xlsx` is retired and must not be recreated.

Success means: fresh payload readback + delta upsert + logical dedupe + payload-landed validation + committed `database.xlsx`.
