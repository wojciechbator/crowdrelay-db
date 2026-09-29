# SCOUT_PL daily persistence

SCOUT POLSKA — VIRYA writes one current-run text payload to scout_runs/current.json. The Render SCOUT_PL snapshot workflow converts that payload into database_festivals.xlsx.

database.xlsx remains the canonical workbook. The binary snapshot is generated only inside GitHub Actions, not by the automation connector.

Every SCOUT_PL run must replace current.json and therefore replace database_festivals.xlsx. It must not append yesterday's research.

Required snapshot sheets: RUN_INFO, OPPORTUNITIES, ORGANIZERS, CONTACTS, CANONICAL_APPEND. SUPPORT_TARGETS is optional.

Canonical changes continue through exact-schema CSV files in updates/pending/. The existing Apply database updates workflow owns all mutations of database.xlsx.

SCOUT automation must never upload or rebuild database.xlsx. It writes text deltas only.

Success is write + render + package validation + snapshot validation + Apply validation/readback. A commit alone is not success.

If a run has zero new or materially changed results, current.json must still be replaced with a valid zero-result snapshot so yesterday's data cannot masquerade as today's research.