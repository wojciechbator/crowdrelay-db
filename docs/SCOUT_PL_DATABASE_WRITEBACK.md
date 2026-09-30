# SCOUT_PL daily persistence

SCOUT POLSKA — VIRYA writes one current-run text payload to `scout_runs/current.json`.

There is now **one workbook only**: `database.xlsx`.

The SCOUT merge workflow reads `scout_runs/current.json` and refreshes these operational sheets inside the canonical workbook:

- Festival Run Info
- Festival Opportunities
- Festival Organizers
- Festival Contacts
- Festival Canonical Append
- Festival Support Targets (optional)

`database_festivals.xlsx` is retired and must not be recreated.

Before writing the festival sheets, the merge pass also deduplicates the canonical sheets using the same logical keys as the database importer:

- Venues: Name + City
- Peer Bands: Name + City
- Beacons: Kind + City + canonical Destination_URL
- Booking Agents: Name + Agency, fallback Name
- Contacts: Email + City, fallback Email, fallback Name + Organization

Festival sheets are independently deduplicated using stable URL/email/name keys. Canonical append rows for Contacts are normalized from `NEW_PENDING` to `UPDATE_PENDING` when that contact already exists in the canonical Contacts sheet.

The Apply database updates workflow remains the owner of CSV mutations under `updates/pending/`.

Success means: merge + logical dedupe + workbook validation + commit of `database.xlsx`.
