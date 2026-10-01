# Shared SCOUT → GitHub persistence

Both VIRYA scout jobs publish **current-run deltas** from the live canonical Google SCOUT to `scout_runs/current.json`:

- `SCOUT_PL` — Poland discovery
- `SCOUT_EUROPE` — Europe discovery
- `SCOUT_SHARED` — explicit repair/recovery runs

The canonical GitHub workbook is **`database.xlsx`**. `database_festivals.xlsx` is retired and must not be recreated.

## Persistence semantics

`scout_runs/current.json` is a delta, not a full snapshot. The merge workflow UPSERTS that delta into persistent operational sheets:

- Festival Run Info
- Festival Opportunities
- Festival Organizers
- Festival Contacts
- Festival Canonical Append
- Festival Support Targets (optional)

Rows already present in `database.xlsx` are preserved when absent from the current-run delta. Incoming non-empty fields update the matched persistent row. This prevents a Poland or Europe run from deleting state written by the other scout.

## Dedupe / identity

Festival sheets use stable identities appropriate to their role:
- Run Info: Run_Date + Scope
- Opportunities: Name + Organizer, with Source_URL fallback
- Organizers: Website, then Email, then Organization + City
- Contacts: Email + City, then Email, then Name + Organization/Event + City
- Support Targets: Dedupe Key, then Band/Artist + Region, then Public URL
- Canonical Append: Target_Sheet + contact/url identity

Canonical append Contacts are normalized from `NEW_PENDING` to `UPDATE_PENDING` when the canonical Contacts sheet already contains the contact.

## Acceptance gate

A GitHub sync is PASS only when:
1. `current.json` is written using a fresh blob SHA and read back.
2. GitHub Actions **Merge current Scout delta** succeeds.
3. **Validate merged workbook** succeeds.
4. **Commit unified database workbook** succeeds.
5. The validator confirms every non-empty field from the current payload is present in the merged workbook.
6. Existing persistent festival state is not erased merely because it was absent from the delta.

The `crowdrelay-database-apply` concurrency group serializes Scout merges with the canonical database update workflow.

The Apply database updates workflow remains the owner of CSV mutations under `updates/pending/`.
