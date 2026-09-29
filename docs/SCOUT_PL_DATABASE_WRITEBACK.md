# SCOUT_PL → CrowdRelay DB writeback contract

database.xlsx remains the canonical workbook. database_festivals.xlsx is a replace-on-each-run snapshot owned exclusively by the daily SCOUT POLSKA — VIRYA automation.

## Per-run contract

1. Research starts from fresh VIRYA_MASTER, fresh SCOUT, latest RUN_INFO, Gmail history, then current web research.
2. Only new or materially changed, source-verified results from the current run enter database_festivals.xlsx.
3. The snapshot is replaced on every run; it is not a historical append-only workbook.
4. Required snapshot sheets are RUN_INFO, OPPORTUNITIES, ORGANIZERS, CONTACTS, and CANONICAL_APPEND. SUPPORT_TARGETS is optional.
5. CANONICAL_APPEND documents the exact records queued to the canonical database. It must never claim an applied record that is not present in updates/pending/.
6. Canonical data enters through exact-schema CSV deltas in updates/pending/. The existing Apply database updates workflow owns the XLSX mutation.
7. Organizer/event/opportunity rows are not forced into canonical sheets. Only records that genuinely match Venues, Peer Bands, Beacons, Booking Agents, or Contacts are queued.
8. The writer must re-fetch main immediately before commit and handle races without overwriting newer changes.
9. After push, the writer must read back database_festivals.xlsx, inspect the snapshot-validation workflow, inspect the triggered Apply run, and read back database.xlsx after Apply.
10. A green commit is not success by itself. Success means write + validation + Apply + readback were confirmed.
11. A zero-result run still replaces the snapshot with a valid current-run workbook containing zero result rows rather than retaining yesterday's data.

The validator is intentionally read-only and uses only the XLSX package structure. It checks workbook integrity, required sheets/headers, identities/deduplication, source URLs, and canonical-writeback actions.
