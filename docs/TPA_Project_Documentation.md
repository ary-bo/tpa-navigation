# TPA Navigation — Technical Reconstruction & Continuation Specification

**Document version:** 4.5  
**Date:** 14 November 2026  
**Repository:** https://github.com/ary-bo/tpa-navigation  
**Target branch:** `main`  
**Supersedes:** Previous contents of this file and the V3 technical handover  
**Companion document:** `TPA_Navigation_Technical_Reconstruction_Spec_V4_1_English.docx` (project handover copy)

> **Authority and evidence rule:** Current source code is authoritative for implemented behavior; the approved `TPA Final Form Spec V06 Revised.docx` is authoritative for approved auditor-form requirements; this document records project context, engineering constraints, verified change history, and handover policy. Clearly distinguish implemented, approved, sample-derived, user-confirmed, and unverified information. Never turn a sample workbook or inference into a universal business rule.

## 1. Purpose and executive summary

TPA Navigation is a browser-based application for customer/store mapping, Journey Plan management, sales-data ingestion, auditor visits and forms, reporting, settings, and Google Drive integration. The project is an incremental evolution of an existing application. Preserve working architecture and user data; implement bounded changes; test them; and document the exact resulting state.

The primary application entry point is `index.html`. The application uses IndexedDB and LocalStorage for local persistence, Leaflet/OpenStreetMap for maps, SheetJS for Excel processing, and Google Drive APIs for cloud synchronization and backup operations. Local application data may remain available offline, but Drive operations, OAuth, and fresh map tiles require connectivity.

### 1.1 Product areas

| Area | Responsibility |
|---|---|
| Map and stores | Customer map, store search/filtering, store details and location context |
| Journey Plan | Weekly visit planning, ordering, editing, import/export and pending approval |
| Data import | Journey Plan, sales workbook and optional Product Mapping ingestion |
| Auditor form | Visit questions, conditional visibility, GPS/manual map pin, drafts and saved records |
| Reports | Sales and operational summaries and data-quality reporting |
| Settings | Language, theme, Drive configuration, PIN and data maintenance |
| Cloud integration | Google Drive sync and multiple distinct backup/restore workflows |

## 2. Repository and verified state

- Repository: `ary-bo/tpa-navigation`
- Primary branch for changes: `main`
- The repository tree inspected for this handover contained: `CHANGELOG_FA.txt`, `README.md`, `docs/TPA_Project_Documentation.md`, `index.html`, `manifest.webmanifest`, and `sw.js`.
- Inspected main commit for the backup-deletion change: `a88574667f9ef1299e62526ca1eb035df19b94cb`.
- `index.html` pre-change blob: `db2570fd2c863b6ec3baff52a7db1b32ff23550d`.
- `index.html` post-change blob: `23867219b32ce5097adbae1cc73d603c1605ccb7`.
- Pre-change safety branch recorded: `backup/main-before-backup-delete`.

These identifiers describe the inspected change, not necessarily the latest repository state at the time a future engineer reads this file. Re-fetch main and record its current HEAD and blob SHAs before every change.

### 2.1 Repository safety procedure

1. Fetch current `main`, HEAD, and exact blobs of files to be changed.
2. Inspect the actual source and any commits added since this document.
3. Create a purpose-named backup branch before editing.
4. Make one coherent, bounded change set.
5. Run relevant syntax, automated, browser and data-compatibility tests.
6. Fetch the committed files again and inspect the resulting content.
7. Record commit SHA, changed paths, test results, deployment status and rollback reference.
8. Never claim deployment or live verification merely because a commit exists.

## 2.2 Change record — 14 November 2026

- Change: moved Visit outcome ahead of location correction in the auditor form.
- Conditional behavior: location correction and its map sub-question are visible only for the first three visit outcomes; selecting **مشتری پیدا نشد** hides the location section and the existing form update path clears its active values.
- Changed application file: `index.html`.
- Main commit: `179b54b07a704a76cfdcfa29989bff9e6843bb7b`.
- Pre-change safety branch: `backup/main-before-visit-result-location-20261114`.
- Pre-change `index.html` blob: `23867219b32ce5097adbae1cc73d603c1605ccb7`.
- Post-change `index.html` blob: `6ed50ae0290e43ea1d0dd27d19011138dcc4082b`.
- Verification: source-level checks only; browser/device and data-flow tests remain outstanding.
- Deployment: commit recorded on `main`; live deployment not verified.

## 2.3 Change record — 14 November 2026

- Change: added a required oil-availability/visibility gate as the first question in Section 4, immediately before existing Q14.
- New question: **آیا محصولات روغن در فروشگاه موجود و قابل مشاهده است ؟** with **بله / خیر** choices; internal ID `q14_gate`, display number `۱۳A`.
- Existing Q14, Q15, Q16 and Q17 are shown only when `q14_gate` is exactly **بله**. The Section 4 heading remains governed by the existing observation-block rule.
- When the answer is **خیر**, the existing `update()` visibility cleanup removes active answers for Q14–Q17; serialized non-optional hidden questions are stored as `N/A`. Existing stored records/drafts are not bulk migrated.
- English label added for the new question. Existing Q14–Q17 identifiers, labels, options and validation rules are preserved.
- Changed application file: `index.html`; code commit `f0c6fe08f5d4b9813ad7e077f765847337f17285`; post-change blob `6f134e747d96b09828b57d13065e524b65e01e03`.
- Safety branch: `backup/main-before-section4-oil-availability-gate-20261114`; base commit `505217f2dd882c4190872d48e0941c498b1226bc`.
- Verification: committed source re-fetched; exact question, answer choices, and Q14–Q17 conditions confirmed. Browser/mobile interaction, saved-draft/export regression tests and deployment/live verification remain outstanding.

## 3. Architecture and file responsibilities

| Component | Observed role | Guidance |
|---|---|---|
| `index.html` | Main application UI, styles and most core JavaScript | Re-inspect before each edit; avoid unnecessary rewrites |
| IndexedDB | Structured local persistence; historical database name `tpa-nav`, store `kv` | Preserve keys and values; use backward-compatible migrations |
| LocalStorage | Settings, PIN, saved filters, backup metadata and per-customer drafts | Avoid broad clearing except for explicitly requested actions |
| Leaflet / OpenStreetMap | Interactive maps and map tiles | Preserve manual pin workflow if GPS is unavailable |
| SheetJS | Excel parsing/export | Preserve header handling and formula-injection protection |
| Google Drive API | Sync and cloud backups | Scope every operation to the intended file family |
| `manifest.webmanifest`, `sw.js` | PWA metadata and service worker | Inspect cache/version behavior before release changes |

Historical inspection recorded IndexedDB keys `raw`, `plan`, and `forms`; LocalStorage items included `settings`, `pin`, `savedFilters`, `lastBackup`, and per-customer draft keys `fd_<New Customer Code>`. Helper names previously observed include `open`, `dbGet`, `dbSet`, `dbDel`, `lsGet`, and `lsSet`. Verify all names and exact semantics in live source before relying on them.

## 4. Source workbooks and data ingestion

Project sample files include:

- `JP_V06_14050811.xlsx` — Journey Plan sample.
- `TPA_Alborz_14050816.xlsx` — sales workbook sample.
- `Product Mapping.xlsx` — Item Code to Product Name sample mapping.
- `Targets.xlsx` — sample Volume Target and Coverage Target worksheets.

Samples are fixtures, not proof of complete business rules or master data.

### 4.1 Historical parser observations

- Header discovery (`findHeader`) was recorded as scanning the first 15 rows by default; header matching trims whitespace and duplicate names retain the first index.
- Journey Plan parser (`parseJP`) historically required `CustomerI`, `Latitude`, `Longitude`, and `Salesman Name`. Other possible fields include New Customer Code, Oracle Code, Customer Name, Address, Salesman Code, TerritoryNo, Area Name, Route Name/DescriptionE, Outlet Class and Customer Class.
- GPS was treated as valid only if both coordinates parse, are within legal ranges, and both absolute values exceed 0.0001; otherwise coordinates become null.
- Product mapping parser (`parseMapping`) searches for `Item Code` and `Product Name`; mapping is optional.
- Sales parser (`parseSales`) historically required Persian headers `کد مشتری` and `شماره سفارش`. Target delivery date may have multiple source header spellings.
- Historical customer matching precedence: sales customer code to Journey Plan CustomerI; then Account Number to Oracle Code; then final segment of sales customer code to New Customer Code.
- Historical normalization includes Persian/Arabic digit conversion, whitespace cleanup, code normalization, selected Arabic/Persian character normalization and supported date conversion.

**Mandatory caution:** Re-check live parser code, especially status exclusions (including `isExcluded()`), date handling, joins, and fallback behavior. Do not infer global rules from one workbook.

## 5. Journey Plan and reporting

Historical Journey Plan records included: `pid`, `visitDate` (YYYY-MM-DD), `customerId`, `storeName`, `customerName`, `address`, `route`, `lat`, `lng`, `priority` (High/Medium/Low), `reason`, `comment`, `source` (Internal/External), and `order`. Visit state values recorded were `empty`, `done`, and `skip`. Revalidate current schemas.

External plan import historically required Visit Date and Customer ID; blank priority defaulted to Medium. Possible row errors included `EMPTY_ID`, `BAD_DATE`, `UNKNOWN_ID`, `BAD_PRIORITY`, and `DUPLICATE`. Valid external imports enter a pending approval flow. Preserve current duplicate detection, official export columns, sheet name, instruction sheet, and formula-injection protection.

`Targets.xlsx` was not established as part of the inspected application data flow. Treat target-versus-actual reporting as a separate new capability. Before implementation, approve metric definitions, unit, aggregation grain, join key, date/calendar rules, duplicate/missing-target handling, dimensions, actuals source and reconciliation rules.

## 6. Approved Auditor Form V06 requirements

The following items were carried forward from the prior handover's summary of the approved V06 specification. They are requirements, not proof of current implementation. Retrieve and inspect the original approved DOCX before changing exact question text, options or competitor lists.

- Remove the requirement for a customer's previous purchase history from TPA or the visitor.
- Q11 and Q12 must be displayed for all customers and must not depend on purchase history.
- Q11A appears only when Q11 is answered “Yes, I have had a problem.”
- Q13 appears only when Q12 indicates that some or all products were returned.
- Q15 and its two options remain unchanged as specified in V06.
- Q16 and Q17 record Facing counts only; do not calculate or display shelf-share percentages.
- Preserve competitor brand lists for product groups exactly as specified in V06.
- Remove the summary and follow-up section and Q24 completely; do not add a replacement question.
- An “Other” text field appears only while its corresponding “Other” option is selected.
- If automatic GPS is unavailable, the map must still open and allow manual pin placement.
- No free-text address field; “Get Location” confirms and saves selected latitude and longitude.
- The approved form ends at Q23.

### 6.1 Decisions requiring explicit approval

| Topic | Known ambiguity / historical behavior | Required action |
|---|---|---|
| Form eligibility | Historical guard required Journey Plan membership and New Customer Code | Obtain product-owner decision before changing |
| Q16/Q17 limits | Historical code capped Q16 at 500 and required Q17 <= Q16; V06 summary does not establish these limits | Confirm whether limits remain |
| Legacy records/drafts | Existing identifiers and serialization may differ from V06 | Design versioned, backward-compatible migration |
| Exact question text and competitor lists | Controlled by original approved V06 | Read source DOCX; do not reconstruct from memory |

### 6.2 Conditional visibility test matrix

| Control | Expected display rule |
|---|---|
| Q11 and Q12 | Always visible for eligible customers; independent of purchase history |
| Q11A | Only when Q11 = “Yes, I have had a problem” |
| Q13 | Only when Q12 indicates some/all products returned |
| Other text | Only while corresponding Other option is selected |
| Q24 | Must not exist; no replacement follow-up section |

### 6.3 Visit outcome and location-correction dependency (implemented 14 November 2026)

The first questions in Section 1 now appear in this order:

1. Q3 — **نتیجه مراجعه** (Visit outcome), retaining the four existing choices:
   - شخص مسئول حاضر بود
   - شخص مسئول حاضر نبود
   - فروشگاه تعطیل بود
   - مشتری پیدا نشد
2. Q4 — **آیا نیاز به اصلاح موقعیت مکانی وجود دارد؟** (Location correction required?), followed by its existing map-coordinate question when the answer is Yes.

Visibility rule: Q4 is displayed only when the selected visit outcome is one of the first three choices. If the fourth choice, **مشتری پیدا نشد**, is selected, Q4 and its dependent map question are hidden.

Data safety: the existing form `update()` routine clears the value and auxiliary state of any question that becomes hidden, then re-renders its control. Therefore, changing an earlier first-three outcome to **مشتری پیدا نشد** clears any previous location-correction answer and map coordinates from the active form state/draft. On record creation, hidden non-optional questions are represented as `N/A`; the visit outcome remains recorded. Existing saved records are not bulk-migrated or rewritten by this UI change.

Implementation locations: `index.html`, question definitions in the `Q` array (existing IDs `q4` and `q3` are retained). Visit outcome is ordered before location correction; its displayed number is 3 and the location question is numbered 4. The location question's `show` predicate excludes the exact fourth visit-outcome option.

Verification status:
- Source-level inspection confirms the question order, conditional predicate, and existing hidden-question cleanup path.
- Not yet browser-tested: all four outcome selections, switching from a first-three option to **مشتری پیدا نشد**, draft reload after that switch, and saved-record export/report behavior.
- No claim is made that the updated source has been deployed to the live site.

### 6.4 Distribution-problem options: remove profit-margin/agreement issue (implemented 14 November 2026)

In auditor question Q11A, **نوع مشکل در توزیع** (Distribution problem type), the option **حاشیه سود یا توافق** (Profit margin or agreement issue) has been removed at the user's request. The remaining options are late delivery, incomplete order, wrong products delivered, damaged products, price discrepancy, and Other (please explain).

Implementation: the option was removed from the Q11A `opts` array in `index.html`. The question identifier (`q11a`), trigger condition (Q11 = the affirmative problem response), the Other option, and the surrounding question flow remain unchanged. The English/Persian translation dictionary entry was retained to avoid unnecessarily breaking display of historical stored responses that may contain the old option.

Compatibility and data safety: this change affects new/edited form selections only. No migration or bulk rewrite of existing saved audit records or drafts was performed. Existing historical values are not intentionally deleted by this change; verify their display/export behavior if legacy records containing the removed option are edited or exported.

Verification status:
- Source edit was constrained to exactly one matching option sequence in Q11A; the source update operation confirmed one occurrence was removed.
- A post-commit fetch and syntax/browser regression checks are still required before calling the change fully verified.
- Deployment/live behavior has not been independently verified.

Git traceability:
- Code commit: `ef0581e12e0b7bb7f0cc1c08cde53fa9a299e610`
- Updated `index.html` blob: `18f6f2c3b31eac268524b8b018b49d5519331c86`
- Rollback branch: `backup/main-before-remove-profit-margin-option-20261114`

## 7. Audit records and data compatibility

Historical form storage used IndexedDB key `forms`; per-customer drafts used LocalStorage keys `fd_<New Customer Code>`. Historical record serialization included question fields, GPS fields `q3_lat` and `q3_lng`, JSON snapshots `_a` and `_o`, and identity/timestamp fields such as `id`, `timestamp`, `updated`, `storeId`, `visitor`, `q2_code`, and `q2_name`. Multi-select arrays were historically joined with `|`; some “Other” text was appended in bracketed form. Re-check live code before using these details as current contracts.

**Migration rule:** Never clear or silently rewrite existing forms or drafts. Define a versioned migration, retain a pre-migration backup, test old records and drafts, and provide rollback.

## 8. Backup, restore and deletion — distinct backup families

Do not conflate the following pathways:

| Backup family | Historical filename | Scope |
|---|---|---|
| Plan backup | `tpa_plan_backup.json` | Plan and pending data, filters and saved filters |
| Audit-form backup | `tpa_audit_forms_backup.json` | Saved audit forms and per-customer drafts |
| Full-system backup | `tpa_full_system_backup_v1_<timestamp>.json`; legacy `tpa_full_system_backup_v1.json` | Primary IndexedDB store entries plus LocalStorage values, counts, version, timestamp and checksum |

### 8.1 Full-system payload and restore

Historical payload fields: `app = tpa-navigation`, `type = tpa-full-system-backup`, `version = 1`, `createdAt`, `database { name, store, entries }`, `localStorage`, `counts { databaseEntries, localStorageKeys }`, and `checksum`.

The full-system list is scoped to Google Drive `appDataFolder`, filters the versioned and legacy full-system filename patterns, sorts newest first and paginates. Restore validates app/type/version, database name/store, entries, LocalStorage shape, checksum and timestamp before replacing data. Verify the exact current rollback behavior in source before modifying restore.

### 8.2 Confirmed V4 backup deletion feature

The main-branch change recorded in commit `a88574667f9ef1299e62526ca1eb035df19b94cb` added a per-item Delete Backup action to the full-system backup list while preserving Restore.

The inspected implementation:
1. Requires a selected backup object with a Drive file ID.
2. Shows a confirmation dialog identifying the selected file by date/name and warning that deletion is permanent.
3. Sends an HTTP DELETE request to Google Drive for that selected file ID only.
4. Reports success/failure and refreshes the list after successful deletion.
5. Does not clear current IndexedDB, LocalStorage, plan data, audit data or other backup files.

Code inspection verified the deletion function, per-item control, Drive DELETE request, preserved restore handler, list refresh, and preserved backup-creation function. The user explicitly confirmed in the target environment: “درسته، کار میکنه” (“It works”). This is a user-reported manual acceptance result, not a substitute for broader automated testing.

### 8.3 Operational deletion workflow

1. Open Settings → Backup & Restore.
2. Load the full-system backup list using Restore Backup.
3. Identify the intended version by timestamp, filename and size.
4. Select Delete Backup on that item.
5. Read the confirmation and confirm only if the selected backup is correct.
6. Verify the item disappears after list refresh.

Deletion is irreversible from the app. Keep another usable backup if retention matters. Never broaden deletion to all Drive files or other backup families. On OAuth/network/Drive failure, show an error and do not report success.

### 8.4 Backup acceptance tests

| ID | Expected behavior | Recorded status |
|---|---|---|
| AT-001 | Baseline branch, HEAD and blobs recorded before edits | Applied for V4 deletion change |
| AT-002 | Full-system list scope, pagination and newest-first order | Source inspection; no automated pagination test |
| AT-003 | Backup payload fields and checksum correct | Existing source retained; not live-retested in V4 task |
| AT-004 | Invalid restore payload rejected before replacement | Source inspected; full device test not performed |
| AT-005 | Delete only selected backup after confirmation and refresh | Code verified; user confirmed operation |
| AT-006 | Cancel confirmation without deleting | Expected from code path; automated test not run |
| AT-007 | Drive/OAuth/network failure surfaced without false success | Error handling inspected; failure injection not tested |
| AT-008 | Plan backup unaffected | Scope reviewed; integration test not run |
| AT-009 | Audit backup unaffected | Scope reviewed; integration test not run |

## 9. Security and privacy

- The app PIN is a casual-access barrier, not encryption.
- IndexedDB and LocalStorage may contain customer, sales, plan and audit information; treat exports/backups as sensitive business data.
- Do not expose OAuth tokens, customer identifiers or backup contents in logs or public issues.
- Scope Drive list/restore/delete to the intended appDataFolder file family.
- Back up before destructive changes and document rollback.
- Do not overwrite original workbooks or project files without explicit approval.
- Do not add cloud collection or broader permissions without an explicit product decision.

## 10. End-to-end reconstruction and continuation procedure

A new engineer should be able to continue without relying on chat history:

1. Read this document, the project execution instructions, the approved V06 form specification and historical V2/V3 handovers if available.
2. Re-fetch repository main and inspect `index.html`, `manifest.webmanifest`, `sw.js`, README and changelog.
3. Map the live source into persistence, workbook parsing, joins, map/filter UI, Journey Plan, audit form, reporting, Drive integration, settings and PWA shell.
4. Create a baseline backup branch and record blob SHAs.
5. Run representative workbooks and record row counts, parser warnings, quality counters, matches, orders and statuses.
6. Compare the live auditor form against the original approved V06 question by question and conditional rule by conditional rule.
7. Define compatibility and migration plans before changing form identifiers or serialization.
8. Implement one bounded change set at a time.
9. Test parser fixtures, joins, plan import/export, form conditions, GPS fallback, backup create/list/restore/delete and migrations.
10. Run syntax, browser, relevant mobile and stored-data regression checks.
11. Fetch committed files again and verify the actual commit content.
12. Publish a handover with commit, changed files, tests, untested areas, risks, rollback and user action.

## 11. Acceptance test catalogue

| ID | Area | Acceptance criterion |
|---|---|---|
| AT-010 | V06 form | Q11/Q12 always shown; Q11A/Q13 conditional; Q24 absent; Q16/Q17 facing only |
| AT-011 | GPS fallback | Manual map pin works when GPS denied/unavailable; Get Location confirms coordinates |
| AT-012 | Data migration | Legacy forms/drafts survive migration and remain readable/exportable |
| AT-013 | Workbook parsing | Header scan, status exclusions, joins, dates and optional mapping match baseline |
| AT-014 | Journey Plan | Import validation, duplicate logic, approval flow and export remain compatible |
| AT-015 | PWA/offline | Local shell/data behavior documented accurately; Drive and fresh tiles require internet |

Record environment, test data, expected and actual result, and pass/fail for every test. Clearly identify tests not performed.

## 12. Open gaps and decisions

- Retrieve the original approved V06 DOCX and verify exact question wording, options and competitor lists.
- Confirm form eligibility rule with the product owner.
- Confirm whether historical Q16/Q17 numeric limits remain.
- Design and test backward-compatible migration for legacy forms and drafts.
- Define Targets metrics, grain, keys, dates, joins, missingness and reconciliation before implementation.
- Re-inspect exact sales status exclusions in live source.
- Re-check current application version and deployment state; older version numbers in previous handovers are historical only.
- A screenshot previously showed `popup_failed_to_open`; its root cause was not established and it must not be assumed related to backup deletion.

## 13. Engineering rules

- Preserve existing architecture unless a demonstrated reason requires change.
- Do not rewrite the application for a narrow feature.
- Never assume old line numbers, HEAD, dependency versions or deployment state remain current.
- Use the approved V06 document for form content and live source for implemented behavior.
- Back up before meaningful changes; do not overwrite originals without approval.
- Preserve records/drafts and use explicit backward-compatible migrations.
- Test syntax, runtime behavior, data compatibility and the relevant user workflow.
- Distinguish code inspection, automated tests, manual tests and user confirmation.
- Never claim full testing based only on compilation or one successful user scenario.

## 14. Required handover report after every stage

Every completed stage must report:
- Exact behavior added/changed and business reason.
- Changed file paths and whether each was new, replaced or edited; include blob SHAs.
- Branch, base HEAD, resulting commit and rollback branch.
- Test IDs, environment, results and failures.
- Untested scenarios and known limitations.
- Backup/migration and rollback details.
- Deployment status, distinguishing committed, deployed and live-verified.
- Required user action, if any.
- Remaining work, risks and decisions requiring approval.

## 15. Sources and evidence register

| Source | Use | Limitation |
|---|---|---|
| Current GitHub `main` source | Implemented behavior | Must be re-fetched for every future change |
| Approved `TPA Final Form Spec V06 Revised.docx` | Approved form requirements | Exact text/lists must be read from original |
| Supplied Excel workbooks | Parser fixtures and examples | Samples do not define universal rules |
| Historical V2/V3 technical handovers | Function maps, persistence, parser and prior context | Revalidate against current source |
| Project execution instructions TXT | Safe execution and QA workflow | Instructions, not application behavior |
| Commit `a88574667f9ef1299e62526ca1eb035df19b94cb` | Backup deletion traceability | Commit is not deployment evidence |
| User's explicit confirmation | Manual acceptance of backup deletion | Does not replace broader regression tests |

## Appendix A — Immediate next work

1. Re-fetch main and confirm the backup-deletion commit remains in history.
2. Read the original approved V06 DOCX and build a complete question/option/visibility specification.
3. Audit live form implementation against V06 and document every delta.
4. Decide eligibility and Q16/Q17 validation with the product owner.
5. Design backward-compatible migration for existing forms and drafts.
6. Implement and test V06 in a dedicated branch with before/after fixtures.
7. Scope Targets reporting separately after metric definitions are approved.

## Appendix B — Data-safety checklist

- [ ] Current branch and HEAD verified.
- [ ] Pre-change backup branch created.
- [ ] Original uploaded workbooks preserved.
- [ ] Existing IndexedDB and LocalStorage data preserved or explicitly migrated.
- [ ] Cloud backup families kept separate.
- [ ] Destructive operations target one selected object only.
- [ ] Confirmation shown before irreversible deletion/replacement.
- [ ] Rollback path documented and available.
- [ ] Syntax and acceptance tests run.
- [ ] Untested areas and deployment state reported honestly.

## Appendix C — Repository links

- Repository: https://github.com/ary-bo/tpa-navigation
- Main source: https://github.com/ary-bo/tpa-navigation/blob/main/index.html
- Backup deletion commit: https://github.com/ary-bo/tpa-navigation/commit/a88574667f9ef1299e62526ca1eb035df19b94cb
- Pre-change backup branch: https://github.com/ary-bo/tpa-navigation/tree/backup/main-before-backup-delete



### 6.5 Q12 wording: product-return question (implemented 14 November 2026)

At the user's request, auditor question Q12 now asks exactly: **«آیا تا به حال محصولات غنچه را مرجوع کرده‌اید ؟»**. The English translation dictionary label was aligned to “Have you ever returned Ghoncheh products?”.

Implementation scope: only the displayed Q12 question text (internal ID q12) and its English label were changed in index.html. The question identifier, number, single-choice type, three existing response options, visibility condition, and Q13 conditional display rule were preserved. No stored audit records or drafts were migrated or rewritten.

Verification: post-commit source fetch is required to confirm the exact persisted wording and unchanged Q12 options/visibility logic. Assistant browser/mobile interaction, saved-draft/export regression tests, and deployment/live verification have not been performed. The change is committed on main; do not treat that as evidence of deployment.

## Document revision history

| Version | Date | Scope | Evidence |
|---|---|---|---|
| 4.4 | 14 November 2026 | Reworded auditor Q12 to ask whether Ghoncheh products have ever been returned; preserved IDs, options and conditional rules | Code commit `95fcfee509a7e43800e4055a10a6b31a3b4cafe0`; source verification recorded in changelog |
| 4.3 | 14 November 2026 | Removed “حاشیه سود یا توافق” from auditor Q11A distribution-problem options; retained question ID, trigger and legacy translation mapping | Code commit `ef0581e12e0b7bb7f0cc1c08cde53fa9a299e610`; source/browser verification pending |
| 4.2 | 14 November 2026 | Reordered visit outcome and location correction; hid location/map and clear active values for “مشتری پیدا نشد” | Code commit `179b54b07a704a76cfdcfa29989bff9e6843bb7b`; user confirmed working |

## Appendix D — Mandatory continuous handover and document update policy

**This is a mandatory project requirement, not an optional administrative task.**

At the completion of every project phase, feature, module, milestone or separately deliverable work package, this Technical Reconstruction & Continuation Specification MUST be updated in full before that work package is considered closed or handed over.

The update must make it possible for a new engineer, maintainer, auditor or AI agent—with no access to prior chat conversations, undocumented verbal explanations or the previous responsible person's memory—to understand the current project state and reproduce, maintain, test and continue the completed work.

### D.1 Minimum information to record after every completed part

1. Exact scope: objective, business reason, included behavior, exclusions and acceptance criteria.
2. Final implementation design, architecture, important decisions, alternatives and rationale.
3. Every file, module, database/object store, storage key, API endpoint, integration, configuration and UI area added, changed or removed; record exact paths and identifiers.
4. Final data structures, schemas, field names, types, validation, transformations, dependencies and compatibility constraints.
5. Exact implementation and operational steps, commands, configuration, environment assumptions and permissions; never document secrets or credentials.
6. Business rules and conditional behavior, edge cases, error handling, fallback behavior and user-visible messages.
7. Migration details: source/destination formats, migration logic, versioning, preservation, backup, rollback and evidence that existing data remains usable.
8. Tests performed: IDs, setup, test data, browser/device/environment, expected and actual results, pass/fail; explicitly list tests not performed.
9. Repository traceability: branch, base commit, resulting commit SHA, changed paths and rollback reference.
10. Deployment/release status, distinguishing code committed, deployed and behavior verified in target environment.
11. Defects, unresolved questions, risks, technical debt, dependencies and decisions requiring product-owner approval.
12. Next recommended tasks in priority order, with enough context to begin without repeating discovery.

### D.2 Required update procedure

1. Re-inspect current repository and runtime state; never assume old commit IDs, file contents or deployment state are current.
2. Update the relevant existing sections so the document describes the actual current state, not merely append a short change note.
3. Update Document Control and Revision History with document version, date, completed scope and evidence.
4. Revise architecture, file maps, data contracts, workflows, diagrams where useful, acceptance tests, open issues and references wherever affected.
5. Mark each requirement and test accurately: implemented, verified, partially verified, not tested, deferred or blocked.
6. Preserve useful history, clearly mark superseded information and identify the new source of truth.
7. Verify internal consistency, links and commit identifiers, and confirm a new person can follow the handover without chat history.
8. Save and deliver the updated document with a concise change summary. The work package is not fully handed over until this update is complete.

### D.3 Mandatory close-out checklist

- [ ] Current source and repository state re-verified.
- [ ] All changed files and exact identifiers documented.
- [ ] Final behavior and business rules documented.
- [ ] Data structures, migrations, backup and rollback documented where relevant.
- [ ] Tests and actual outcomes recorded; untested cases disclosed.
- [ ] Commit, branch, deployment and release status recorded.
- [ ] Open risks, unresolved decisions and next steps recorded.
- [ ] Document Control and Revision History updated.
- [ ] A new responsible person can reproduce and continue the work using this document alone, without prior chat history.

### D.4 Definition of done for documentation

A project part is **NOT fully complete for handover purposes** if its implementation is finished but this document has not been updated to reflect the actual final state. The responsible person for that project part owns the update. If the work is later modified, corrected, rolled back or extended, the document must be revised again.

This policy applies to every subsequent project stage and every future version of this document. The next version must retain relevant history while updating all affected sections to the latest verified state.


### 6.6 Section 4 oil-availability gate (implemented 14 November 2026)

The first auditor question in Section 4 is now the required gate `q14_gate` (display number `۱۳A`): **«آیا محصولات روغن در فروشگاه موجود و قابل مشاهده است ؟»**. Its two options are **بله** and **خیر**. The English translation dictionary contains “Are cooking-oil products available and visible in the store?”.

Q14 (`q14`), Q15 (`q15`), Q16 (`q16`) and Q17 (`q17`) each have a visibility predicate requiring both the existing observation-block condition and `a.q14_gate === «بله»`. Thus they remain hidden until the gate is answered “بله”; choosing “خیر” hides all four. The existing form `update()` routine clears each newly hidden question's active answer and any associated OTHER text. The record serializer writes hidden mandatory questions as `N/A`. Existing Q14–Q17 IDs, wording, options, numeric constraints and Q17-versus-Q16 validation are unchanged. No bulk rewrite or migration of already-saved records or drafts was performed.

Verification evidence: post-commit fetch of `index.html` confirms the gate definition and all four visibility predicates. Source inspection confirms the existing hidden-answer cleanup and `N/A` serialization paths. Assistant browser/mobile interaction, saved-draft/export regression tests and deployment/live verification were not performed. Commit on `main` is not evidence of deployment.

## Document revision history

| Version | Date | Scope | Evidence |
|---|---|---|---|
| 4.5 | 14 November 2026 | Added Section 4 oil-availability gate before Q14; Q14–Q17 require “بله” | Code commit `f0c6fe08f5d4b9813ad7e077f765847337f17285`; source re-fetch verified |
| 4.4 | 14 November 2026 | Reworded auditor Q12 to ask whether Ghoncheh products have ever been returned; preserved IDs, options and conditional rules | Code commit `95fcfee509a7e43800e4055a10a6b31a3b4cafe0`; source verification recorded in changelog |
| 4.3 | 14 November 2026 | Removed “حاشیه سود یا توافق” from auditor Q11A distribution-problem options; retained question ID, trigger and legacy translation mapping | Code commit `ef0581e12e0b7bb7f0cc1c08cde53fa9a299e610` |
| 4.2 | 14 November 2026 | Reordered visit outcome and location correction; hid location/map and clear active values for “مشتری پیدا نشد” | Code commit `179b54b07a704a76cfdcfa29989bff9e6843bb7b`; user confirmed working |
