# Dual Teams + Email HTML Reports — Design

**Date:** 2026-07-09  
**Status:** Approved (Patrick, 2026-07-09)  
**Project:** PurviewTeamsPstToHtmlApp  
**Version target:** next minor after 1.0.31.0

## Goal

From one PST conversion run, optionally produce a Teams conversation HTML report, an email HTML report, or both. Users choose via two GUI checkboxes (both on by default). Email content is whatever the Teams folder filter excludes, narrowed to real mail (`IPM.Note*`), not calendar/contacts/tasks/meeting invites.

## Non-goals (v1)

- Separate EXE or second PST attach for email
- MessageClass denylist as the primary email definition
- Meeting invites/responses in the email report (`IPM.Schedule.Meeting.*`)
- Full MIME / inline image extraction to disk
- Unread-only or server-side search index
- Merging Teams and email into one HTML page

---

## 1. Classification

Each message-like Outlook item is classified once during the existing recursive PST scan:

| Order | Condition | Destination |
|------|-----------|-------------|
| 1 | Folder path matches Teams allowlist segment: `TeamsMessagesData`, `TeamsMeetings`, `Migrated-Teams-Chat`, or `SubstrateHolds` (case-insensitive path segment) | **Teams** |
| 2 | Else `MessageClass` matches `IPM.Note*` **and** does **not** match meeting invite/response classes (`IPM.Schedule.Meeting.*`) | **Email** |
| 3 | Else | **Skip** (calendar, contacts, tasks, sticky notes, unknown junk, etc.) |

Notes:

- Teams routing remains folder-first because Purview Teams chats are often stored as `IPM.Note`.
- Email uses MessageClass allowlist (`IPM.Note*`) so unknown classes fail closed.
- Meeting invites/responses are explicitly excluded from email even if they appear outside Teams folders.
- An optional documented denylist of known non-mail prefixes may exist as a safety comment/helper, but allowlist + Teams folder gate are authoritative.

---

## 2. Run modes (GUI + CLI)

### GUI

- Checkboxes: **Teams report** and **Email report**.
- Both checked by default.
- Convert remains clickable when both are unchecked; on click, show an error that at least one report type must be selected (do not start conversion).

### Core / NoGui parameters

- Launcher and NoGui always pass explicit `-TeamsReport:$true|$false` and `-EmailReport:$true|$false` (do not rely on bare switch presence/absence).
- Core parameters default to `$true` when invoked directly without those args (preserves today’s “Teams report” behavior for simple core calls; email also defaults on when calling core directly — document in README; GUI/NoGui always pass both explicitly).
- If both false: fail with a clear `CONVERSION_ERROR` / non-zero exit (same rule as GUI).
- Sample data path must support Teams-only, Email-only, and both.

---

## 3. Paths, naming, and open behavior

One **output path** text box and one **log path** text box (existing layout).

| Checkboxes | Path boxes show | Files written |
|------------|-----------------|---------------|
| Both | Base name, e.g. `LArtley Messages.html` / `.log` | `LArtley Messages_Teams.html` + `LArtley Messages_Email.html` (same pattern for `.log`) |
| Teams only | `…_Teams.html` / `…_Teams.log` | Those exact paths |
| Email only | `…_Email.html` / `…_Email.log` | Those exact paths |

Rules:

- Changing checkboxes **rewrites** the path boxes so the displayed names match what will be written.
- Preserve the user’s base name when toggling by stripping/reapplying `_Teams` / `_Email` suffixes cleanly (before extension).
- On success, **auto-open every HTML report** that was generated.
- Status/progress text should name which report(s) are running and list final paths.
- Overwrite confirmation should cover every target file that already exists.

Default stamp filenames should follow the same suffix rules (base vs `_Teams` / `_Email`).

---

## 4. Email report UX (v1)

Reuse the Teams report shell (left filters, right pane, date range + sort toolbar) with email-oriented behavior:

**Grouping**

- One conversation card per email thread.
- Prefer Outlook conversation topic / conversation id when available.
- Fallback: normalized subject (strip `Re:` / `Fw:` / `Fwd:`) + sorted participant set.
- Messages within a thread sorted by time; threads sortable by newest/oldest (existing toolbar pattern).

**Filters**

- **Folders** — multi-select of folders that appear in the email export.
- **People** — From / To / Cc participants (any-selected / all-selected style modes acceptable for v1).
- Date range + sort controls retained.

**Message card**

- Subject as thread title.
- From prominent; To/Cc secondary or in details.
- Folder path in details.
- Body + attachment summary (same encoding/safety patterns as Teams).

---

## 5. Architecture

**Approach:** one PST pass, two in-memory buckets, two HTML writers (recommended over a separate email converter).

```text
Attach PST (existing)
  → Scan folders
      → Classify item → Teams list | Email list | skip
  → If TeamsReport: write Teams HTML (+ Teams log when separate)
  → If EmailReport: write Email HTML (+ Email log when separate)
Detach / cleanup (existing)
```

**Core**

- Extend `Read-OutlookFolder` (or post-classify after `Get-MessageRecord`) to bucket records.
- Keep existing Teams HTML writer; add email HTML writer sharing helpers (HTML encode, dates, attachments, progress stages).
- Stats should report per-type export counts (and skips).

**Launcher**

- Add checkboxes; wire path rewrite helper; validate at least one selection on Convert.
- Pass report-type flags into embedded core argument list.
- Success gate: exit code 0 **and** every requested report file exists on disk.
- Open each generated HTML on success.
- **Open Report** button: if both reports exist from the last run, open both; if only one exists, open that one.
- **Open Log** button: same rule for log files from the last run.

**Contracts**

- Extend stdout so the launcher learns both paths and counts. Preferred shape:
  - Keep `CONVERSION_RESULT` with primary fields for backward compatibility where possible, **and**
  - Add explicit fields such as `TeamsOutputPath`, `EmailOutputPath`, `TeamsLogPath`, `EmailLogPath`, `TeamsItemsExported`, `EmailItemsExported` (omit or empty when that report was not requested).
- Update `docs/DATA_CONTRACTS.md` and `docs/schemas/stdout-contract.schema.json`.
- Per-report log files when both run; each log is truncated for that report’s write path. Shared scan progress may be duplicated into both logs or written to both — **v1: write a full run log into each requested log file** (same content is acceptable) so either file is self-contained.

**Verification**

- Sample smokes: Teams-only, Email-only, both → expected export counts and file existence.
- GUI/NoGui: both unchecked → error, no conversion.
- Path suffix rewrite unit/regression coverage where practical.
- Rebuild via `build.ps1`; ship Debug/Release to Cursor Output deliverables.

---

## 6. Sample data

Extend `-UseSampleData` to include:

- Existing two Teams messages under `TeamsMessagesData`.
- At least two email messages under a non-Teams folder (e.g. `SamplePst\Inbox`) with `MessageClass = IPM.Note`.
- Optionally one skipped junk item (e.g. contact or sticky) to prove it is not exported.

Expectations:

- Both reports: Teams `exported=2`, Email `exported≥2`.
- Teams-only / Email-only: only the selected side is written.

---

## 7. Open questions resolved in this spec

| Topic | Decision |
|-------|----------|
| Email vs Teams | Folder allowlist → Teams; else `IPM.Note*` minus meeting classes → Email |
| Meeting invites | Excluded from email |
| GUI selection | Two checkboxes, both default on; both off → error on Convert |
| Naming | Sibling `_Teams` / `_Email`; both → base name in boxes |
| Path UI | One output + one log box; rewrite on checkbox change |
| Auto-open | Every generated HTML |
| Email UX | Folder + people filters; thread grouping |
| Architecture | Single PST pass, two writers |

---

## 8. Implementation follow-up

After Patrick approves this written spec, create an implementation plan (`writing-plans`) covering core classification, email HTML writer, launcher checkboxes/path rewrite, contract updates, sample data, tests, build, and ship.
