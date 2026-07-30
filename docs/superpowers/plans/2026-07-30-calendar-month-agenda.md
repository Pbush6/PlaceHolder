# Calendar Month Grid + Agenda Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the Calendar HTML report’s flat card list with the approved A2 month-grid + left agenda + detail panel layout, driven by existing Calendar records.

**Architecture:** Keep extraction and typed paths unchanged. Rewrite Calendar-only HTML/CSS/JS writers in `src/Convert-PurviewTeamsPstToHtml.ps1` so the report embeds a self-contained month board, chronological agenda, and detail pane that share one selection model and the existing filter dimensions. Extend `tests/CalendarContactsHtml.Tests.ps1` with structure, sync, encoding, and filter assertions.

**Tech Stack:** PowerShell 5.1-compatible report generation, static HTML/CSS/JS, Pester, existing `Get-CalendarRecord` fields.

**Spec:** `docs/superpowers/specs/2026-07-30-calendar-month-agenda-design.md`  
**Mockup:** `Cursor Output/PurviewTeamsPstToHtmlApp/Mockups/calendar-views-2026-07-30/A2-month-grid-agenda.html`

## Global Constraints

- Preserve exclusive bucket order and Calendar classification from 1.2.0.0.
- Recurring series remain one record; do not expand occurrences.
- HTML-encode every PST/sample-derived value in markup and attributes.
- Filters stay client-side via precomputed `data-*` / JSON payload fields; do not scan `textContent` to build indexes.
- Calendar output remains typed `*_Calendar.html` with matching log; launcher/browser open unchanged.
- Do not change Contacts/Teams/Email report UX except unavoidable shared helper cleanup.
- PowerShell 5.1-safe syntax; `LiteralPath` for spaced paths.
- No CDN dependency required for the shipped report (system fonts OK; do not require Google Fonts in production HTML).
- Do not edit the attached plan file during execution; do not commit unless Patrick asks.

---

### Task 1: Failing Calendar layout/sync contracts

**Files:**
- Modify: `tests/CalendarContactsHtml.Tests.ps1`
- Modify: `src/Convert-PurviewTeamsPstToHtml.ps1` (only if a tiny test seam is required; prefer asserting generated HTML)

**Interfaces:**
- Consumes: existing sample Calendar records and `Write-CalendarHtmlReport`
- Produces: failing Pester coverage that defines the A2 surface

- [ ] **Step 1: Add failing tests for the A2 shell**

Assert generated Calendar HTML contains:
- agenda region/list (`id` such as `calendarAgenda` / `calendarAgendaList`)
- month grid (`id` such as `calendarMonthGrid`)
- detail panel (`id` such as `calendarDetail`)
- desktop CSS with a skinny left agenda track and wider center month track (no orphan 10px resize track)
- no focusable `resizeHandle` in the Calendar report
- chronological agenda entries derived from records
- shared selection hooks (agenda item `data-entry-id` matching calendar chips)
- existing filters still present (search, from/to, folder, type, all-day, recurring)
- hostile encoded content still cannot inject markup

- [ ] **Step 2: Run focused Calendar HTML tests and confirm RED**

Run:
```powershell
pwsh -NoProfile -Command "Invoke-Pester -Path './tests/CalendarContactsHtml.Tests.ps1' -Output Detailed"
```
Expected: new assertions fail against the current card-list Calendar report.

- [ ] **Step 3: Commit only if Patrick requested commits; otherwise stop at RED evidence**

---

### Task 2: Calendar HTML shell, month placement, agenda, detail, sync JS

**Files:**
- Modify: `src/Convert-PurviewTeamsPstToHtml.ps1` (`Get-StaticRecordReportCss` Calendar path / new Calendar CSS helper, `Get-CalendarReportScript`, `Write-CalendarReportHeader`, `Write-CalendarRecordHtml` or replacement JSON/card emitters, `Write-CalendarHtmlReport`)
- Modify: `tests/CalendarContactsHtml.Tests.ps1` as needed for green
- Sync: launcher embedded core via established embed/build mechanism if the core script is inlined

**Interfaces:**
- Consumes: sorted Calendar records with existing fields (`StartTime`/`SortTime`, `EndTime`, `Subject`, `ItemType`, `AllDayEvent`, `Location`, `Organizer`, `RequiredAttendees`, `OptionalAttendees`, `IsRecurring`, `RecurrenceSummary`, `Categories`, `Sensitivity`, `FolderPath`, body/notes, attachments, timestamps, `MessageClass`, `EntryId`)
- Produces: self-contained Calendar HTML matching the A2 interaction model

- [ ] **Step 1: Emit a three-pane Calendar shell**

Replace the Calendar card-list chrome with:
- left sticky scrollable agenda
- center month toolbar + grid
- right sticky detail panel  
Keep typed log footer path behavior.

- [ ] **Step 2: Embed record data safely for client rendering**

Prefer one encoded JSON array (or equivalent per-record `data-*` chips + agenda rows) built with HTML/JSON-safe encoding. Include normalized filter fields and start-date keys for month placement.

- [ ] **Step 3: Implement month grid placement**

Place each visible record on its start date. Support month navigation defaulting to the earliest visible record’s month. Show up to a small chip budget per day plus “+N more”. Color-code meeting / all-day / recurring.

- [ ] **Step 4: Implement chronological agenda**

Render visible records oldest→newest, grouped by date. Clicking a row selects the same `EntryId`/stable id on the grid, scrolls/flashes the day, and populates the detail panel.

- [ ] **Step 5: Implement detail panel**

On selection, show the comprehensive fields already in the Calendar contract. Empty state when nothing is selected.

- [ ] **Step 6: Wire filters to both agenda and month**

Existing filter controls must hide/show records in both surfaces and update the visible count. Reset restores defaults.

- [ ] **Step 7: Run Calendar HTML tests GREEN, then Task 1–3 related suites**

Run:
```powershell
pwsh -NoProfile -Command "Invoke-Pester -Path './tests/CalendarContactsHtml.Tests.ps1','./tests/RecordExtraction.Tests.ps1','./tests/Classification.Tests.ps1','./tests/PathSuffix.Tests.ps1' -Output Detailed"
```
Expected: all pass.

- [ ] **Step 8: Sync embedded launcher core if required by repo convention**

---

### Task 3: Sample verification, docs, and release readiness

**Files:**
- Modify: `README.md` (Calendar report description)
- Modify: `docs/DATA_CONTRACTS.md` only if Calendar UX contract text must mention month/agenda
- Modify: version/build/package files only if Patrick wants this shipped as a versioned release in the same change
- Test: `tests/VerificationSuite.Tests.ps1` / release tests only if contracts or sample assertions mention Calendar HTML structure

**Interfaces:**
- Consumes: Task 2 writers
- Produces: documented behavior + verified sample Calendar HTML

- [ ] **Step 1: Generate sample Calendar HTML and visually confirm A2 layout**

Use `-UseSampleData` Calendar-only or all-four sample run. Confirm agenda, month chips, detail sync, and desktop column proportions. Save screenshots under Cursor Output if useful.

- [ ] **Step 2: Update README Calendar section to describe month grid + agenda + detail**

- [ ] **Step 3: Run canonical verification appropriate to scope**

At minimum:
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\scripts\Validate-DataContracts.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File .\scripts\Run-VerificationSuite.ps1
```
If packaging/version bump is in scope for this change, rebuild and run the explicit release artifact gate afterward.

- [ ] **Step 4: Update Curt memory / inflight with the Calendar UX decision and outcome**

---

## Done when

- Calendar report matches approved A2 interaction: skinny agenda, wider month, detail panel, bidirectional selection.
- Existing Calendar filters, encoding, typed paths, and extraction semantics still hold.
- Focused and canonical verification for the touched scope are green.
- README describes the new Calendar UX.
