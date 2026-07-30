# Purview PST Report Converter

Converts a Microsoft Purview eDiscovery PST into searchable Teams, Email, Calendar, and Contacts reports.

**Current version:** 1.2.1.0
**Status:** Independent Cursor project (not the Hermes originals/output tree)

## What it does

- Temporarily attaches a PST via Outlook COM and scans it once.
- All four report types are selected by default and are produced in one PST scan.
- Writes Teams items to `Base_Teams.html`.
- Writes Email items to `Base_Email.db`; Email HTML is no longer generated.
- Writes appointments and meetings to `Base_Calendar.html` with a month grid, chronological agenda, and synchronized detail pane.
- Writes contacts and distribution lists to `Base_Contacts.html`.
- Exports Teams items from `TeamsMessagesData`, `TeamsMeetings`, `Migrated-Teams-Chat`, or `SubstrateHolds`.
- Exports Email items when that report is enabled using the message-class classifier (`IPM.Note*`, excluding meeting classes).
- Exports Calendar items from Outlook calendar folders and Contacts items from Outlook contacts folders.
- GUI launcher (`PurviewTeamsPstToHtmlConverter.exe`) with progress, cancel, and automatic report open.
- Core converter (`Convert-PurviewTeamsPstToHtml.ps1`) runs under PowerShell 7 (`pwsh -Sta`).
- After conversion, verifies the PST was detached from Outlook (unless **Keep PST attached** is checked).
- If the PST was already attached before the run, reuses it and leaves it attached when finished.

## Requirements

- Windows 10/11
- Classic Microsoft Outlook (installed and registered for COM; new Outlook alone is not supported)
- PowerShell 7 (`pwsh`)
- .NET 8 Desktop Runtime for development/framework-dependent viewer builds only; the portable deployment package includes a self-contained viewer
- PS2EXE (for building EXEs): `Install-Module ps2exe -Scope CurrentUser`

The internal executables are unsigned, so Windows SmartScreen may warn. Close Outlook before conversion when possible. Outlook can expose only the record fields present in the PST; inaccessible or malformed individual properties are omitted without stopping the whole scan.

## Project layout

```
PurviewTeamsPstToHtmlApp/
  src/
    Convert-PurviewTeamsPstToHtml.ps1   # Core converter (PS 7)
    Start-PurviewTeamsPstToHtmlApp.ps1  # WinForms launcher (PS 5.1, embeds core)
  build.ps1                             # Embed core + build Debug/Release EXEs
  build/                                # Built EXEs (gitignored)
  EmailReviewViewer/                    # .NET 8 WinForms SQLite/FTS5 viewer + tests
  assets/                               # App icon
  docs/
    BUILD_NOTES.md                      # Packaging history
    FIX_PLAN_2026-07-07.md              # Hardening plan (completed)
  scripts/                              # Verification helpers
```

## Verification suite

Install a modern Pester locally if needed:

```powershell
Install-Module Pester -Scope CurrentUser -Force -SkipPublisherCheck
```

Run the canonical suite from the repo root:

```powershell
pwsh -NoProfile -File .\scripts\Run-VerificationSuite.ps1
```

The suite covers:
- parser checks for launcher, core, and regression script
- embedded-core consistency between launcher and core source
- core sample-data smoke path
- launcher sample-data smoke path
- permanent stop-process-tree regression script
- build smoke for debug/release EXEs

## Build

From this folder:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Version 1.2.1.0
```

Outputs:

- `build\PurviewTeamsPstToHtmlConverter.exe` (GUI, no console)
- `build\PurviewTeamsPstToHtmlConverter_Debug.exe` (console for debugging)
- `EmailReviewViewer\artifacts\publish\win-x64\EmailReviewViewer.App.exe` (framework-dependent viewer publish)

The no-console release converter is built with PS2EXE `-NoOutput`, so it does not
provide stdout or `CONVERSION_RESULT` to scripted callers. Automation that
consumes machine-readable stdout must run the source launcher or the packaged
Debug converter under `Tools\`. For the release converter, unattended `-NoGui`
success is determined from its exit code and the selected typed artifacts/logs.

Build the portable Windows x64 deployment folder and ZIP:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\scripts\Build-DeploymentPackage.ps1
```

This publishes Email Review Viewer self-contained as a single file, so target
computers do not need the .NET 8 Desktop Runtime. The release is written under
`Cursor Output\PurviewTeamsPstToHtmlApp\Releases`.

Shipped copies for end users live in:

`C:\Users\pbush\OneDrive - Perfection Learning\Documents\AI\Cursor Output\PurviewTeamsPstToHtmlApp\`

## Quick test (no Outlook/PST)

```powershell
pwsh -NoProfile -File .\src\Start-PurviewTeamsPstToHtmlApp.ps1 -NoGui -UseSampleData `
  -OutputPath $env:TEMP\sample.html -LogPath $env:TEMP\sample.log
```

Expect `ItemsExported=6`, `TeamsItemsExported=2`, `EmailItemsExported=2`, `CalendarItemsExported=1`, `ContactsItemsExported=1`, `itemReadFailures=0`, and exit code 0.

## Four-report behavior

- Teams, Email, Calendar, and Contacts are selected by default in the GUI. The checkboxes map to the four corresponding `*Report` switches.
- The shared base path produces `_Teams.html`, `_Email.db`, `_Calendar.html`, and `_Contacts.html`, plus one typed `.log` file for each selected report.
- Any subset may be selected; only those reports and logs are written. If all four are unchecked, the core exits with `CONVERSION_ERROR`.
- Email records are staged as UTF-8 NDJSON, imported into a temporary SQLite database, count-validated, then renamed into place.
- Teams HTML supports participant/conversation text search plus date and sort filters.
- Calendar HTML provides a navigable month grid, a scrollable chronological agenda of all matching meetings, and a sticky appointment detail pane. Agenda rows and month chips share selection; clicking an agenda item outside the visible month jumps the grid to that month. Search, date, folder, item-type, all-day, and recurring filters update both views.
- Contacts HTML supports text search plus folder and category filters.
- Email Review Viewer provides SQLite FTS5 search, folder filtering, date filtering, sorting, paging, and on-demand message detail.
- After a successful conversion, the GUI opens Teams, Calendar, and Contacts HTML in the default browser and opens Email databases in `EmailReviewViewer.App.exe`. Reports are not launched after a failed or incomplete conversion.
- In Email Review Viewer, use **File > Open Database…** or the **Open Database…** button to open or switch `.db` files. Invalid, corrupt, and incompatible databases are rejected without replacing the current database.

## Related paths

| Purpose | Path |
|---|---|
| Working copy (source) | `...\Cursor Working Directory\PurviewTeamsPstToHtmlApp` |
| Deliverables / EXEs | `...\Cursor Output\PurviewTeamsPstToHtmlApp` |
| Hermes originals (do not edit) | `...\Hermes Working Directory\PurviewTeamsPstToHtmlApp` |

## Recent changes (2026-07-30)

- Version 1.2.1.0 redesigns Calendar review as a three-pane month grid, chronological agenda of all matching meetings, and synchronized appointment detail pane.
- Detail fields omit Message class and Entry ID; All-day appears only for all-day appointments.

## Earlier changes (2026-07-29)

- Version 1.2.0.0 adds searchable Calendar and Contacts HTML reports.
- All four report types are selected by default, share one PST scan, and receive typed output/log paths.
- The launcher validates every selected output and log before opening HTML reports in the default browser or Email SQLite in the bundled reviewer.
- The six-item sample contains 2 Teams, 2 Email, 1 Calendar, and 1 Contacts record.

## Earlier changes (2026-07-13)

- Version 1.1.0.1 fixes packaged automatic viewer discovery and Windows PowerShell-safe database path passing.
- Email Review Viewer now starts without a sample database and can open existing databases from **File > Open Database…** or the **Open Database…** button.
- Version 1.1.0.0 integrates the standalone Email Review Viewer.
- Email report selection now produces SQLite/FTS5 (`_Email.db`) instead of large Email HTML.
- Import is offline, parameterized, duplicate-safe by Entry ID (deterministic fallback), and atomic via a temporary database.

## Earlier changes (2026-07-11)

- Email Folders filter shows message count next to each leaf label (e.g. `Inbox (2)`); checkbox value remains full folder path.
- Email Folders filter: all folders checked on load; uncheck to hide; none checked = show all (same as all checked). Select all / Clear buttons added. Keeps people search + large-report performance path.
- Large email HTML open freeze fixed earlier in 1.0.33.1 (no `textContent` indexing).

## Recent changes (2026-07-09)

- Teams folder filter expanded: also includes `TeamsMeetings`, `Migrated-Teams-Chat`, and `SubstrateHolds`.
- Dual-report contract now includes explicit Teams/Email output paths and item counts.
- Sample-data smoke now expects `ItemsExported=4` with `TeamsItemsExported=2` and `EmailItemsExported=2`.
- Build now fails fast if the shared helper names drift out of sync.

## Recent changes (2026-07-08)

- Reverted T13 disk-spill (user prefers speed over lower RAM).
- Added Outlook PST detach verification + GUI warning on cleanup failure.
- Hardened cancel path: `taskkill` success requires exit code 0.
