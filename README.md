# Purview Teams PST to HTML Converter

Converts a Microsoft Purview eDiscovery Teams PST export into a searchable, filterable HTML conversation report on Windows.

**Current version:** 1.0.32.2  
**Status:** Independent Cursor project (not the Hermes originals/output tree)

## What it does

- Temporarily attaches a PST via Outlook COM, reads message items, and writes one or two HTML reports.
- Exports Teams items from `TeamsMessagesData`, `TeamsMeetings`, `Migrated-Teams-Chat`, or `SubstrateHolds`.
- Exports Email items when that report is enabled using the message-class classifier (`IPM.Note*`, excluding meeting classes).
- GUI launcher (`PurviewTeamsPstToHtmlConverter.exe`) with progress, cancel, and automatic report open.
- Core converter (`Convert-PurviewTeamsPstToHtml.ps1`) runs under PowerShell 7 (`pwsh -Sta`).
- After conversion, verifies the PST was detached from Outlook (unless **Keep PST attached** is checked).
- If the PST was already attached before the run, reuses it and leaves it attached when finished.

## Requirements

- Windows 10/11
- Microsoft Outlook (installed and registered for COM)
- PowerShell 7 (`pwsh`)
- PS2EXE (for building EXEs): `Install-Module ps2exe -Scope CurrentUser`

## Project layout

```
PurviewTeamsPstToHtmlApp/
  src/
    Convert-PurviewTeamsPstToHtml.ps1   # Core converter (PS 7)
    Start-PurviewTeamsPstToHtmlApp.ps1  # WinForms launcher (PS 5.1, embeds core)
  build.ps1                             # Embed core + build Debug/Release EXEs
  build/                                # Built EXEs (gitignored)
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
pwsh -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Version 1.0.32.1
```

Outputs:

- `build\PurviewTeamsPstToHtmlConverter.exe` (GUI, no console)
- `build\PurviewTeamsPstToHtmlConverter_Debug.exe` (console for debugging)

Shipped copies for end users live in:

`C:\Users\pbush\OneDrive - Perfection Learning\Documents\AI\Cursor Output\PurviewTeamsPstToHtmlApp\`

## Quick test (no Outlook/PST)

```powershell
pwsh -NoProfile -File .\src\Start-PurviewTeamsPstToHtmlApp.ps1 -NoGui -UseSampleData `
  -OutputPath $env:TEMP\sample.html -LogPath $env:TEMP\sample.log
```

Expect `ItemsExported=4`, `TeamsItemsExported=2`, `EmailItemsExported=2`, `itemReadFailures=0`, exit code 0.

## Dual report behavior

- The GUI checkboxes map to `-TeamsReport:$true|$false` and `-EmailReport:$true|$false`.
- When both are checked, the shared display path is split into `_Teams` and `_Email` siblings for both report and log files.
- When only one is checked, only that report and log file are written.
- If both are unchecked, the core exits with `CONVERSION_ERROR`.

## Related paths

| Purpose | Path |
|---|---|
| Working copy (source) | `...\Cursor Working Directory\PurviewTeamsPstToHtmlApp` |
| Deliverables / EXEs | `...\Cursor Output\PurviewTeamsPstToHtmlApp` |
| Hermes originals (do not edit) | `...\Hermes Working Directory\PurviewTeamsPstToHtmlApp` |

## Recent changes (2026-07-09)

- Teams folder filter expanded: also includes `TeamsMeetings`, `Migrated-Teams-Chat`, and `SubstrateHolds`.
- Dual-report contract now includes explicit Teams/Email output paths and item counts.
- Sample-data smoke now expects `ItemsExported=4` with `TeamsItemsExported=2` and `EmailItemsExported=2`.
- Build now fails fast if the shared helper names drift out of sync.

## Recent changes (2026-07-08)

- Reverted T13 disk-spill (user prefers speed over lower RAM).
- Added Outlook PST detach verification + GUI warning on cleanup failure.
- Hardened cancel path: `taskkill` success requires exit code 0.
