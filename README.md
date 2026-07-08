# Purview Teams PST to HTML Converter

Converts a Microsoft Purview eDiscovery Teams PST export into a searchable, filterable HTML conversation report on Windows.

**Current version:** 1.0.29.0  
**Status:** Independent Cursor project (not the Hermes originals/output tree)

## What it does

- Temporarily attaches a PST via Outlook COM, reads Teams message items, and writes one HTML report.
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

## Build

From this folder:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Version 1.0.29.0
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

Expect `exported=3`, `itemReadFailures=0`, exit code 0.

## Related paths

| Purpose | Path |
|---|---|
| Working copy (source) | `...\Cursor Working Directory\PurviewTeamsPstToHtmlApp` |
| Deliverables / EXEs | `...\Cursor Output\PurviewTeamsPstToHtmlApp` |
| Hermes originals (do not edit) | `...\Hermes Working Directory\PurviewTeamsPstToHtmlApp` |

## Recent changes (2026-07-08)

- Reverted T13 disk-spill (user prefers speed over lower RAM).
- Added Outlook PST detach verification + GUI warning on cleanup failure.
- Hardened cancel path: `taskkill` success requires exit code 0.
