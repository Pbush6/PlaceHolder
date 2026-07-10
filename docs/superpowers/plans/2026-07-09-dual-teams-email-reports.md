# Dual Teams + Email HTML Reports Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** From one PST scan, optionally write a Teams HTML report, an Email HTML report, or both, controlled by two default-on checkboxes with sibling `_Teams` / `_Email` path naming.

**Architecture:** Single Outlook PST pass classifies each message-like item into Teams (folder allowlist), Email (`IPM.Note*` minus meeting classes), or skip; then one or two HTML writers run. Launcher passes explicit `-TeamsReport` / `-EmailReport`, rewrites path boxes on checkbox change, and opens every generated report.

**Tech Stack:** PowerShell 7 core (`Convert-PurviewTeamsPstToHtml.ps1`), Windows PowerShell 5.1 WinForms launcher, Outlook COM, Pester verification suite, PS2EXE packaging.

**Spec:** `docs/superpowers/specs/2026-07-09-dual-teams-email-reports-design.md`

## Global Constraints

- Teams folders (unchanged): path segment `TeamsMessagesData|TeamsMeetings|Migrated-Teams-Chat|SubstrateHolds`
- Email: `MessageClass` matches `(?i)^IPM\.Note` and does **not** match `(?i)^IPM\.Schedule\.Meeting`
- Meeting invites/responses never go to the email report
- Both report types default **on**; both off → error before/at conversion start
- Path naming: both → base name in UI; files always `*_Teams.*` / `*_Email.*` when that type is selected
- One PST attach; no second EXE
- Do not edit Hermes originals; ship EXEs to Cursor Output deliverables after `build.ps1`
- Commits only when Patrick asks (skip commit steps unless he requests them)

## File map

| File | Responsibility |
|------|----------------|
| `src/Convert-PurviewTeamsPstToHtml.ps1` | Classify, bucket, resolve dual paths, write Teams and/or Email HTML, emit extended `CONVERSION_RESULT` |
| `src/Start-PurviewTeamsPstToHtmlApp.ps1` | Checkboxes, path rewrite, arg pass-through, success gate, open all reports/logs |
| `docs/DATA_CONTRACTS.md` | Document new result fields |
| `docs/schemas/stdout-contract.schema.json` | Schema for new fields |
| `scripts/Validate-DataContracts.ps1` | Assert new fields present on sample both-run |
| `tests/VerificationSuite.Tests.ps1` | Teams-only / Email-only / both / neither smokes |
| `README.md` | Document dual reports + checkbox/CLI behavior |
| `build.ps1` | Re-embed + package (no logic change beyond version bump when shipping) |

---

### Task 1: Pure path-suffix helpers + unit tests (launcher-side first)

**Files:**
- Create: `tests/PathSuffix.Tests.ps1`
- Modify: `src/Start-PurviewTeamsPstToHtmlApp.ps1` (add helper functions near other path helpers, before GUI; also callable from tests by dot-sourcing a small extracted block **or** duplicate the pure functions in the test file until Task 5 wires GUI — prefer defining helpers in launcher and copying the same function bodies into core in Task 2 for path resolution)

**Interfaces:**
- Produces:
  - `Get-ReportPathBaseName([string]$FilePath) -> string` — strips trailing `_Teams` / `_Email` before extension
  - `Get-ReportOutputPaths([string]$DisplayPath, [bool]$TeamsReport, [bool]$EmailReport) -> [pscustomobject]@{ DisplayPath; TeamsPath; EmailPath }` — DisplayPath is what the textbox should show; TeamsPath/EmailPath are `$null` when that type is off
  - Same for logs via extension `.log`

- [ ] **Step 1: Write failing tests**

```powershell
# tests/PathSuffix.Tests.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Dot-source only the helper region after Task 1 implements it in a tiny shared file:
# Create: src/ReportPathNaming.ps1 (dot-sourced by launcher + core + tests)
. (Join-Path $PSScriptRoot '..\src\ReportPathNaming.ps1')

Describe 'Report path naming' {
    It 'strips Teams and Email suffixes from base' {
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Teams.html' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Email.html' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages.html' | Should -Be 'LArtley Messages'
    }

    It 'both selected -> display base, write two sibling files' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.html' -TeamsReport $true -EmailReport $true
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages.html'
        $r.TeamsPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.EmailPath | Should -Be 'C:\out\LArtley Messages_Email.html'
    }

    It 'Teams only -> display and write _Teams' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.html' -TeamsReport $true -EmailReport $false
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.TeamsPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.EmailPath | Should -BeNullOrEmpty
    }

    It 'Email only -> display and write _Email' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages_Teams.html' -TeamsReport $false -EmailReport $true
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Email.html'
        $r.EmailPath | Should -Be 'C:\out\LArtley Messages_Email.html'
        $r.TeamsPath | Should -BeNullOrEmpty
    }
}
```

- [ ] **Step 2: Run tests — expect fail**

```powershell
pwsh -NoProfile -Command "Invoke-Pester -Path '.\tests\PathSuffix.Tests.ps1' -Output Detailed"
```

Expected: fail (missing `ReportPathNaming.ps1` / functions).

- [ ] **Step 3: Implement `src/ReportPathNaming.ps1`**

```powershell
Set-StrictMode -Version Latest

function Get-ReportPathBaseName {
    param([Parameter(Mandatory = $true)][string]$FilePath)
    $dir = [IO.Path]::GetDirectoryName($FilePath)
    $name = [IO.Path]::GetFileNameWithoutExtension($FilePath)
    $ext = [IO.Path]::GetExtension($FilePath)
    if ($name -match '(?i)_Teams$') { $name = $name.Substring(0, $name.Length - 6) }
    elseif ($name -match '(?i)_Email$') { $name = $name.Substring(0, $name.Length - 6) }
    return $name
}

function Get-ReportOutputPaths {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayPath,
        [Parameter(Mandatory = $true)][bool]$TeamsReport,
        [Parameter(Mandatory = $true)][bool]$EmailReport
    )
    if (-not $TeamsReport -and -not $EmailReport) {
        throw 'At least one of TeamsReport or EmailReport must be true.'
    }
    $dir = [IO.Path]::GetDirectoryName($DisplayPath)
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = (Get-Location).Path }
    $ext = [IO.Path]::GetExtension($DisplayPath)
    if ([string]::IsNullOrWhiteSpace($ext)) { $ext = '.html' }
    $base = Get-ReportPathBaseName -FilePath $DisplayPath
    $teamsPath = $null
    $emailPath = $null
    $display = $null
    if ($TeamsReport -and $EmailReport) {
        $display = Join-Path $dir ($base + $ext)
        $teamsPath = Join-Path $dir ($base + '_Teams' + $ext)
        $emailPath = Join-Path $dir ($base + '_Email' + $ext)
    }
    elseif ($TeamsReport) {
        $display = Join-Path $dir ($base + '_Teams' + $ext)
        $teamsPath = $display
    }
    else {
        $display = Join-Path $dir ($base + '_Email' + $ext)
        $emailPath = $display
    }
    [pscustomobject]@{ DisplayPath = $display; TeamsPath = $teamsPath; EmailPath = $emailPath }
}
```

- [ ] **Step 4: Re-run PathSuffix tests — expect PASS**

- [ ] **Step 5: Commit only if Patrick asks** (otherwise skip)

---

### Task 2: Core classification helpers + params

**Files:**
- Modify: `src/Convert-PurviewTeamsPstToHtml.ps1` (param block + helpers near `Test-IsTeamsMessagesFolder`)
- Dot-source or inline-copy: path naming from Task 1 (prefer `#region` copy of `Get-ReportOutputPaths` into core **or** read/embed `ReportPathNaming.ps1` at start of core — for packaged EXE, **inline the two functions into core and launcher** so embed stays one file; keep `ReportPathNaming.ps1` as the testable source of truth and sync by copy in build **or** duplicate carefully). **Decision for implementer:** duplicate the two functions into core and launcher (and keep `src/ReportPathNaming.ps1` for tests only), with a one-line comment `ponytail: keep in sync with ReportPathNaming.ps1`.

**Interfaces:**
- Produces:
  - `Test-IsEmailMessageClass([string]$MessageClass) -> bool`
  - `Get-ItemReportBucket([string]$FolderPath, [string]$MessageClass) -> 'Teams'|'Email'|'Skip'`
  - Params: `[bool]$TeamsReport = $true`, `[bool]$EmailReport = $true`

- [ ] **Step 1: Add failing classification tests** in `tests/Classification.Tests.ps1`

```powershell
. (Join-Path $PSScriptRoot '..\src\Convert-PurviewTeamsPstToHtml.ps1') # DO NOT — script executes. Instead extract helpers to src/ReportClassification.ps1 OR parse/dot-source a functions-only file.

# Prefer Create: src/ReportClassification.ps1 with Test-IsTeamsMessagesFolder, Test-IsEmailMessageClass, Get-ItemReportBucket
. (Join-Path $PSScriptRoot '..\src\ReportClassification.ps1')

Describe 'Item report bucket' {
    It 'Teams folder wins even for IPM.Note' {
        Get-ItemReportBucket -FolderPath 'X\TeamsMessagesData' -MessageClass 'IPM.Note' | Should -Be 'Teams'
    }
    It 'TeamsMeetings is Teams' {
        Get-ItemReportBucket -FolderPath 'X\SkypeSpacesData\TeamsMeetings' -MessageClass 'IPM.AppointmentSnapshot.SkypeTeams.Meeting' | Should -Be 'Teams'
    }
    It 'Inbox IPM.Note is Email' {
        Get-ItemReportBucket -FolderPath 'X\Inbox' -MessageClass 'IPM.Note' | Should -Be 'Email'
    }
    It 'IPM.Note.SMIME is Email' {
        Get-ItemReportBucket -FolderPath 'X\Inbox' -MessageClass 'IPM.Note.SMIME' | Should -Be 'Email'
    }
    It 'meeting request is Skip' {
        Get-ItemReportBucket -FolderPath 'X\Inbox' -MessageClass 'IPM.Schedule.Meeting.Request' | Should -Be 'Skip'
    }
    It 'contact is Skip' {
        Get-ItemReportBucket -FolderPath 'X\Contacts' -MessageClass 'IPM.Contact' | Should -Be 'Skip'
    }
}
```

- [ ] **Step 2: Run — expect fail**

- [ ] **Step 3: Implement `src/ReportClassification.ps1`**

```powershell
function Test-IsTeamsMessagesFolder {
    param([Parameter(Mandatory = $true)][string]$FolderPath)
    return $FolderPath -match '(?i)(^|\\)(TeamsMessagesData|TeamsMeetings|Migrated-Teams-Chat|SubstrateHolds)(\\|$)'
}

function Test-IsEmailMessageClass {
    param([AllowNull()][string]$MessageClass)
    if ([string]::IsNullOrWhiteSpace($MessageClass)) { return $false }
    if ($MessageClass -match '(?i)^IPM\.Schedule\.Meeting') { return $false }
    return $MessageClass -match '(?i)^IPM\.Note'
}

function Get-ItemReportBucket {
    param(
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [AllowNull()][string]$MessageClass
    )
    if (Test-IsTeamsMessagesFolder -FolderPath $FolderPath) { return 'Teams' }
    if (Test-IsEmailMessageClass -MessageClass $MessageClass) { return 'Email' }
    return 'Skip'
}
```

- [ ] **Step 4: In core, replace inline `Test-IsTeamsMessagesFolder` with the shared file content inlined (same bodies), add params:**

```powershell
[Parameter(Mandatory = $false)]
[bool]$TeamsReport = $true,

[Parameter(Mandatory = $false)]
[bool]$EmailReport = $true,
```

At start of `Invoke-ReportConversion`:

```powershell
if (-not $TeamsReport -and -not $EmailReport) {
    $msg = 'At least one report type must be selected (TeamsReport and/or EmailReport).'
    Write-Output ("CONVERSION_ERROR|{0}ExitCode=2|Message={1}" -f (Get-RunIdField), ($msg -replace '[\r\n|]', ' '))
    throw $msg
}
```

- [ ] **Step 5: Run Classification tests — PASS**

---

### Task 3: Core scan bucketing + dual path resolution + sample data

**Files:**
- Modify: `src/Convert-PurviewTeamsPstToHtml.ps1` (`Read-OutlookFolder`, `Get-SampleRecord`, `Invoke-ReportConversion`, stats, `CONVERSION_RESULT`)

**Interfaces:**
- Consumes: `Get-ItemReportBucket`, `Get-ReportOutputPaths`
- Produces: `$teamsRecords`, `$emailRecords` lists; `$script:TeamsOutputPath` / `$script:EmailOutputPath` / log counterparts; stats `TeamsItemsExported`, `EmailItemsExported`

- [ ] **Step 1: Extend stats**

```powershell
$script:Stats = [ordered]@{
    # ...existing...
    TeamsItemsExported = 0
    EmailItemsExported = 0
}
```

- [ ] **Step 2: Change `Read-OutlookFolder` to accept two lists** (`TeamsRecords`, `EmailRecords`) instead of one `Records`. After message-like check:

```powershell
$bucket = Get-ItemReportBucket -FolderPath $FolderPath -MessageClass $messageClass
if ($bucket -eq 'Teams' -and $TeamsReport) {
    [void]$TeamsRecords.Add((Get-MessageRecord -Item $item -FolderPath $FolderPath))
    $script:Stats.TeamsItemsExported++
    $script:Stats.ItemsExported++
}
elseif ($bucket -eq 'Email' -and $EmailReport) {
    [void]$EmailRecords.Add((Get-MessageRecord -Item $item -FolderPath $FolderPath))
    $script:Stats.EmailItemsExported++
    $script:Stats.ItemsExported++
}
else {
    # message-like but wrong bucket for this run, or Skip
    if (-not $loggedNonTeamsSkip -and $bucket -eq 'Skip' -and $isMessageLike) {
        Write-ReportLog "Skipping non-report item in folder: $FolderPath (class=$messageClass)"
        $loggedNonTeamsSkip = $true
    }
    $script:Stats.ItemsSkipped++
}
```

When `$TeamsReport` is false, Teams-bucket items count as skipped (do not export). Same for email.

- [ ] **Step 3: Resolve dual output paths after reading user `OutputPath`/`LogPath`**

```powershell
$htmlPaths = Get-ReportOutputPaths -DisplayPath $script:OutputPath -TeamsReport $TeamsReport -EmailReport $EmailReport
$logPaths  = Get-ReportOutputPaths -DisplayPath $script:LogPath -TeamsReport $TeamsReport -EmailReport $EmailReport
$script:TeamsOutputPath = $htmlPaths.TeamsPath
$script:EmailOutputPath = $htmlPaths.EmailPath
$script:TeamsLogPath = $logPaths.TeamsPath
$script:EmailLogPath = $logPaths.EmailPath
# Primary OutputPath/LogPath for backward compat = first requested (Teams preferred)
$script:OutputPath = if ($TeamsReport) { $script:TeamsOutputPath } else { $script:EmailOutputPath }
$script:LogPath    = if ($TeamsReport) { $script:TeamsLogPath } else { $script:EmailLogPath }
```

Assert safety for every non-null path. Open log writer(s): if both logs, write the same lines to both (tee) **or** write once then copy file at end — **v1 implement tee via a small `Write-ReportLog` that writes to all open writers**.

- [ ] **Step 4: Expand `Get-SampleRecord`**

Return object with `.Teams` and `.Email` arrays (or two functions). Sample:

- 2 Teams under `SamplePst\TeamsMessagesData` (existing)
- 2 Email under `SamplePst\Inbox`, `MessageClass='IPM.Note'`, subjects `Re: Budget` / `Budget`, distinct EntryIds
- Do not include meeting request or contact in exported lists (optional third skipped item only if you assert skip counts)

Wire `-UseSampleData` to fill both lists; respect `$TeamsReport`/`$EmailReport` when writing.

- [ ] **Step 5: Smoke core both-on sample**

```powershell
pwsh -NoProfile -File .\src\Convert-PurviewTeamsPstToHtml.ps1 -UseSampleData -NoPrompt `
  -TeamsReport:$true -EmailReport:$true `
  -OutputPath $env:TEMP\dual.html -LogPath $env:TEMP\dual.log
```

Expected: creates `dual_Teams.html`, `dual_Email.html`, matching logs; `CONVERSION_RESULT` includes new fields; exit 0.

(If path helper currently requires display base without suffix, pass `dual.html` as OutputPath.)

---

### Task 4: Email HTML writer (v1)

**Files:**
- Modify: `src/Convert-PurviewTeamsPstToHtml.ps1` (new functions after Teams HTML writers)

**Interfaces:**
- Consumes: email record list (same shape as `Get-MessageRecord`, plus optional ConversationTopic if you extend the record)
- Produces: `Write-EmailHtmlReport -Records -PstItem -ReportPath`

- [ ] **Step 1: Extend `Get-MessageRecord`** to capture conversation fields when present:

```powershell
$conversationTopic = Get-PropSafe -Object $Item -Name 'ConversationTopic' -Default ''
$conversationId = Get-PropSafe -Object $Item -Name 'ConversationID' -Default ''
# add to pscustomobject: ConversationTopic, ConversationId
```

- [ ] **Step 2: Add helpers**

```powershell
function Get-NormalizedEmailSubject {
    param([string]$Subject)
    $s = [string]$Subject
    while ($s -match '(?i)^(re|fw|fwd)\s*:\s*') { $s = $s -replace '(?i)^(re|fw|fwd)\s*:\s*','' }
    if ([string]::IsNullOrWhiteSpace($s)) { return '(no subject)' }
    return $s.Trim()
}

function Get-EmailThreadKey {
    param($Record)
    if (-not [string]::IsNullOrWhiteSpace([string]$Record.ConversationId)) {
        return 'id:' + [string]$Record.ConversationId
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Record.ConversationTopic)) {
        return 'topic:' + ([string]$Record.ConversationTopic).Trim().ToLowerInvariant()
    }
    $subj = Get-NormalizedEmailSubject -Subject $Record.Subject
    return "subj:$subj`n$($Record.ParticipantsKey)"
}
```

- [ ] **Step 3: Implement `Write-EmailHtmlReport`**

Clone Teams report structure with these differences:

- Title: `Purview Email PST Report - …`
- Hero: email wording
- Left filter: **Folders** checklist (`data-folder` on each conversation) + **People** checklist (reuse participant modes simply: Any / All)
- Group by `Get-EmailThreadKey`; title = normalized subject; subtitle = From list / participant summary
- Message card: From as speaker; To/Cc in details; folder in details
- JS: filter by selected folders (empty selection = show all **or** require at least one — **v1: no folder checked = show all folders**; same for people when none checked)
- Include `hero-credit` with `By Patrick Bush` (existing contract)
- Include date toolbar ids used by verification: `startDateFilter`, `endDateFilter`, `sortOrder`, `conversation-toolbar`, `participantMatchMode` (people mode select — keep id for contract grep)

- [ ] **Step 4: Call writer from `Invoke-ReportConversion`**

```powershell
if ($TeamsReport) {
    Write-ConversionStage -Stage 'WritingReport' # or WritingTeamsReport if you extend enum — prefer keep WritingReport and log "Writing Teams report"
    Write-HtmlReport -Records $teamsRecords -PstItem $pstItem -ReportPath $script:TeamsOutputPath
}
if ($EmailReport) {
    Write-ReportLog 'Writing Email HTML report.'
    Write-EmailHtmlReport -Records $emailRecords -PstItem $pstItem -ReportPath $script:EmailOutputPath
}
```

- [ ] **Step 5: Emit extended result**

```powershell
Write-Output ("CONVERSION_RESULT|{0}OutputPath={1}|LogPath={2}|ItemsExported={3}|ItemReadFailures={4}|AttachmentReadFailures={5}|SubfolderScanFailures={6}|TeamsOutputPath={7}|EmailOutputPath={8}|TeamsLogPath={9}|EmailLogPath={10}|TeamsItemsExported={11}|EmailItemsExported={12}" -f `
  (Get-RunIdField), $script:OutputPath, $script:LogPath, $script:Stats.ItemsExported, $script:Stats.ItemReadFailures, $script:Stats.AttachmentReadFailures, $script:Stats.SubfolderScanFailures,
  $script:TeamsOutputPath, $script:EmailOutputPath, $script:TeamsLogPath, $script:EmailLogPath,
  $script:Stats.TeamsItemsExported, $script:Stats.EmailItemsExported)
```

Empty string for unused paths when that report is off.

- [ ] **Step 6: Verify sample email HTML contains folder filter + thread markers**

```powershell
Select-String -LiteralPath $env:TEMP\dual_Email.html -Pattern 'folder-filter|data-folder|Purview Email' | Select-Object -First 5
```

---

### Task 5: Launcher GUI + NoGui wiring

**Files:**
- Modify: `src/Start-PurviewTeamsPstToHtmlApp.ps1`

**Interfaces:**
- Consumes: `Get-ReportOutputPaths` (duplicated helpers)
- Produces: passes `-TeamsReport:$true|$false` `-EmailReport:$true|$false`; success requires all requested files

- [ ] **Step 1: Extend param block + `ConvertTo-ArgumentList`**

```powershell
[bool]$TeamsReport = $true,
[bool]$EmailReport = $true,
# in ConvertTo-ArgumentList:
$args.Add('-TeamsReport'); $args.Add($TeamsReport.ToString())
$args.Add('-EmailReport'); $args.Add($EmailReport.ToString())
```

(Use bool parameters, not switches, so `$false` can be passed.)

- [ ] **Step 2: Add checkboxes** below Keep PST (shift Convert row down ~36px as needed):

```powershell
$teamsCheck = [System.Windows.Forms.CheckBox]::new()
$teamsCheck.Text = 'Teams report'
$teamsCheck.Checked = $true
# location near keepCheck

$emailCheck = [System.Windows.Forms.CheckBox]::new()
$emailCheck.Text = 'Email report'
$emailCheck.Checked = $true
```

- [ ] **Step 3: Path rewrite on checkbox change**

```powershell
function Update-PathsForReportSelection {
    $rHtml = Get-ReportOutputPaths -DisplayPath $outputBox.Text -TeamsReport $teamsCheck.Checked -EmailReport $emailCheck.Checked
    $rLog  = Get-ReportOutputPaths -DisplayPath $logBox.Text -TeamsReport $teamsCheck.Checked -EmailReport $emailCheck.Checked
    $outputBox.Text = $rHtml.DisplayPath
    $logBox.Text = $rLog.DisplayPath
}
$teamsCheck.add_CheckedChanged({ Update-PathsForReportSelection })
$emailCheck.add_CheckedChanged({ Update-PathsForReportSelection })
# Also call at end of Update-OutputPathsFromPstPath
```

Default PST-derived name: change `Get-SafeOutputBaseNameFromPstPath` to return `"$baseName Messages"` (not `Teams Messages`) so both-on base name is neutral.

- [ ] **Step 4: Convert click validation**

```powershell
if (-not $teamsCheck.Checked -and -not $emailCheck.Checked) {
    [System.Windows.Forms.MessageBox]::Show($form, 'At least one box must be checked (Teams report and/or Email report).', 'Report type required', 'OK', 'Error')
    return
}
```

Overwrite prompt: compute both target HTML paths; if any exist, list them all in the MessageBox.

- [ ] **Step 5: Parse `CONVERSION_RESULT` for Teams/Email paths; on success open all HTML files that exist; store arrays `$script:lastReportPaths` / `$script:lastLogPaths`. Open Report / Open Log iterate those arrays.**

- [ ] **Step 6: NoGui mode** — require at least one report flag; pass through; success = all requested outputs exist.

- [ ] **Step 7: Manual GUI smoke** (or NoGui):

```powershell
pwsh -NoProfile -File .\src\Start-PurviewTeamsPstToHtmlApp.ps1 -NoGui -UseSampleData `
  -TeamsReport:$true -EmailReport:$false `
  -OutputPath $env:TEMP\only_teams.html -LogPath $env:TEMP\only_teams.log
# expect only *_Teams.* files (or path already suffixed)
```

---

### Task 6: Contracts, verification suite, docs, build, ship

**Files:**
- Modify: `docs/DATA_CONTRACTS.md`, `docs/schemas/stdout-contract.schema.json`, `scripts/Validate-DataContracts.ps1`, `tests/VerificationSuite.Tests.ps1`, `README.md` (working + Output copies), `.cursor/rules/project.mdc`
- Run: `build.ps1 -Version 1.0.32.0` (or next unused), copy EXEs to Cursor Output

- [ ] **Step 1: Schema** — add optional string/integer properties to `resultLine`:

`TeamsOutputPath`, `EmailOutputPath`, `TeamsLogPath`, `EmailLogPath`, `TeamsItemsExported`, `EmailItemsExported`

Keep existing required fields. Update DATA_CONTRACTS table.

- [ ] **Step 2: VerificationSuite** — add cases:

```powershell
It 'sample both writes Teams and Email reports' { ... TeamsItemsExported=2; EmailItemsExported=2; files exist }
It 'sample Teams-only skips email file' { ... }
It 'sample Email-only skips teams file' { ... }
It 'both report flags false fails' {
  $result = & core -TeamsReport:$false -EmailReport:$false ...
  $result.ExitCode | Should -Not -Be 0
  $result.StdOut | Should -Match 'CONVERSION_ERROR'
}
```

Update any hard-coded `exported=2` expectations if totals change with email sample (Teams-only still 2; both may show `ItemsExported=4`).

- [ ] **Step 3: README** — document checkboxes, path naming, classifier, CLI bools.

- [ ] **Step 4: Build + verify**

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Version 1.0.32.0
pwsh -NoProfile -File .\scripts\Run-VerificationSuite.ps1
```

Expected: Suite Green, 8+ tests passing (count will increase).

- [ ] **Step 5: Copy EXEs** to `Cursor Output\PurviewTeamsPstToHtmlApp\` (if release EXE locked, write `*_1.0.32.0.exe` alternate and tell Patrick).

- [ ] **Step 6: Commit only if Patrick asks**

---

## Spec coverage checklist

| Spec requirement | Task |
|------------------|------|
| Folder → Teams, else IPM.Note* email, meeting skip | Task 2 |
| Checkboxes default on; both off error | Task 5 |
| Path base / `_Teams` / `_Email` rewrite | Tasks 1, 5 |
| Auto-open all; Open buttons open all | Task 5 |
| Email UX folder + people + threads | Task 4 |
| One PST pass two writers | Task 3–4 |
| Extended CONVERSION_RESULT | Tasks 3–4, 6 |
| Sample data Teams+Email | Task 3 |
| Verification + build/ship | Task 6 |

## Placeholder / consistency review

- Path helpers and classification live in small `src/*.ps1` files for tests; **must be inlined/duplicated into core + launcher** because the EXE embeds a single launcher with embedded core base64 (core cannot depend on sibling files at runtime).
- Bool params (not switches) for `$false` pass-through.
- No TBD left for classifier, naming, or open behavior.
