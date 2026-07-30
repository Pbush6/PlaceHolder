<#
.SYNOPSIS
Converts a Microsoft Purview eDiscovery PST export of Teams messages into a readable, filterable HTML conversation report.

.DESCRIPTION
Uses the locally installed Microsoft Outlook COM object to temporarily attach a PST to the current Outlook
profile, recursively read message-like items, and write a single searchable HTML report. The report includes
an easy participant filter so reviewers can show conversations involving only the people they choose. Messages
are displayed in a conversational format with different sender colors. The PST is detached from the Outlook
profile at the end when possible. No Microsoft 365 credentials are used and no data is uploaded.

.NOTES
Run in Windows PowerShell/PowerShell on the Windows machine where Outlook is installed. Close Outlook before
running for best results. For very large PSTs, the export can take a while.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Top-level script parameters are consumed inside helper functions; PSScriptAnalyzer does not always track script-scope captures.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseBOMForUnicodeEncodedFile', '', Justification = 'The report generator intentionally writes UTF-8 without BOM; PowerShell 7 parses this script correctly.')]
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$PstPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    # Optional: pre-select these people in the HTML filter when the report opens.
    # Example: -DefaultConversationParticipants 'Torey Page','Linda Artley'
    [Parameter(Mandatory = $false)]
    [string[]]$DefaultConversationParticipants = @(),

    [Parameter(Mandatory = $false)]
    [switch]$KeepPstAttached,

    [Parameter(Mandatory = $false)]
    [Alias('FlushEvery')]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$LogEvery = 100,

    # Renders a small built-in data set without Outlook/PST access. Useful for validating report output safely.
    [Parameter(Mandatory = $false)]
    [switch]$UseSampleData,

    # Packaging-friendly mode: fail clearly instead of prompting when required inputs are missing.
    [Parameter(Mandatory = $false)]
    [switch]$NoPrompt,

    # Optional correlation id. When non-empty, every machine-readable stdout line
    # (CONVERSION_PROGRESS / CONVERSION_STAGE / CONVERSION_RESULT) carries RunId=<id>
    # so the GUI can ignore any line not tagged with the current run's id.
    [Parameter(Mandatory = $false)]
    [string]$RunId = '',

    [Parameter(Mandatory = $false)]
    [bool]$TeamsReport = $true,

    [Parameter(Mandatory = $false)]
    [bool]$EmailReport = $true,

    [Parameter(Mandatory = $false)]
    [bool]$CalendarReport = $true,

    [Parameter(Mandatory = $false)]
    [bool]$ContactsReport = $true
)

$legacyReportFlagBound = $PSBoundParameters.ContainsKey('TeamsReport') -or $PSBoundParameters.ContainsKey('EmailReport')
$newReportFlagBound = $PSBoundParameters.ContainsKey('CalendarReport') -or $PSBoundParameters.ContainsKey('ContactsReport')
if ($legacyReportFlagBound -and -not $newReportFlagBound) {
    # Compatibility for pre-Calendar/Contacts direct callers: explicitly binding either
    # legacy report flag without either new flag retains the historical two-report surface.
    $CalendarReport = $false
    $ContactsReport = $false
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Stats = [ordered]@{
    StartedAt               = Get-Date
    FoldersScanned          = 0
    ItemsAttempted          = 0
    ItemsExported           = 0
    ItemsSkipped            = 0
    TeamsItemsExported      = 0
    EmailItemsExported      = 0
    CalendarItemsExported   = 0
    ContactsItemsExported   = 0
    ItemReadFailures        = 0
    AttachmentReadFailures  = 0
    SubfolderScanFailures   = 0
}

# Single no-BOM, auto-flushing writer over the log file. Opened in Invoke-ReportConversion
# (create/truncate = fresh log each run) and disposed in its finally block.
$script:LogWriter = $null
$script:LogWriters = @()

function ConvertTo-NormalizedInputPath {
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }

    # Supports pasted paths, quoted paths, and paths dragged from File Explorer.
    $clean = $Path.Trim()
    $clean = $clean.Trim('"').Trim("'")

    if ($clean.StartsWith('file://', [System.StringComparison]::OrdinalIgnoreCase)) {
        try {
            $clean = ([System.Uri]$clean).LocalPath
        }
        catch {
            Write-Verbose "Could not parse file URI as a local path: $clean"
            # Keep the original cleaned text so normal path validation can report it.
        }
    }

    return [Environment]::ExpandEnvironmentVariables($clean)
}

function Resolve-OutputFilePath {
    param(
        [AllowNull()][string]$Path,
        [Parameter(Mandatory = $true)][string]$DefaultDirectory,
        [Parameter(Mandatory = $true)][string]$DefaultFileName
    )

    $clean = ConvertTo-NormalizedInputPath -Path $Path
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return (Join-Path $DefaultDirectory $DefaultFileName)
    }

    $parent = Split-Path -LiteralPath $clean
    if ([string]::IsNullOrWhiteSpace($parent)) {
        return (Join-Path $DefaultDirectory $clean)
    }

    return $clean
}

function Assert-OutputPathsSafe {
    param(
        [Parameter(Mandatory = $true)][string]$ReportPath,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    if ([string]::IsNullOrWhiteSpace($ReportPath)) { throw 'OutputPath is required.' }
    if ([string]::IsNullOrWhiteSpace($LogPath)) { throw 'LogPath is required.' }

    try { $reportFullPath = [IO.Path]::GetFullPath($ReportPath) }
    catch { throw "OutputPath is not a valid file path: $ReportPath. $($_.Exception.Message)" }
    try { $logFullPath = [IO.Path]::GetFullPath($LogPath) }
    catch { throw "LogPath is not a valid file path: $LogPath. $($_.Exception.Message)" }

    if ($reportFullPath -ieq $logFullPath) {
        throw 'OutputPath and LogPath must be different files. Choose a separate .html report path and .log path.'
    }

    foreach ($pathToCheck in @($reportFullPath, $logFullPath)) {
        if (Test-Path -LiteralPath $pathToCheck -PathType Container) {
            throw "Output path points to a folder, not a file: $pathToCheck"
        }
        $parent = Split-Path -LiteralPath $pathToCheck
        if ([string]::IsNullOrWhiteSpace($parent)) { throw "Output path must include a parent folder: $pathToCheck" }
        if (Test-Path -LiteralPath $parent -PathType Leaf) {
            throw "Output parent path is a file, not a folder: $parent"
        }
    }
}

function Request-PstPath {
    param([AllowNull()][string]$InitialPath)

    $candidate = ConvertTo-NormalizedInputPath -Path $InitialPath
    if ($NoPrompt -and ([string]::IsNullOrWhiteSpace($candidate) -or -not (Test-Path -LiteralPath $candidate -PathType Leaf))) {
        throw "PST file path is required and must point to an existing file when -NoPrompt is used. Provided path: $candidate"
    }

    while ([string]::IsNullOrWhiteSpace($candidate) -or -not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            Write-Warning "PST file not found: $candidate"
        }

        Write-Information '' -InformationAction Continue
        Write-Information 'Enter the full path to the PST file you want to convert.' -InformationAction Continue
        Write-Information 'Tip: you can drag the .pst file from File Explorer into this window, then press Enter.' -InformationAction Continue
        $typed = Read-Host 'PST file path'
        $candidate = ConvertTo-NormalizedInputPath -Path $typed

        if ([string]::IsNullOrWhiteSpace($candidate)) {
            Write-Information 'No PST path entered. Press Ctrl+C to cancel, or enter a PST path.' -InformationAction Continue
        }
    }

    $item = Get-Item -LiteralPath $candidate
    if ($item.Extension -ne '.pst') {
        Write-Warning "The selected file does not end in .pst: $($item.FullName)"
        $continue = Read-Host 'Continue anyway? Type Y to continue'
        if ($continue -notin @('Y','y','Yes','yes')) {
            return Request-PstPath -InitialPath $null
        }
    }

    return $item.FullName
}

function Write-ReportLog {
    param(
        [Parameter(Mandatory = $true)] [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )

    # Collapse CR/LF so a message (e.g. a folder name) with an embedded newline cannot forge
    # extra log lines or break the single-line stdout parser downstream.
    $safeMessage = ([string]$Message) -replace "[\r\n]+", ' '
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 's'), $Level, $safeMessage
    foreach ($writer in @($script:LogWriters)) {
        if ($null -ne $writer) { $writer.WriteLine($line) }
    }
    Write-Information $line -InformationAction Continue
}

function Get-RunIdField {
    # Emits "RunId=<id>|" for tagged runs (GUI), or '' when -RunId is empty (NoGui),
    # so the field order is otherwise unchanged and the NoGui '-like' matches still work.
    if ([string]::IsNullOrEmpty($RunId)) { return '' }
    return "RunId=$RunId|"
}

function Write-ConversionStage {
    # Machine-readable report-stage marker for the GUI progress bar. Stage names and
    # numeric Extra fields are literals, so no untrusted text is interpolated here.
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [AllowNull()][string]$Extra
    )
    $line = "CONVERSION_STAGE|$(Get-RunIdField)Stage=$Stage"
    if (-not [string]::IsNullOrEmpty($Extra)) { $line += "|$Extra" }
    Write-Output $line
}

function Write-ConversionProgress {
    param([Parameter(Mandatory = $true)] [string]$FolderPath)

    $elapsed = (Get-Date) - $script:Stats.StartedAt
    $elapsedSeconds = [Math]::Max(1, [int][Math]::Round($elapsed.TotalSeconds))
    $ratePerMinute = [Math]::Round(($script:Stats.ItemsExported / $elapsedSeconds) * 60, 1)
    $safeFolderPath = (([string]$FolderPath) -replace "[\r\n]+", ' ').Replace('\', '/').Replace('|', '/')
    Write-Output ("CONVERSION_PROGRESS|{0}ItemsAttempted={1}|ItemsExported={2}|FoldersScanned={3}|ItemReadFailures={4}|ElapsedSeconds={5}|RatePerMinute={6}|FolderPath={7}" -f (Get-RunIdField), $script:Stats.ItemsAttempted, $script:Stats.ItemsExported, $script:Stats.FoldersScanned, $script:Stats.ItemReadFailures, $elapsedSeconds, $ratePerMinute, $safeFolderPath)
}

function Write-ConversionError {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [AllowNull()][string]$ExitCode = '1'
    )
    $safeMessage = (([string]$Message) -replace "[\r\n]+", ' ').Replace('|', '/')
    if ($safeMessage.Length -gt 500) { $safeMessage = $safeMessage.Substring(0, 500) }
    Write-Output ("CONVERSION_ERROR|{0}ExitCode={1}|Message={2}" -f (Get-RunIdField), $ExitCode, $safeMessage)
}

function ConvertTo-HtmlEncodedText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-HtmlBody {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '<em>(empty)</em>' }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return '<em>(empty)</em>' }
    $encoded = [System.Net.WebUtility]::HtmlEncode($text.Trim())
    return ($encoded -replace "`r?`n", '<br/>')
}

function Close-ComObjectSafe {
    param([AllowNull()][object]$ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Test-OutlookRpcRejected {
    param([AllowNull()][System.Exception]$Exception)

    $current = $Exception
    while ($null -ne $current) {
        if ($current.HResult -eq -2147418111) { return $true } # RPC_E_CALL_REJECTED / 0x80010001
        if ($current.Message -match 'RPC_E_CALL_REJECTED|Call was rejected by callee') { return $true }
        $current = $current.InnerException
    }
    return $false
}

function Invoke-OutlookComOperation {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory = $true)][string]$Operation,
        [int]$MaxAttempts = 8,
        [int]$DelaySeconds = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try { return & $ScriptBlock }
        catch {
            if ((Test-OutlookRpcRejected -Exception $_.Exception) -and $attempt -lt $MaxAttempts) {
                Write-ReportLog ("Outlook was busy while $Operation; retrying attempt $($attempt + 1) of $MaxAttempts. Close Outlook if this continues. $($_.Exception.Message)") 'WARN'
                Start-Sleep -Seconds $DelaySeconds
                continue
            }
            if (Test-OutlookRpcRejected -Exception $_.Exception) {
                throw "Outlook did not respond while $Operation after $MaxAttempts attempts. Close Outlook completely, wait a few seconds, then run the converter again. Original error: $($_.Exception.Message)"
            }
            throw
        }
    }
}

function Get-PropSafe {
    param(
        [Parameter(Mandatory = $true)] [object]$Object,
        [Parameter(Mandatory = $true)] [string]$Name,
        [AllowNull()][object]$Default = $null
    )
    try { return $Object.$Name } catch { return $Default }
}

function ConvertTo-NormalizedPersonName {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    $name = [string]$Value
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }

    # Remove addresses and Exchange legacy DN fragments, then normalize common company suffixes.
    $name = $name -replace '<[^>]*>', ''
    $name = $name -replace '/O=.*$', ''
    $name = $name -replace '\s+-\s+(Mastery Education|Perfection Learning).*$', ''
    $name = $name.Trim()

    # Convert "Last, First" to "First Last" for easier filtering.
    if ($name -match '^([^,]+),\s*([^,]+)$') {
        $name = ("$($Matches[2]) $($Matches[1])").Trim()
    }

    # If only an email address remains, keep the local part as a readable fallback.
    if ($name -match '^[^@\s]+@[^@\s]+$') {
        $name = ($name -replace '@.*$', '')
    }

    return (($name -replace '\s+', ' ').Trim())
}

function Test-LikelyPersonName {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $false }
    $name = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { return $false }
    if ($name -match '^Fireflies\.ai\b') { return $false }
    if ($name.Length -gt 60) { return $false }
    if ($name -match '@|https?://|www\.|/|\\|\d|[{}\[\]_]|\b(thread|communication|meeting|call|teamsvisitor|visitor|unknown|recipient|exchange|admin|onmicrosoft|perfectionlearning\.com)\b') { return $false }

    $parts = @($name -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -lt 2 -or $parts.Count -gt 5) { return $false }

    foreach ($part in $parts) {
        if ($part -notmatch "^[\p{L}][\p{L}'.\-]*$") { return $false }
    }

    return $true
}

function Split-RecipientName {
    param([AllowNull()][object]$Value)

    $names = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $names }

    foreach ($part in ([string]$Value -split '\s*;\s*')) {
        $normalized = ConvertTo-NormalizedPersonName $part
        if (-not [string]::IsNullOrWhiteSpace($normalized)) { [void]$names.Add($normalized) }
    }
    return $names
}

function Get-ParticipantName {
    param(
        [AllowNull()][object]$SenderName,
        [AllowNull()][object]$SenderEmail,
        [AllowNull()][object]$To,
        [AllowNull()][object]$Cc
    )

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $senderNameValue = ConvertTo-NormalizedPersonName $SenderName
    if ([string]::IsNullOrWhiteSpace($senderNameValue)) { $senderNameValue = ConvertTo-NormalizedPersonName $SenderEmail }
    if (-not [string]::IsNullOrWhiteSpace($senderNameValue)) { [void]$set.Add($senderNameValue) }

    foreach ($n in (Split-RecipientName $To)) { [void]$set.Add($n) }
    foreach ($n in (Split-RecipientName $Cc)) { [void]$set.Add($n) }

    return @($set | Sort-Object)
}

function Get-ParticipantKey {
    param([AllowNull()][string[]]$Participants)
    if ($null -eq $Participants -or $Participants.Count -eq 0) { return '(unknown participants)' }
    return (($Participants | Sort-Object -Unique) -join ' || ')
}

function Get-ParticipantDisplayOrder {
    param(
        [AllowNull()][string[]]$Participants,
        [AllowNull()][hashtable]$PreferredParticipants
    )

    $names = @($Participants | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    if ($names.Count -eq 0) { return @() }

    return @($names | Sort-Object `
        @{ Expression = { if ($null -ne $PreferredParticipants -and $PreferredParticipants.ContainsKey($_)) { 0 } else { 1 } } }, `
        @{ Expression = { $_ } })
}

function Format-ParticipantPreview {
    param(
        [AllowNull()][string[]]$Participants,
        [int]$MaximumNames = 4,
        [AllowNull()][hashtable]$PreferredParticipants
    )

    $names = @(Get-ParticipantDisplayOrder -Participants $Participants -PreferredParticipants $PreferredParticipants)
    if ($names.Count -eq 0) { return '' }

    if ($names.Count -gt $MaximumNames) {
        return (($names | Select-Object -First $MaximumNames) -join ', ') + ', ...'
    }

    return ($names -join ', ')
}

function Get-AttachmentSummaryHtml {
    param([AllowNull()][object]$Item)

    $rows = New-Object System.Collections.Generic.List[string]
    foreach ($attachment in @(Get-AttachmentMetadata -Item $Item)) {
        [void]$rows.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>' -f (ConvertTo-HtmlEncodedText $attachment.FileName), (ConvertTo-HtmlEncodedText $attachment.DisplayName), (ConvertTo-HtmlEncodedText $attachment.Size)))
    }
    if ($rows.Count -le 0) { return '' }
    return "<div class='attachments'><strong>Attachments:</strong><table><thead><tr><th>File name</th><th>Display name</th><th>Size bytes</th></tr></thead><tbody>$($rows -join "`n")</tbody></table></div>"
}

function Get-AttachmentMetadata {
    param([AllowNull()][object]$Item)

    $attachments = $null
    $metadata = New-Object System.Collections.Generic.List[object]
    try {
        $attachments = Get-PropSafe -Object $Item -Name 'Attachments' -Default $null
        if ($null -eq $attachments) { return @() }
        $count = [int](Get-PropSafe -Object $attachments -Name 'Count' -Default 0)
        if ($count -le 0) { return @() }
        for ($i = 1; $i -le $count; $i++) {
            $attachment = $null
            try {
                $attachment = $attachments.Item($i)
                [void]$metadata.Add([pscustomobject]@{
                    FileName = [string](Get-PropSafe -Object $attachment -Name 'FileName' -Default '')
                    DisplayName = [string](Get-PropSafe -Object $attachment -Name 'DisplayName' -Default '')
                    Size = (Get-PropSafe -Object $attachment -Name 'Size' -Default $null)
                })
            }
            catch {
                $script:Stats.AttachmentReadFailures++
                Write-ReportLog "Could not read attachment $i on item. $($_.Exception.Message)" 'WARN'
            }
            finally {
                Close-ComObjectSafe $attachment
            }
        }
    }
    catch {
        $script:Stats.AttachmentReadFailures++
        Write-ReportLog "Could not enumerate attachments on item. $($_.Exception.Message)" 'WARN'
    }
    finally {
        Close-ComObjectSafe $attachments
    }
    return $metadata.ToArray()
}

function Test-MissingDate {
    # Outlook/MAPI returns 1/1/4501 (year 4501) as a "no date" sentinel rather than $null, which
    # would otherwise sort items to the far future and display an absurd timestamp. Treat empty,
    # null, and any date in year >= 4500 as missing.
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $true }
    if ($Value -is [datetime] -and $Value.Year -ge 4500) { return $true }
    return $false
}

# ponytail: keep in sync with ReportPathNaming.ps1
function Get-ReportPathBaseName {
    param([Parameter(Mandatory = $true)][string]$FilePath)
    $name = [IO.Path]::GetFileNameWithoutExtension($FilePath)
    if ($name -match '(?i)_Teams$') { $name = $name.Substring(0, $name.Length - 6) }
    elseif ($name -match '(?i)_Email$') { $name = $name.Substring(0, $name.Length - 6) }
    elseif ($name -match '(?i)_Calendar$') { $name = $name.Substring(0, $name.Length - 9) }
    elseif ($name -match '(?i)_Contacts$') { $name = $name.Substring(0, $name.Length - 9) }
    return $name
}

# ponytail: keep in sync with ReportPathNaming.ps1
function Get-ReportOutputPaths {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayPath,
        [Parameter(Mandatory = $true)][bool]$TeamsReport,
        [Parameter(Mandatory = $true)][bool]$EmailReport,
        [bool]$CalendarReport = $false,
        [bool]$ContactsReport = $false
    )
    if (-not $TeamsReport -and -not $EmailReport -and -not $CalendarReport -and -not $ContactsReport) {
        throw 'At least one report type must be selected.'
    }
    $dir = [IO.Path]::GetDirectoryName($DisplayPath)
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = (Get-Location).Path }
    $inputExt = [IO.Path]::GetExtension($DisplayPath)
    $isLogPath = $inputExt -in @('.log', '.txt')
    $teamsExt = if ($isLogPath) { $inputExt } else { '.html' }
    $emailExt = if ($isLogPath) { $inputExt } else { '.db' }
    $calendarExt = if ($isLogPath) { $inputExt } else { '.html' }
    $contactsExt = if ($isLogPath) { $inputExt } else { '.html' }
    $displayExt = if ($isLogPath) { $inputExt } else { '.html' }
    $base = Get-ReportPathBaseName -FilePath $DisplayPath
    $teamsPath = if ($TeamsReport) { Join-Path $dir ($base + '_Teams' + $teamsExt) } else { $null }
    $emailPath = if ($EmailReport) { Join-Path $dir ($base + '_Email' + $emailExt) } else { $null }
    $calendarPath = if ($CalendarReport) { Join-Path $dir ($base + '_Calendar' + $calendarExt) } else { $null }
    $contactsPath = if ($ContactsReport) { Join-Path $dir ($base + '_Contacts' + $contactsExt) } else { $null }
    $selectedPaths = @(@($teamsPath, $emailPath, $calendarPath, $contactsPath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $display = if (@($selectedPaths).Count -eq 1) { $selectedPaths[0] } else { Join-Path $dir ($base + $displayExt) }
    [pscustomobject]@{
        DisplayPath = $display
        TeamsPath = $teamsPath
        EmailPath = $emailPath
        CalendarPath = $calendarPath
        ContactsPath = $contactsPath
    }
}

# ponytail: keep in sync with ReportClassification.ps1
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

function Test-IsCalendarMessageClass {
    param([AllowNull()][string]$MessageClass)
    if ([string]::IsNullOrWhiteSpace($MessageClass)) { return $false }
    return $MessageClass -match '(?i)^IPM\.(Appointment|Schedule\.Meeting)'
}

function Test-IsContactsMessageClass {
    param([AllowNull()][string]$MessageClass)
    if ([string]::IsNullOrWhiteSpace($MessageClass)) { return $false }
    return $MessageClass -match '(?i)^IPM\.(Contact|DistList)'
}

function Get-ItemReportBucket {
    param(
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [AllowNull()][string]$MessageClass
    )
    if (Test-IsTeamsMessagesFolder -FolderPath $FolderPath) { return 'Teams' }
    if (Test-IsEmailMessageClass -MessageClass $MessageClass) { return 'Email' }
    if (Test-IsCalendarMessageClass -MessageClass $MessageClass) { return 'Calendar' }
    if (Test-IsContactsMessageClass -MessageClass $MessageClass) { return 'Contacts' }
    return 'Skip'
}

function Get-MessageRecord {
    param(
        [Parameter(Mandatory = $true)] [object]$Item,
        [Parameter(Mandatory = $true)] [string]$FolderPath
    )

    $subject = Get-PropSafe -Object $Item -Name 'Subject' -Default ''
    $messageClass = Get-PropSafe -Object $Item -Name 'MessageClass' -Default ''
    $senderName = Get-PropSafe -Object $Item -Name 'SenderName' -Default ''
    $senderEmail = Get-PropSafe -Object $Item -Name 'SenderEmailAddress' -Default ''
    $to = Get-PropSafe -Object $Item -Name 'To' -Default ''
    $cc = Get-PropSafe -Object $Item -Name 'CC' -Default ''
    $sentOn = Get-PropSafe -Object $Item -Name 'SentOn' -Default $null
    $receivedTime = Get-PropSafe -Object $Item -Name 'ReceivedTime' -Default $null
    $creationTime = Get-PropSafe -Object $Item -Name 'CreationTime' -Default $null
    $entryId = Get-PropSafe -Object $Item -Name 'EntryID' -Default ''
    $body = Get-PropSafe -Object $Item -Name 'Body' -Default ''
    $conversationTopic = Get-PropSafe -Object $Item -Name 'ConversationTopic' -Default ''
    $conversationId = Get-PropSafe -Object $Item -Name 'ConversationID' -Default ''
    $sortTime = $receivedTime
    if (Test-MissingDate $sortTime) { $sortTime = $sentOn }
    if (Test-MissingDate $sortTime) { $sortTime = $creationTime }
    if (Test-MissingDate $sortTime) { $sortTime = $null }

    $senderDisplay = ConvertTo-NormalizedPersonName $senderName
    if ([string]::IsNullOrWhiteSpace($senderDisplay)) { $senderDisplay = ConvertTo-NormalizedPersonName $senderEmail }
    if ([string]::IsNullOrWhiteSpace($senderDisplay)) { $senderDisplay = '(unknown sender)' }

    $participants = @(Get-ParticipantName -SenderName $senderName -SenderEmail $senderEmail -To $to -Cc $cc)
    $participantKey = Get-ParticipantKey -Participants $participants
    $conversationSubject = if ([string]::IsNullOrWhiteSpace([string]$subject)) { '(no subject)' } else { [string]$subject }

    $conversationKey = "$participantKey`n$conversationSubject`n$FolderPath"

    [pscustomobject]@{
        SortTime             = $sortTime
        FolderPath           = $FolderPath
        Subject              = $subject
        MessageClass         = $messageClass
        SenderName           = $senderName
        SenderEmail          = $senderEmail
        SenderDisplay        = $senderDisplay
        To                   = $to
        Cc                   = $cc
        Participants         = $participants
        ParticipantsKey      = $participantKey
        ConversationKey      = $conversationKey
        ConversationTitle    = $conversationSubject
        SentOn               = $sentOn
        ReceivedTime         = $receivedTime
        CreationTime         = $creationTime
        EntryId              = $entryId
        ConversationTopic    = $conversationTopic
        ConversationId       = $conversationId
        BodyText             = [string]$body
        AttachmentsHtml      = Get-AttachmentSummaryHtml -Item $Item
    }
}

function Get-RecurrenceSummary {
    param([AllowNull()][object]$Item)

    if (-not [bool](Get-PropSafe -Object $Item -Name 'IsRecurring' -Default $false)) { return '' }

    $pattern = $null
    try {
        if ($null -eq $Item.PSObject.Methods['GetRecurrencePattern']) { return '' }
        $pattern = $Item.GetRecurrencePattern()
        if ($null -eq $pattern) { return '' }

        $typeValue = [int](Get-PropSafe -Object $pattern -Name 'RecurrenceType' -Default -1)
        $typeLabel = switch ($typeValue) {
            0 { 'Daily' }
            1 { 'Weekly' }
            2 { 'Monthly' }
            3 { 'MonthlyNth' }
            5 { 'Yearly' }
            6 { 'YearlyNth' }
            default { "Type$typeValue" }
        }

        $parts = New-Object System.Collections.Generic.List[string]
        [void]$parts.Add($typeLabel)

        $interval = Get-PropSafe -Object $pattern -Name 'Interval' -Default $null
        if ($null -ne $interval -and [int]$interval -gt 0) {
            [void]$parts.Add("every $interval")
        }

        $patternStart = Get-PropSafe -Object $pattern -Name 'PatternStartDate' -Default $null
        if (-not (Test-MissingDate $patternStart)) {
            [void]$parts.Add(("starting {0}" -f ([datetime]$patternStart).ToString('yyyy-MM-dd')))
        }

        if ([bool](Get-PropSafe -Object $pattern -Name 'NoEndDate' -Default $false)) {
            [void]$parts.Add('no end date')
        }
        else {
            $patternEnd = Get-PropSafe -Object $pattern -Name 'PatternEndDate' -Default $null
            if (-not (Test-MissingDate $patternEnd)) {
                [void]$parts.Add(("ending {0}" -f ([datetime]$patternEnd).ToString('yyyy-MM-dd')))
            }
        }

        $occurrences = Get-PropSafe -Object $pattern -Name 'Occurrences' -Default $null
        if ($null -ne $occurrences -and [int]$occurrences -gt 0) {
            $occurrenceLabel = if ([int]$occurrences -eq 1) { 'occurrence' } else { 'occurrences' }
            [void]$parts.Add("$occurrences $occurrenceLabel")
        }

        return ($parts -join '; ')
    }
    catch {
        Write-ReportLog "Could not read recurrence details on calendar item. $($_.Exception.Message)" 'WARN'
        return ''
    }
    finally {
        Close-ComObjectSafe $pattern
    }
}

function Get-CalendarRecord {
    param(
        [Parameter(Mandatory = $true)] [object]$Item,
        [Parameter(Mandatory = $true)] [string]$FolderPath
    )

    $startTime = Get-PropSafe -Object $Item -Name 'Start' -Default $null
    if (Test-MissingDate $startTime) { $startTime = $null }
    $endTime = Get-PropSafe -Object $Item -Name 'End' -Default $null
    if (Test-MissingDate $endTime) { $endTime = $null }
    $creationTime = Get-PropSafe -Object $Item -Name 'CreationTime' -Default $null
    if (Test-MissingDate $creationTime) { $creationTime = $null }
    $lastModificationTime = Get-PropSafe -Object $Item -Name 'LastModificationTime' -Default $null
    if (Test-MissingDate $lastModificationTime) { $lastModificationTime = $null }
    $messageClass = [string](Get-PropSafe -Object $Item -Name 'MessageClass' -Default '')
    $sortTime = $startTime
    if (Test-MissingDate $sortTime) { $sortTime = $creationTime }
    if (Test-MissingDate $sortTime) { $sortTime = $lastModificationTime }
    if (Test-MissingDate $sortTime) { $sortTime = $null }

    $itemType = if ($messageClass -match '(?i)^IPM\.Schedule\.Meeting') { 'Meeting' } else { 'Appointment' }
    $bodyText = [string](Get-PropSafe -Object $Item -Name 'Body' -Default '')

    [pscustomobject]@{
        SortTime             = $sortTime
        StartTime            = $startTime
        EndTime              = $endTime
        Subject              = [string](Get-PropSafe -Object $Item -Name 'Subject' -Default '')
        ItemType             = $itemType
        AllDayEvent          = [bool](Get-PropSafe -Object $Item -Name 'AllDayEvent' -Default $false)
        Location             = [string](Get-PropSafe -Object $Item -Name 'Location' -Default '')
        Organizer            = [string](Get-PropSafe -Object $Item -Name 'Organizer' -Default '')
        RequiredAttendees    = [string](Get-PropSafe -Object $Item -Name 'RequiredAttendees' -Default '')
        OptionalAttendees    = [string](Get-PropSafe -Object $Item -Name 'OptionalAttendees' -Default '')
        IsRecurring          = [bool](Get-PropSafe -Object $Item -Name 'IsRecurring' -Default $false)
        RecurrenceSummary    = Get-RecurrenceSummary -Item $Item
        Categories           = [string](Get-PropSafe -Object $Item -Name 'Categories' -Default '')
        Sensitivity          = Get-PropSafe -Object $Item -Name 'Sensitivity' -Default $null
        FolderPath           = $FolderPath
        BodyText             = $bodyText
        Notes                = $bodyText
        MessageClass         = $messageClass
        EntryId              = [string](Get-PropSafe -Object $Item -Name 'EntryID' -Default '')
        CreationTime         = $creationTime
        LastModificationTime = $lastModificationTime
        Attachments          = @(Get-AttachmentMetadata -Item $Item)
    }
}

function Get-DistributionListMembers {
    param([AllowNull()][object]$Item)

    $members = New-Object System.Collections.Generic.List[string]
    $memberCount = [int](Get-PropSafe -Object $Item -Name 'MemberCount' -Default 0)
    if ($memberCount -le 0) { return @() }
    if ($null -eq $Item.PSObject.Methods['GetMember']) { return @() }

    for ($i = 1; $i -le $memberCount; $i++) {
        $member = $null
        try {
            $member = $Item.GetMember($i)
            $memberName = [string](Get-PropSafe -Object $member -Name 'Name' -Default '')
            if ([string]::IsNullOrWhiteSpace($memberName)) {
                $memberName = [string](Get-PropSafe -Object $member -Name 'Address' -Default '')
            }
            if (-not [string]::IsNullOrWhiteSpace($memberName)) {
                [void]$members.Add($memberName)
            }
        }
        catch {
            Write-ReportLog "Could not read distribution list member $i on contact item. $($_.Exception.Message)" 'WARN'
        }
        finally {
            Close-ComObjectSafe $member
        }
    }

    return $members.ToArray()
}

function Get-ContactRecord {
    param(
        [Parameter(Mandatory = $true)] [object]$Item,
        [Parameter(Mandatory = $true)] [string]$FolderPath
    )

    $firstName = [string](Get-PropSafe -Object $Item -Name 'FirstName' -Default '')
    $middleName = [string](Get-PropSafe -Object $Item -Name 'MiddleName' -Default '')
    $lastName = [string](Get-PropSafe -Object $Item -Name 'LastName' -Default '')
    $messageClass = [string](Get-PropSafe -Object $Item -Name 'MessageClass' -Default '')
    $isDistributionList = $messageClass -match '(?i)^IPM\.DistList'
    $dlName = [string](Get-PropSafe -Object $Item -Name 'DLName' -Default '')
    $fullName = [string](Get-PropSafe -Object $Item -Name 'FullName' -Default '')
    if ($isDistributionList -and [string]::IsNullOrWhiteSpace($fullName)) {
        $fullName = $dlName
    }
    if ([string]::IsNullOrWhiteSpace($fullName)) {
        $fullName = (@($firstName, $middleName, $lastName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' '
    }
    $displayName = [string](Get-PropSafe -Object $Item -Name 'FileAs' -Default '')
    if ($isDistributionList -and [string]::IsNullOrWhiteSpace($displayName)) { $displayName = $dlName }
    if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $fullName }
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = [string](Get-PropSafe -Object $Item -Name 'CompanyName' -Default '')
    }

    $birthday = Get-PropSafe -Object $Item -Name 'Birthday' -Default $null
    if (Test-MissingDate $birthday) { $birthday = $null }
    $anniversary = Get-PropSafe -Object $Item -Name 'Anniversary' -Default $null
    if (Test-MissingDate $anniversary) { $anniversary = $null }
    $creationTime = Get-PropSafe -Object $Item -Name 'CreationTime' -Default $null
    if (Test-MissingDate $creationTime) { $creationTime = $null }
    $lastModificationTime = Get-PropSafe -Object $Item -Name 'LastModificationTime' -Default $null
    if (Test-MissingDate $lastModificationTime) { $lastModificationTime = $null }
    $bodyText = [string](Get-PropSafe -Object $Item -Name 'Body' -Default '')

    [pscustomobject]@{
        DisplayName             = $displayName
        FullName                = $fullName
        FirstName               = $firstName
        MiddleName              = $middleName
        LastName                = $lastName
        CompanyName             = [string](Get-PropSafe -Object $Item -Name 'CompanyName' -Default '')
        JobTitle                = [string](Get-PropSafe -Object $Item -Name 'JobTitle' -Default '')
        Department              = [string](Get-PropSafe -Object $Item -Name 'Department' -Default '')
        Email1                  = [string](Get-PropSafe -Object $Item -Name 'Email1Address' -Default '')
        Email2                  = [string](Get-PropSafe -Object $Item -Name 'Email2Address' -Default '')
        Email3                  = [string](Get-PropSafe -Object $Item -Name 'Email3Address' -Default '')
        BusinessPhone           = [string](Get-PropSafe -Object $Item -Name 'BusinessTelephoneNumber' -Default '')
        HomePhone               = [string](Get-PropSafe -Object $Item -Name 'HomeTelephoneNumber' -Default '')
        MobilePhone             = [string](Get-PropSafe -Object $Item -Name 'MobileTelephoneNumber' -Default '')
        OtherPhone              = [string](Get-PropSafe -Object $Item -Name 'OtherTelephoneNumber' -Default '')
        BusinessAddress         = [string](Get-PropSafe -Object $Item -Name 'BusinessAddress' -Default '')
        HomeAddress             = [string](Get-PropSafe -Object $Item -Name 'HomeAddress' -Default '')
        OtherAddress            = [string](Get-PropSafe -Object $Item -Name 'OtherAddress' -Default '')
        WebPage                 = [string](Get-PropSafe -Object $Item -Name 'WebPage' -Default '')
        Birthday                = $birthday
        Anniversary             = $anniversary
        Categories              = [string](Get-PropSafe -Object $Item -Name 'Categories' -Default '')
        DistributionListMembers = @(Get-DistributionListMembers -Item $Item)
        FolderPath              = $FolderPath
        BodyText                = $bodyText
        Notes                   = $bodyText
        MessageClass            = $messageClass
        EntryId                 = [string](Get-PropSafe -Object $Item -Name 'EntryID' -Default '')
        CreationTime            = $creationTime
        LastModificationTime    = $lastModificationTime
        Attachments             = @(Get-AttachmentMetadata -Item $Item)
    }
}

function Read-OutlookFolder {
    param(
        [Parameter(Mandatory = $true)] [object]$Folder,
        [Parameter(Mandatory = $true)] [string]$FolderPath,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]]$TeamsRecords,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]]$EmailRecords,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]]$CalendarRecords,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]]$ContactsRecords
    )

    $script:Stats.FoldersScanned++
    Write-ReportLog "Scanning folder: $FolderPath"
    $loggedNonReportSkip = $false
    $items = $null
    try {
        $items = $Folder.Items
        $itemCount = [int]$items.Count
        Write-ReportLog "Folder item count: $itemCount ($FolderPath)"

        for ($i = 1; $i -le $itemCount; $i++) {
            $item = $null
            try {
                $script:Stats.ItemsAttempted++
                # Retry a transient Outlook-busy rejection on the per-item fetch instead of
                # letting the ItemReadFailures catch swallow the message. Genuine (non-RPC)
                # read failures still throw here and are handled by the catch below.
                $item = Invoke-OutlookComOperation -Operation "reading item $i in $FolderPath" -ScriptBlock { $items.Item($i) }
                $messageClass = Get-PropSafe -Object $item -Name 'MessageClass' -Default ''
                $body = Get-PropSafe -Object $item -Name 'Body' -Default $null
                $subject = Get-PropSafe -Object $item -Name 'Subject' -Default $null
                $bucket = Get-ItemReportBucket -FolderPath $FolderPath -MessageClass $messageClass
                $isMessageLike = ($messageClass -or $body -or $subject -or $bucket -in @('Calendar', 'Contacts'))

                # Purview Teams messages are usually mail-like items in an Exchange PST. Export any
                # item that has normal message properties instead of requiring one exact MessageClass.
                if ($isMessageLike) {
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
                    elseif ($bucket -eq 'Calendar' -and $CalendarReport) {
                        [void]$CalendarRecords.Add((Get-CalendarRecord -Item $item -FolderPath $FolderPath))
                        $script:Stats.CalendarItemsExported++
                        $script:Stats.ItemsExported++
                    }
                    elseif ($bucket -eq 'Contacts' -and $ContactsReport) {
                        [void]$ContactsRecords.Add((Get-ContactRecord -Item $item -FolderPath $FolderPath))
                        $script:Stats.ContactsItemsExported++
                        $script:Stats.ItemsExported++
                    }
                    else {
                        if (-not $loggedNonReportSkip -and $bucket -eq 'Skip') {
                            Write-ReportLog "Skipping non-report item in folder: $FolderPath (class=$messageClass)"
                            $loggedNonReportSkip = $true
                        }
                        $script:Stats.ItemsSkipped++
                    }
                }
                else {
                    $script:Stats.ItemsSkipped++
                }

                if ($script:Stats.ItemsExported -gt 0 -and $script:Stats.ItemsExported % $LogEvery -eq 0) {
                    Write-ConversionProgress -FolderPath $FolderPath
                    Write-ReportLog "Collected $($script:Stats.ItemsExported) message-like items so far."
                }
            }
            catch {
                $script:Stats.ItemReadFailures++
                Write-ReportLog "Could not read item $i in $FolderPath. $($_.Exception.Message)" 'WARN'
            }
            finally {
                Close-ComObjectSafe $item
            }
        }
    }
    finally {
        Close-ComObjectSafe $items
    }

    $subFolders = $null
    try {
        $subFolders = $Folder.Folders
        $subFolderCount = [int]$subFolders.Count
        for ($j = 1; $j -le $subFolderCount; $j++) {
            $child = $null
            try {
                $child = $subFolders.Item($j)
                $childName = Get-PropSafe -Object $child -Name 'Name' -Default "Folder$j"
                Read-OutlookFolder -Folder $child -FolderPath "$FolderPath\$childName" -TeamsRecords $TeamsRecords -EmailRecords $EmailRecords -CalendarRecords $CalendarRecords -ContactsRecords $ContactsRecords
            }
            catch {
                $script:Stats.SubfolderScanFailures++
                Write-ReportLog "Could not scan child folder under $FolderPath. $($_.Exception.Message)" 'WARN'
            }
            finally {
                Close-ComObjectSafe $child
            }
        }
    }
    finally {
        Close-ComObjectSafe $subFolders
    }
}

function Find-StoreRootForPst {
    param(
        [Parameter(Mandatory = $true)] [object]$Namespace,
        [Parameter(Mandatory = $true)] [string]$TargetPath
    )

    $stores = $Namespace.Stores
    try {
        $storeCount = [int]$stores.Count
        for ($i = 1; $i -le $storeCount; $i++) {
            $store = $null
            try {
                $store = $stores.Item($i)
                $filePath = Get-PropSafe -Object $store -Name 'FilePath' -Default ''
                if ($filePath -and ([IO.Path]::GetFullPath($filePath) -ieq [IO.Path]::GetFullPath($TargetPath))) {
                    return $store.GetRootFolder()
                }
            }
            finally { Close-ComObjectSafe $store }
        }
    }
    finally { Close-ComObjectSafe $stores }
    return $null
}

function Wait-StoreRootForPst {
    param(
        [Parameter(Mandatory = $true)] [object]$Namespace,
        [Parameter(Mandatory = $true)] [string]$TargetPath,
        [int]$TimeoutSeconds = 30
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $root = Find-StoreRootForPst -Namespace $Namespace -TargetPath $TargetPath
        if ($null -ne $root) { return $root }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    return $null
}

function Test-PstStoreAttached {
    param(
        [Parameter(Mandatory = $true)] [object]$Namespace,
        [Parameter(Mandatory = $true)] [string]$TargetPath
    )

    # Outlook may already be dead (RPC server unavailable). Never throw from here —
    # a throw in the conversion finally-block would flip a successful report write to exit 1.
    try {
        $root = Find-StoreRootForPst -Namespace $Namespace -TargetPath $TargetPath
        if ($null -ne $root) { Close-ComObjectSafe $root; return $true }
        return $false
    }
    catch {
        Write-ReportLog "Could not verify whether PST is still attached. $($_.Exception.Message)" 'WARN'
        return $true
    }
}

function Remove-PstStoreFromOutlook {
    param(
        [Parameter(Mandatory = $true)] [object]$Namespace,
        [Parameter(Mandatory = $true)] [string]$TargetPath,
        [AllowNull()] [object]$RootFolder
    )

    $targetRoot = $RootFolder
    $discovered = $false
    if ($null -eq $targetRoot) {
        $targetRoot = Find-StoreRootForPst -Namespace $Namespace -TargetPath $TargetPath
        $discovered = $null -ne $targetRoot
    }
    if ($null -eq $targetRoot) { return $false }

    Invoke-OutlookComOperation -Operation 'detaching the PST from Outlook' -ScriptBlock { $Namespace.RemoveStore($targetRoot) } | Out-Null
    if ($discovered) { Close-ComObjectSafe $targetRoot }
    return $true
}

function Confirm-OutlookPstCleanup {
    param(
        [Parameter(Mandatory = $true)] [object]$Namespace,
        [Parameter(Mandatory = $true)] [string]$TargetPath,
        [AllowNull()] [object]$RootFolder
    )

    [void](Remove-PstStoreFromOutlook -Namespace $Namespace -TargetPath $TargetPath -RootFolder $RootFolder)
    if (Test-PstStoreAttached -Namespace $Namespace -TargetPath $TargetPath) {
        Write-ReportLog 'PST still attached after detach; retrying once.' 'WARN'
        [void](Remove-PstStoreFromOutlook -Namespace $Namespace -TargetPath $TargetPath -RootFolder $null)
    }
    if (Test-PstStoreAttached -Namespace $Namespace -TargetPath $TargetPath) {
        Write-ReportLog "PST is still attached to your Outlook profile: $TargetPath. Remove it manually in Outlook (File -> Account Settings -> Data Files), then verify your profile is unchanged." 'WARN'
        return $false
    }

    Write-ReportLog 'Verified PST detached from Outlook profile; Outlook left as before this conversion.'
    return $true
}

function Assert-StaForOutlookCom {
    if ($UseSampleData) { return }
    $state = [System.Threading.Thread]::CurrentThread.ApartmentState
    if ($state -ne [System.Threading.ApartmentState]::STA) {
        $command = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh -Sta -File' } else { 'powershell.exe -Sta -File' }
        throw "Outlook COM automation should be run in STA mode. Current apartment state is $state. Re-run with: $command `"$PSCommandPath`" -PstPath `"<path-to-pst>`""
    }
}

function Get-ReportCss {
    return @'
:root { --bg: #f4f6fb; --card: #ffffff; --text: #1f2937; --muted: #5b6472; --line: #d9e0ea; --accent: #315fbd; }
* { box-sizing: border-box; }
body { font-family: "Segoe UI", Arial, sans-serif; margin: 0; color: var(--text); background: var(--bg); line-height: 1.45; }
.page { max-width: 1600px; margin: 0 auto; padding: 16px; }
.hero { position: relative; background: linear-gradient(135deg, #213b78, #5c7dde); color: white; border-radius: 18px; padding: 14px 22px; box-shadow: 0 10px 28px rgba(24, 44, 92, .22); }
.hero h1 { margin: 0 0 6px 0; font-size: 1.65rem; }
.hero p { margin: 4px 0; opacity: .95; }
.hero-credit { position: absolute; right: 18px; bottom: 8px; font-size: .78rem; opacity: .85; font-weight: 600; }
.summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 10px; margin: 10px 0; }
.summary-card, .filter-panel, .folder-summary, .conversation, .conversation-toolbar { background: var(--card); border: 1px solid var(--line); border-radius: 16px; box-shadow: 0 2px 8px rgba(22, 34, 51, .06); }
.summary-card { padding: 7px 12px; }
.summary-card .label { color: var(--muted); font-size: .82rem; }
.summary-card .value { font-weight: 700; font-size: .98rem; overflow-wrap: anywhere; }
.review-layout { --left-column-width: 340px; display: grid; grid-template-columns: minmax(180px, var(--left-column-width)) 10px minmax(0, 1fr); gap: 10px; align-items: start; }
.filter-panel { padding: 14px; position: sticky; top: 8px; z-index: 5; max-height: calc(100vh - 16px); overflow: auto; min-width: 170px; }
.resize-handle { position: sticky; top: 8px; height: calc(100vh - 16px); border-radius: 999px; background: linear-gradient(90deg, transparent 0, transparent 3px, #b7c3d6 3px, #b7c3d6 7px, transparent 7px); cursor: col-resize; touch-action: none; }
.resize-handle:hover, .resize-handle.dragging { background: linear-gradient(90deg, transparent 0, transparent 2px, #315fbd 2px, #315fbd 8px, transparent 8px); }
.resize-handle:focus { outline: 3px solid rgba(49, 95, 189, .35); outline-offset: 2px; }
.conversation-pane { max-height: calc(100vh - 16px); overflow: auto; padding-right: 6px; }
.filter-title { display: block; margin-bottom: 10px; }
.filter-title h2 { margin: 0 0 6px; font-size: 1.18rem; }
.filter-help { color: var(--muted); margin: 4px 0 10px; font-size: .93rem; }
.controls { display: grid; grid-template-columns: 1fr; gap: 9px; margin-bottom: 10px; }
.controls input[type="text"], .controls input[type="date"], .controls select { width: 100%; padding: 10px 11px; border: 1px solid #b9c4d3; border-radius: 10px; font-size: .98rem; background: white; }
.conversation-toolbar input[type="date"], .conversation-toolbar select { width: 100%; padding: 7px 10px; border: 1px solid #b9c4d3; border-radius: 10px; font-size: .9rem; background: white; }
.conversation-toolbar { display: grid; grid-template-columns: repeat(3, minmax(160px, 1fr)); gap: 10px; padding: 6px 14px; margin: 0 0 10px 0; position: sticky; top: 0; z-index: 4; }
.toolbar-field { display: grid; gap: 3px; }
.toolbar-field label { font-weight: 800; color: #334155; font-size: .82rem; }
.filter-toggle { display: flex; gap: 8px; align-items: flex-start; padding: 9px 10px; border: 1px solid var(--line); border-radius: 10px; background: #fbfcff; font-size: .92rem; font-weight: 700; color: #334155; }
.filter-toggle input { margin-top: 3px; }
.people-box { max-height: 36vh; overflow: auto; border: 1px solid var(--line); border-radius: 12px; padding: 8px; background: #fbfcff; display: grid; grid-template-columns: 1fr; gap: 5px; }
.people-heading { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin: 14px 0 8px; }
.people-heading h3 { margin: 0; }
.people-heading button { padding: 7px 10px; font-size: .88rem; }
.person-option { display: flex; gap: 8px; align-items: center; padding: 6px 7px; border-radius: 9px; cursor: pointer; }
.person-option:hover { background: #eef3ff; }
.other-names { margin-top: 10px; border: 1px solid var(--line); border-radius: 12px; background: #fbfcff; }
.other-names summary { cursor: pointer; padding: 10px; font-weight: 800; color: #334155; }
.other-names .people-box { border: 0; border-top: 1px solid var(--line); border-radius: 0 0 12px 12px; max-height: 20vh; }
button { border: 0; border-radius: 10px; padding: 9px 10px; background: var(--accent); color: white; font-weight: 700; cursor: pointer; white-space: nowrap; }
button.secondary { background: #e8edf7; color: #243047; }
.result-count { color: var(--accent); font-weight: 800; display: block; }
.notice { background: #fff8df; border-left: 6px solid #f5bc2f; padding: 10px; border-radius: 10px; margin-top: 10px; font-size: .9rem; }
.folder-summary { padding: 10px; margin-top: 10px; }
.folder-summary summary { cursor: pointer; font-weight: 800; }
table { border-collapse: collapse; width: 100%; margin-top: 8px; }
th, td { border: 1px solid var(--line); padding: 8px; vertical-align: top; }
th { background: #edf2fb; text-align: left; }
.conversation { margin: 0 0 16px 0; overflow: hidden; }
.conversation[hidden] { display: none; }
.conversation-header { display: flex; justify-content: space-between; gap: 14px; padding: 11px 18px; background: #eef3ff; border-bottom: 1px solid var(--line); }
.conversation-header h2 { margin: 0 0 2px; font-size: 1.04rem; }
.conversation-people { color: var(--muted); font-size: .9rem; }
.conversation-count { white-space: nowrap; color: var(--accent); font-weight: 800; font-size: .92rem; }
.conversation-messages { padding: 18px; }
.message-card { border-left: 7px solid #5f6f89; background: #fbfcff; border-radius: 14px; margin: 0 0 14px 0; padding: 13px 15px; box-shadow: 0 1px 4px rgba(20, 33, 61, .06); }
.message-card[hidden] { display: none; }
.speaker-row { display: flex; justify-content: space-between; gap: 12px; align-items: flex-start; margin-bottom: 8px; }
.speaker-block { display: flex; flex-direction: column; gap: 2px; min-width: 0; max-width: 100%; }
.speaker-name { font-weight: 800; font-size: 1.04rem; white-space: nowrap; }
.speaker-context { display: block; color: var(--muted); font-size: .78rem; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0; }
.message-time { color: var(--muted); font-size: .9rem; white-space: nowrap; }
.message-body { background: white; border-radius: 10px; padding: 12px; border: 1px solid #e5eaf2; overflow-wrap: anywhere; }
.message-details { margin-top: 10px; color: var(--muted); font-size: .9rem; }
.message-details summary { cursor: pointer; font-weight: 700; }
.attachments { margin-top: 10px; font-size: .9rem; }
.wrap { overflow-wrap: anywhere; word-break: break-word; }
.sender-blue { border-left-color: #2f6fec; } .sender-blue .speaker-name { color: #2459bf; }
.sender-green { border-left-color: #16875d; } .sender-green .speaker-name { color: #116b4a; }
.sender-purple { border-left-color: #7c3aed; } .sender-purple .speaker-name { color: #6427c5; }
.sender-orange { border-left-color: #e26a00; } .sender-orange .speaker-name { color: #b45100; }
.sender-red { border-left-color: #dc3545; } .sender-red .speaker-name { color: #b02a37; }
.sender-teal { border-left-color: #008c95; } .sender-teal .speaker-name { color: #006f76; }
.sender-pink { border-left-color: #c026d3; } .sender-pink .speaker-name { color: #9b1fac; }
.sender-brown { border-left-color: #8b5e34; } .sender-brown .speaker-name { color: #6f4b2a; }
.sender-slate { border-left-color: #64748b; } .sender-slate .speaker-name { color: #475569; }
.sender-indigo { border-left-color: #4f46e5; } .sender-indigo .speaker-name { color: #3730a3; }
.footer { color: var(--muted); font-size: .9rem; margin: 28px 0 8px; }
@media (max-width: 900px) { .page { padding: 12px; } .review-layout { grid-template-columns: 1fr; } .filter-panel { position: static; max-height: none; } .resize-handle { display: none; } .conversation-pane { max-height: none; overflow: visible; } .conversation-toolbar { grid-template-columns: 1fr; position: static; } .conversation-header, .speaker-row { display: block; } .speaker-block { max-width: 100%; } .message-time { display: block; margin-top: 2px; } }
'@
}

function Get-ReportScript {
    return @'
(function () {
  const checks = Array.from(document.querySelectorAll('.person-check'));
  const personSearch = document.getElementById('personSearch');
  const messageSearch = document.getElementById('messageSearch');
  const participantMatchMode = document.getElementById('participantMatchMode');
  const startDateFilter = document.getElementById('startDateFilter');
  const endDateFilter = document.getElementById('endDateFilter');
  const sortOrder = document.getElementById('sortOrder');
  const conversationList = document.getElementById('conversationList');
  const conversations = Array.from(document.querySelectorAll('.conversation'));
  const resultCount = document.getElementById('resultCount');
  const selectAllPeople = document.getElementById('selectAllPeople');
  const selectAllOther = document.getElementById('selectAllOther');
  const clearAll = document.getElementById('clearAll');
  const layout = document.querySelector('.review-layout');
  const resizeHandle = document.getElementById('resizeHandle');

  conversations.forEach(conversation => {
    conversation._searchText = (conversation.textContent || '').replace(/\s+/g, ' ').toLowerCase();
  });

  function selectedPeople() { return checks.filter(c => c.checked).map(c => c.value); }
  function listFromDataset(conversation, name) { return (conversation.dataset[name] || '').split('||').map(v => v.trim()).filter(Boolean); }
  function participantList(conversation) { return listFromDataset(conversation, 'participants'); }
  function messageSender(message) { return (message.dataset.sender || '').trim(); }
  function messageDate(message) { return (message.dataset.date || '').trim(); }
  function messageInDateRange(message, startDate, endDate) {
    const date = messageDate(message);
    if (!date) return true;
    if (startDate && date < startDate) return false;
    if (endDate && date > endDate) return false;
    return true;
  }
  function matchesSelectedSender(message, selected) { return selected.length === 0 || selected.includes(messageSender(message)); }
  function containsAll(actual, selected) { return selected.every(p => actual.includes(p)); }
  function sameSet(actual, selected) { return containsAll(actual, selected) && actual.length === selected.length; }
  function matchesParticipantMode(conversation, selected, mode) {
    if (selected.length === 0) return true;
    const allParticipants = participantList(conversation);
    const personParticipants = listFromDataset(conversation, 'personParticipants');

    if (mode === 'involving') {
      return containsAll(allParticipants, selected);
    }

    if (mode === 'exactAll') {
      return sameSet(allParticipants, selected);
    }

    if (mode === 'exactPeopleOnly') {
      const selectedPeopleOnly = selected.filter(p => personParticipants.includes(p));
      return containsAll(allParticipants, selected) && sameSet(personParticipants, selectedPeopleOnly);
    }

    return true;
  }

  function sortConversations() {
    if (!conversationList || !sortOrder) return;
    const order = sortOrder.value || 'newestFirst';
    const sorted = conversations.slice().sort((a, b) => {
      const aTime = Date.parse(a.dataset.sortTime || '') || 0;
      const bTime = Date.parse(b.dataset.sortTime || '') || 0;
      return order === 'oldestFirst' ? aTime - bTime : bTime - aTime;
    });
    sorted.forEach(conversation => conversationList.appendChild(conversation));
  }

  function applyFilters() {
    const selected = selectedPeople();
    const mode = participantMatchMode ? participantMatchMode.value : 'messagesFromSelected';
    const startDate = startDateFilter ? startDateFilter.value : '';
    const endDate = endDateFilter ? endDateFilter.value : '';
    const text = (messageSearch.value || '').trim().toLowerCase();
    let visibleConversations = 0;
    let visibleMessages = 0;

    conversations.forEach(conversation => {
      const conversationTextOk = !text || (conversation._searchText || '').includes(text);
      const messages = Array.from(conversation.querySelectorAll('.message-card'));
      let conversationVisibleMessages = 0;

      if (mode === 'messagesFromSelected') {
        messages.forEach(message => {
          const showMessage = conversationTextOk && messageInDateRange(message, startDate, endDate) && matchesSelectedSender(message, selected);
          message.hidden = !showMessage;
          if (showMessage) conversationVisibleMessages += 1;
        });
      } else {
        const showConversation = conversationTextOk && matchesParticipantMode(conversation, selected, mode);
        messages.forEach(message => {
          const showMessage = showConversation && messageInDateRange(message, startDate, endDate);
          message.hidden = !showMessage;
          if (showMessage) conversationVisibleMessages += 1;
        });
      }

      const showConversation = conversationVisibleMessages > 0;
      const countEl = conversation.querySelector('.conversation-count');
      if (countEl) countEl.textContent = (selected.length > 0 || text) ? (conversationVisibleMessages + ' of ' + messages.length + ' messages') : (messages.length + ' messages');
      conversation.hidden = !showConversation;
      if (showConversation) { visibleConversations += 1; visibleMessages += conversationVisibleMessages; }
    });

    sortConversations();
    resultCount.textContent = visibleConversations + ' conversations / ' + visibleMessages + ' messages shown';
  }

  function filterPersonList() {
    const text = (personSearch.value || '').trim().toLowerCase();
    checks.forEach(check => {
      const label = check.closest('.person-option');
      label.style.display = !text || check.value.toLowerCase().includes(text) ? '' : 'none';
    });
  }

  function safeGetLocalStorage(key) {
    try { return localStorage.getItem(key); } catch (_) { return null; }
  }
  function safeSetLocalStorage(key, value) {
    try { localStorage.setItem(key, value); } catch (_) { }
  }

  function setupColumnResize() {
    if (!layout || !resizeHandle) return;
    const savedWidth = safeGetLocalStorage('purviewTeamsReport.leftColumnWidth');
    if (savedWidth) layout.style.setProperty('--left-column-width', savedWidth);
    let dragging = false;
    function setWidthFromPointer(clientX) {
      const rect = layout.getBoundingClientRect();
      const min = 180;
      const max = Math.max(min, Math.min(720, rect.width - 320));
      const width = Math.max(min, Math.min(max, clientX - rect.left));
      const value = Math.round(width) + 'px';
      layout.style.setProperty('--left-column-width', value);
      safeSetLocalStorage('purviewTeamsReport.leftColumnWidth', value);
    }
    resizeHandle.addEventListener('pointerdown', event => { dragging = true; resizeHandle.classList.add('dragging'); resizeHandle.setPointerCapture(event.pointerId); event.preventDefault(); });
    resizeHandle.addEventListener('pointermove', event => { if (dragging) setWidthFromPointer(event.clientX); });
    function stopDragging(event) { if (!dragging) return; dragging = false; resizeHandle.classList.remove('dragging'); try { resizeHandle.releasePointerCapture(event.pointerId); } catch (_) { } }
    resizeHandle.addEventListener('pointerup', stopDragging);
    resizeHandle.addEventListener('pointercancel', stopDragging);
    resizeHandle.addEventListener('keydown', event => {
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      const current = parseInt(getComputedStyle(layout).getPropertyValue('--left-column-width'), 10) || 340;
      const delta = event.key === 'ArrowRight' ? 30 : -30;
      const value = Math.max(180, Math.min(720, current + delta)) + 'px';
      layout.style.setProperty('--left-column-width', value);
      safeSetLocalStorage('purviewTeamsReport.leftColumnWidth', value);
      event.preventDefault();
    });
  }

  checks.forEach(c => c.addEventListener('change', applyFilters));
  personSearch.addEventListener('input', filterPersonList);
  messageSearch.addEventListener('input', applyFilters);
  if (participantMatchMode) participantMatchMode.addEventListener('change', applyFilters);
  if (startDateFilter) startDateFilter.addEventListener('change', applyFilters);
  if (endDateFilter) endDateFilter.addEventListener('change', applyFilters);
  if (sortOrder) sortOrder.addEventListener('change', applyFilters);

  function setGroupChecked(selectAllCheckbox, groupId) {
    if (!selectAllCheckbox) return;
    const group = document.getElementById(groupId);
    if (!group) return;
    group.querySelectorAll('.person-check').forEach(c => { c.checked = selectAllCheckbox.checked; });
    applyFilters();
  }

  if (selectAllPeople) selectAllPeople.addEventListener('change', () => setGroupChecked(selectAllPeople, 'peopleBox'));
  if (selectAllOther) selectAllOther.addEventListener('change', () => setGroupChecked(selectAllOther, 'otherPeopleBox'));
  clearAll.addEventListener('click', () => {
    checks.forEach(c => c.checked = false);
    if (selectAllPeople) selectAllPeople.checked = false;
    if (selectAllOther) selectAllOther.checked = false;
    if (participantMatchMode) participantMatchMode.value = 'messagesFromSelected';
    if (startDateFilter) startDateFilter.value = '';
    if (endDateFilter) endDateFilter.value = '';
    if (sortOrder) sortOrder.value = 'newestFirst';
    personSearch.value = '';
    messageSearch.value = '';
    filterPersonList();
    applyFilters();
  });

  setupColumnResize();
  filterPersonList();
  applyFilters();
})();
'@
}

function Format-ParticipantOptionHtml {
    param(
        [Parameter(Mandatory = $true)] [string]$Participant,
        [AllowNull()][string[]]$DefaultParticipants
    )
    $checked = if ($DefaultParticipants -contains $Participant) { ' checked="checked"' } else { '' }
    return "<label class='person-option'><input type='checkbox' class='person-check' value='$(ConvertTo-HtmlEncodedText $Participant)'$checked/> <span>$(ConvertTo-HtmlEncodedText $Participant)</span></label>"
}

function Write-ReportHeader {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)]$PstItem,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SortedRecords,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$AllParticipants,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$OtherParticipants,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ParticipantOptions,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$OtherParticipantOptions,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$FolderSummaryRows
    )

    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss K'
    $css = Get-ReportCss
    $Writer.WriteLine(@"
<!doctype html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Purview Teams PST Conversation Report - $(ConvertTo-HtmlEncodedText $PstItem.Name)</title>
<style>
$css
</style>
</head>
<body>
<div class='page'>
  <header class='hero'>
    <h1>Microsoft Purview eDiscovery Teams Conversation Report</h1>
    <p>Use the simple filter below to show conversations between only the people you choose.</p>
    <div class='hero-credit'>By Patrick Bush</div>
  </header>

  <section class='summary-grid' aria-label='Report summary'>
    <div class='summary-card'><div class='label'>PST</div><div class='value'>$(ConvertTo-HtmlEncodedText $PstItem.Name)</div></div>
    <div class='summary-card'><div class='label'>Generated</div><div class='value'>$(ConvertTo-HtmlEncodedText $generated)</div></div>
    <div class='summary-card'><div class='label'>Messages exported</div><div class='value'>$(ConvertTo-HtmlEncodedText $SortedRecords.Count)</div></div>
    <div class='summary-card'><div class='label'>People detected</div><div class='value'>$(ConvertTo-HtmlEncodedText $AllParticipants.Count)</div></div>
    <div class='summary-card'><div class='label'>Folders scanned</div><div class='value'>$(ConvertTo-HtmlEncodedText $script:Stats.FoldersScanned)</div></div>
    <div class='summary-card'><div class='label'>Read warnings</div><div class='value'>Items: $(ConvertTo-HtmlEncodedText $script:Stats.ItemReadFailures); Attachments: $(ConvertTo-HtmlEncodedText $script:Stats.AttachmentReadFailures)</div></div>
  </section>

  <div class='review-layout'>
    <aside class='filter-panel' aria-label='Conversation filters'>
      <div class='filter-title'>
        <h2>Choose whose conversations to view</h2>
        <span id='resultCount' class='result-count'></span>
      </div>
      <p class='filter-help'>Check one or more names, then choose how strictly to match them. Use "Exact selected people, ignoring Other/IDs" when Purview adds bots, meeting artifacts, or system IDs that should not count as real conversation participants.</p>
      <div class='controls'>
        <input id='personSearch' type='text' placeholder='Find a person in the list...'/>
        <input id='messageSearch' type='text' placeholder='Search within shown conversations...'/>
        <label for='participantMatchMode'>Participant match</label>
        <select id='participantMatchMode'>
          <option value='messagesFromSelected' selected='selected'>Messages from selected people</option>
          <option value='involving'>Conversations involving selected people</option>
          <option value='exactAll'>Exact selected people only</option>
          <option value='exactPeopleOnly'>Exact selected people, ignoring Other/IDs</option>
        </select>
      </div>
      <div class='people-heading'><h3>People</h3><button type='button' class='secondary' id='clearAll'>Clear filters</button></div>
      <div id='peopleBox' class='people-box'>
        <label class='person-option'><input type='checkbox' id='selectAllPeople'/> <span>Select All</span></label>
$($ParticipantOptions -join "`n")
      </div>
      <details class='other-names'>
        <summary>Other detected names / IDs ($($OtherParticipants.Count))</summary>
        <div id='otherPeopleBox' class='people-box'>
          <label class='person-option'><input type='checkbox' id='selectAllOther'/> <span>Select All</span></label>
$($OtherParticipantOptions -join "`n")
        </div>
      </details>
      <div class='notice'>Message bodies are HTML-encoded as text for safe review. Attachment files are not extracted; attachment metadata is listed where Outlook exposes it.</div>
      <details class='folder-summary'>
        <summary>Folder summary</summary>
        <table>
          <thead><tr><th>Folder</th><th>Items</th></tr></thead>
          <tbody>
$($FolderSummaryRows -join "`n")
          </tbody>
        </table>
      </details>
      <div class='footer'>Log file: $(ConvertTo-HtmlEncodedText $script:LogPath)<br/>Created by Convert-PurviewTeamsPstToHtml.ps1</div>
    </aside>

    <div id='resizeHandle' class='resize-handle' role='separator' aria-orientation='vertical' aria-label='Resize filter column' tabindex='0' title='Drag to resize the filter column'></div>

    <main id='conversationList' class='conversation-pane'>
      <section class='conversation-toolbar' aria-label='Date filter and conversation sorting'>
        <div class='toolbar-field'>
          <label for='startDateFilter'>Start date</label>
          <input id='startDateFilter' type='date'/>
        </div>
        <div class='toolbar-field'>
          <label for='endDateFilter'>End date</label>
          <input id='endDateFilter' type='date'/>
        </div>
        <div class='toolbar-field'>
          <label for='sortOrder'>Sort conversations</label>
          <select id='sortOrder'>
            <option value='newestFirst' selected='selected'>Newest first</option>
            <option value='oldestFirst'>Oldest first</option>
          </select>
        </div>
      </section>
"@)
}

function Write-ConversationHtml {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)][object]$GroupMessages,
        [Parameter(Mandatory = $true)][hashtable]$PreferredParticipants,
        [Parameter(Mandatory = $true)][hashtable]$SenderClasses,
        [Parameter(Mandatory = $true)][ref]$NextSenderIndex
    )

    $GroupMessages = @($GroupMessages)
    if ($GroupMessages.Count -eq 0) { return }
    $palette = @('blue','green','purple','orange','red','teal','pink','brown','slate','indigo')
    $first = $GroupMessages[0]
    $groupParticipants = @($first.Participants | Sort-Object -Unique)
    $groupPersonParticipants = @($groupParticipants | Where-Object { Test-LikelyPersonName $_ })
    $groupOtherParticipants = @($groupParticipants | Where-Object { -not (Test-LikelyPersonName $_) })
    $displayGroupParticipants = @(Get-ParticipantDisplayOrder -Participants $groupParticipants -PreferredParticipants $PreferredParticipants)
    $participantData = ConvertTo-HtmlEncodedText ($groupParticipants -join '||')
    $personParticipantData = ConvertTo-HtmlEncodedText ($groupPersonParticipants -join '||')
    $otherParticipantData = ConvertTo-HtmlEncodedText ($groupOtherParticipants -join '||')
    $participantLabel = if ($displayGroupParticipants.Count -gt 0) { Format-ParticipantPreview -Participants $displayGroupParticipants -MaximumNames 4 -PreferredParticipants $PreferredParticipants } else { '(unknown participants)' }
    $conversationTitle = if ([string]::IsNullOrWhiteSpace([string]$first.ConversationTitle)) { $participantLabel } else { [string]$first.ConversationTitle }
    $conversationSortRecord = @($GroupMessages | Sort-Object @{ Expression = { if ($_.SortTime) { ([datetime]$_.SortTime).Ticks } else { 0 } } } | Select-Object -Last 1)
    $conversationSortTime = if ($conversationSortRecord -and $conversationSortRecord[0].SortTime) { ([datetime]$conversationSortRecord[0].SortTime).ToString('o') } else { '' }

    $Writer.WriteLine(@"
<section class='conversation' data-participants='$participantData' data-person-participants='$personParticipantData' data-other-participants='$otherParticipantData' data-sort-time='$(ConvertTo-HtmlEncodedText $conversationSortTime)'>
  <div class='conversation-header'>
    <div>
      <h2>$(ConvertTo-HtmlEncodedText $conversationTitle)</h2>
      <div class='conversation-people'>$(ConvertTo-HtmlEncodedText $participantLabel)</div>
    </div>
    <div class='conversation-count'>$($GroupMessages.Count) messages</div>
  </div>
  <div class='conversation-messages'>
"@)

    # Normalize the conversation's participant set once; every message reuses these instead of
    # re-normalizing the whole group per message when building its "other people" list.
    $groupParticipantInfo = @($groupParticipants | ForEach-Object { [pscustomobject]@{ Name = $_; Normalized = (ConvertTo-NormalizedPersonName $_) } })

    foreach ($record in $GroupMessages) {
        if (-not $SenderClasses.ContainsKey($record.SenderDisplay)) {
            $SenderClasses[$record.SenderDisplay] = $palette[$NextSenderIndex.Value % $palette.Count]
            $NextSenderIndex.Value++
        }
        $senderClass = $SenderClasses[$record.SenderDisplay]
        $timeText = if ($record.SortTime) { ([datetime]$record.SortTime).ToString('MMM d, yyyy h:mm tt') } else { '' }
        $messageDate = if ($record.SortTime) { ([datetime]$record.SortTime).ToString('yyyy-MM-dd') } else { '' }
        $messageSortTime = if ($record.SortTime) { ([datetime]$record.SortTime).ToString('o') } else { '' }
        $senderParticipantName = ConvertTo-NormalizedPersonName $record.SenderDisplay
        $otherPeople = @($groupParticipantInfo | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.Normalized) -and $_.Normalized -ne $senderParticipantName
        } | ForEach-Object { $_.Name })
        if ($otherPeople.Count -eq 0 -and $groupParticipants.Count -gt 1) {
            $otherPeople = @($groupParticipants | Where-Object { $_ -ne $record.SenderDisplay })
        }
        $displayOtherPeople = @(Get-ParticipantDisplayOrder -Participants $otherPeople -PreferredParticipants $PreferredParticipants)
        $otherPeoplePreview = Format-ParticipantPreview -Participants $displayOtherPeople -MaximumNames 4 -PreferredParticipants $PreferredParticipants
        $otherPeopleLabel = if ($displayOtherPeople.Count -gt 0) { 'with ' + $otherPeoplePreview } else { 'no other named participants shown' }
        $speakerContext = "Chat: $conversationTitle • $otherPeopleLabel"
        $bodyHtml = ConvertTo-HtmlBody $record.BodyText
        # Blank the 4501 "no date" sentinel in the detail rows too, matching the header timestamp.
        $sentDisplay = if (Test-MissingDate $record.SentOn) { '' } else { $record.SentOn }
        $receivedDisplay = if (Test-MissingDate $record.ReceivedTime) { '' } else { $record.ReceivedTime }
        $details = @"
<details class='message-details'>
  <summary>Details</summary>
  <div><strong>Folder:</strong> $(ConvertTo-HtmlEncodedText $record.FolderPath)</div>
  <div><strong>Subject:</strong> $(ConvertTo-HtmlEncodedText $record.Subject)</div>
  <div><strong>Message class:</strong> $(ConvertTo-HtmlEncodedText $record.MessageClass)</div>
  <div><strong>From:</strong> $(ConvertTo-HtmlEncodedText $record.SenderName) &lt;$(ConvertTo-HtmlEncodedText $record.SenderEmail)&gt;</div>
  <div><strong>To:</strong> $(ConvertTo-HtmlEncodedText $record.To)</div>
  <div><strong>Cc:</strong> $(ConvertTo-HtmlEncodedText $record.Cc)</div>
  <div><strong>Sent:</strong> $(ConvertTo-HtmlEncodedText $sentDisplay)</div>
  <div><strong>Received:</strong> $(ConvertTo-HtmlEncodedText $receivedDisplay)</div>
  <div><strong>Entry ID:</strong> <span class='wrap'>$(ConvertTo-HtmlEncodedText $record.EntryId)</span></div>
</details>
"@
        $Writer.WriteLine(@"
<article class='message-card sender-$senderClass' data-sender='$(ConvertTo-HtmlEncodedText $record.SenderDisplay)' data-date='$(ConvertTo-HtmlEncodedText $messageDate)' data-time='$(ConvertTo-HtmlEncodedText $messageSortTime)'>
  <div class='speaker-row'>
    <span class='speaker-block'><span class='speaker-name'>$(ConvertTo-HtmlEncodedText $record.SenderDisplay)</span><span class='speaker-context'>$(ConvertTo-HtmlEncodedText $speakerContext)</span></span>
    <span class='message-time'>$(ConvertTo-HtmlEncodedText $timeText)</span>
  </div>
  <div class='message-body'>$bodyHtml</div>
  $($record.AttachmentsHtml)
  $details
</article>
"@)
    }

    $Writer.WriteLine(@"
  </div>
</section>
"@)
}

function Write-ReportFooter {
    param([Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer)

    $script = Get-ReportScript
    $Writer.WriteLine(@"
    </main>
  </div>
</div>
<script>
$script
</script>
</body>
</html>
"@)
}

# ponytail: keep in sync with sample email threading and Write-EmailHtmlReport
function Get-NormalizedEmailSubject {
    param([AllowNull()][object]$Subject)

    $s = [string]$Subject
    while ($s -match '(?i)^(re|fw|fwd)\s*:\s*') {
        $s = $s -replace '(?i)^(re|fw|fwd)\s*:\s*', ''
    }
    if ([string]::IsNullOrWhiteSpace($s)) { return '(no subject)' }
    return $s.Trim()
}

# ponytail: keep in sync with sample email threading and Write-EmailHtmlReport
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

function Get-EmailFolderLeafLabel {
    param([AllowNull()][string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return '(unknown folder)' }
    $parts = @($FolderPath -split '[\\/]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -eq 0) { return '(unknown folder)' }
    return $parts[-1]
}

function Format-FolderOptionHtml {
    param(
        [Parameter(Mandatory = $true)] [string]$FolderPath,
        [Parameter(Mandatory = $true)] [int]$MessageCount
    )

    $folderValue = if ([string]::IsNullOrWhiteSpace($FolderPath)) { '(unknown folder)' } else { $FolderPath }
    $folderLabel = '{0} ({1})' -f (Get-EmailFolderLeafLabel -FolderPath $FolderPath), $MessageCount
    # ponytail: checked on load; JS treats none-checked the same as all-checked (show all)
    return "<label class='folder-option' title='$(ConvertTo-HtmlEncodedText $folderValue)'><input type='checkbox' class='folder-check' value='$(ConvertTo-HtmlEncodedText $folderValue)' checked='checked'/> <span>$(ConvertTo-HtmlEncodedText $folderLabel)</span></label>"
}

function Format-TopSenderOptionHtml {
    param(
        [Parameter(Mandatory = $true)] [string]$SenderDisplay,
        [Parameter(Mandatory = $true)] [int]$MessageCount
    )

    return "<button type='button' class='top-sender' data-query='$(ConvertTo-HtmlEncodedText $SenderDisplay)'>$(ConvertTo-HtmlEncodedText $SenderDisplay) <span class='top-sender-count'>($MessageCount)</span></button>"
}

function Get-EmailReportCss {
    $baseCss = Get-ReportCss
    return $baseCss + @'
.folder-filter .folder-box { max-height: 32vh; overflow: auto; border: 1px solid var(--line); border-radius: 12px; padding: 8px; background: #fbfcff; display: grid; grid-template-columns: 1fr; gap: 5px; }
.folder-option { display: flex; gap: 8px; align-items: center; padding: 6px 7px; border-radius: 9px; cursor: pointer; font-size: .92rem; font-weight: 700; color: #334155; }
.folder-option:hover { background: #eef3ff; }
.folder-actions { display: flex; gap: 8px; flex-wrap: wrap; margin: 8px 0; }
.email-summary { margin-top: 10px; }
.email-summary summary { cursor: pointer; font-weight: 800; color: #334155; }
.email-summary .folder-box { border: 0; border-top: 1px solid var(--line); border-radius: 0 0 12px 12px; max-height: 24vh; }
.people-search-wrap { display: grid; gap: 6px; margin-top: 4px; }
.people-search-wrap input[type='search'] { width: 100%; box-sizing: border-box; }
.top-senders { display: flex; flex-wrap: wrap; gap: 6px; padding-top: 8px; }
.top-sender { border: 1px solid var(--line); background: #fbfcff; border-radius: 999px; padding: 5px 10px; font-size: .85rem; font-weight: 700; color: #334155; cursor: pointer; }
.top-sender:hover { background: #eef3ff; }
.top-sender-count { color: var(--muted); font-weight: 600; }
.read-warnings { margin-top: 12px; color: #7c2d12; font-size: .9rem; }
.read-warnings summary { cursor: pointer; font-weight: 800; }
.email-note { margin-top: 10px; color: var(--muted); font-size: .9rem; }
.email-from-line { font-size: .95rem; color: #334155; font-weight: 700; }
.email-details div { margin: 2px 0; }
/* Large reports: skip offscreen layout/paint without changing UX. */
.conversation { content-visibility: auto; contain-intrinsic-size: auto 220px; }
'@
}

function Get-EmailReportScript {
    # ponytail: no textContent indexing — that freezes 20k+ thread / 250MB reports on open
    return @'
(function () {
  const folderChecks = Array.from(document.querySelectorAll('.folder-check'));
  const selectAllFoldersBtn = document.getElementById('selectAllFoldersBtn');
  const clearFoldersBtn = document.getElementById('clearFoldersBtn');
  const peopleSearch = document.getElementById('peopleSearch');
  const topSenderButtons = Array.from(document.querySelectorAll('.top-sender'));
  const startDateFilter = document.getElementById('startDateFilter');
  const endDateFilter = document.getElementById('endDateFilter');
  const sortOrder = document.getElementById('sortOrder');
  const conversationList = document.getElementById('conversationList');
  const conversations = Array.from(document.querySelectorAll('.conversation'));
  const resultCount = document.getElementById('resultCount');
  const totalConversations = conversations.length;
  const folderCount = folderChecks.length;
  let totalMessages = 0;
  let messagesCached = false;
  let lastSortOrder = sortOrder ? (sortOrder.value || 'newestFirst') : 'newestFirst';
  let filterTimer = 0;

  conversations.forEach(conversation => {
    conversation._participants = (conversation.dataset.participants || '').toLowerCase();
    conversation._folders = null;
    conversation._msgCount = parseInt(conversation.dataset.msgCount || '0', 10) || 0;
    conversation._sortTimeMs = Date.parse(conversation.dataset.sortTime || '') || 0;
    conversation._countEl = null;
    conversation._messages = null;
    totalMessages += conversation._msgCount;
  });

  function selectedValues(checks) { return checks.filter(c => c.checked).map(c => c.value); }
  function getFolders(conversation) {
    if (!conversation._folders) {
      conversation._folders = (conversation.dataset.folder || '').split('||').map(v => v.trim()).filter(Boolean);
    }
    return conversation._folders;
  }
  function getCountEl(conversation) {
    if (conversation._countEl === null) conversation._countEl = conversation.querySelector('.conversation-count');
    return conversation._countEl;
  }
  function messageInDateRange(message, startDate, endDate) {
    const date = (message.dataset.date || '').trim();
    if (!date) return true;
    if (startDate && date < startDate) return false;
    if (endDate && date > endDate) return false;
    return true;
  }
  // Checked folders only; none checked OR all checked => show every folder (same as no folder filter).
  function folderFilterActive(selectedFolders) {
    return selectedFolders.length > 0 && selectedFolders.length < folderCount;
  }
  function matchesFolder(conversation, selectedFolders) {
    if (!folderFilterActive(selectedFolders)) return true;
    return selectedFolders.some(folder => getFolders(conversation).includes(folder));
  }
  function matchesPeople(conversation, query) {
    if (!query) return true;
    return conversation._participants.includes(query.toLowerCase());
  }
  function ensureMessages(conversation) {
    if (!conversation._messages) {
      conversation._messages = Array.from(conversation.querySelectorAll('.message-card'));
      if (!conversation._msgCount) conversation._msgCount = conversation._messages.length;
    }
    return conversation._messages;
  }
  function ensureAllMessages() {
    if (messagesCached) return;
    conversations.forEach(ensureMessages);
    messagesCached = true;
  }
  function setCountText(conversation, text) {
    const countEl = getCountEl(conversation);
    if (countEl) countEl.textContent = text;
  }
  function setResultCount(visibleConversations, visibleMessages) {
    if (resultCount) resultCount.textContent = visibleConversations + ' conversations / ' + visibleMessages + ' messages shown';
  }
  function sortConversations(force) {
    if (!conversationList || !sortOrder) return;
    const order = sortOrder.value || 'newestFirst';
    if (!force && order === lastSortOrder) return;
    lastSortOrder = order;
    const sorted = conversations.slice().sort((a, b) => order === 'oldestFirst' ? a._sortTimeMs - b._sortTimeMs : b._sortTimeMs - a._sortTimeMs);
    const frag = document.createDocumentFragment();
    sorted.forEach(conversation => frag.appendChild(conversation));
    conversationList.appendChild(frag);
  }

  function applyFilters() {
    const selectedFolders = selectedValues(folderChecks);
    const peopleQuery = peopleSearch ? peopleSearch.value.trim() : '';
    const startDate = startDateFilter ? startDateFilter.value : '';
    const endDate = endDateFilter ? endDateFilter.value : '';
    const filtersActive = folderFilterActive(selectedFolders) || !!peopleQuery || !!startDate || !!endDate;

    if (!filtersActive) {
      conversations.forEach(conversation => {
        if (conversation.hidden) conversation.hidden = false;
        if (conversation._messages) {
          conversation._messages.forEach(message => { if (message.hidden) message.hidden = false; });
        }
        setCountText(conversation, conversation._msgCount + ' messages');
      });
      sortConversations(false);
      setResultCount(totalConversations, totalMessages || totalConversations);
      return;
    }

    let visibleConversations = 0;
    let visibleMessages = 0;
    const dateFilterActive = !!(startDate || endDate);

    if (!dateFilterActive) {
      conversations.forEach(conversation => {
        const show = matchesFolder(conversation, selectedFolders) && matchesPeople(conversation, peopleQuery);
        conversation.hidden = !show;
        if (conversation._messages) {
          conversation._messages.forEach(message => { if (message.hidden) message.hidden = false; });
        }
        setCountText(conversation, conversation._msgCount + ' messages');
        if (show) {
          visibleConversations += 1;
          visibleMessages += conversation._msgCount;
        }
      });
    } else {
      ensureAllMessages();
      conversations.forEach(conversation => {
        const conversationMatches = matchesFolder(conversation, selectedFolders) && matchesPeople(conversation, peopleQuery);
        const messages = conversation._messages;
        let conversationVisibleMessages = 0;
        messages.forEach(message => {
          const showMessage = conversationMatches && messageInDateRange(message, startDate, endDate);
          message.hidden = !showMessage;
          if (showMessage) conversationVisibleMessages += 1;
        });
        conversation.hidden = conversationVisibleMessages === 0;
        setCountText(conversation, conversationVisibleMessages + ' of ' + messages.length + ' messages');
        if (!conversation.hidden) {
          visibleConversations += 1;
          visibleMessages += conversationVisibleMessages;
        }
      });
    }

    sortConversations(false);
    setResultCount(visibleConversations, visibleMessages);
  }

  function scheduleFilters() {
    clearTimeout(filterTimer);
    filterTimer = setTimeout(applyFilters, 150);
  }

  function setAllFolders(checked) {
    folderChecks.forEach(c => { c.checked = checked; });
    applyFilters();
  }

  folderChecks.forEach(c => c.addEventListener('change', applyFilters));
  if (selectAllFoldersBtn) selectAllFoldersBtn.addEventListener('click', () => setAllFolders(true));
  if (clearFoldersBtn) clearFoldersBtn.addEventListener('click', () => setAllFolders(false));
  if (peopleSearch) {
    peopleSearch.addEventListener('input', scheduleFilters);
    peopleSearch.addEventListener('search', applyFilters);
  }
  topSenderButtons.forEach(button => button.addEventListener('click', () => {
    if (!peopleSearch) return;
    peopleSearch.value = button.dataset.query || '';
    peopleSearch.focus();
    applyFilters();
  }));
  if (startDateFilter) startDateFilter.addEventListener('change', applyFilters);
  if (endDateFilter) endDateFilter.addEventListener('change', applyFilters);
  if (sortOrder) sortOrder.addEventListener('change', () => { sortConversations(true); applyFilters(); });

  // Defer result-count paint; skip full filter walk when nothing is filtered.
  requestAnimationFrame(() => setResultCount(totalConversations, totalMessages || totalConversations));
})();
'@
}

function Write-EmailReportHeader {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)]$PstItem,
        [Parameter(Mandatory = $true)][int]$ConversationCount,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SortedRecords,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$AllParticipants,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$AllFolders,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$FolderOptions,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$FolderSummaryRows,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$TopSenderOptions
    )

    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss K'
    $css = Get-EmailReportCss
    $itemWarn = [int]$script:Stats.ItemReadFailures
    $attachWarn = [int]$script:Stats.AttachmentReadFailures
    $readWarningsBlock = ''
    if ($itemWarn -gt 0 -or $attachWarn -gt 0) {
        $readWarningsBlock = @"
      <details class='read-warnings'>
        <summary>Read warnings</summary>
        <div>Items: $(ConvertTo-HtmlEncodedText $itemWarn); Attachments: $(ConvertTo-HtmlEncodedText $attachWarn)</div>
      </details>
"@
    }
    $topSendersBlock = ''
    if ($TopSenderOptions.Count -gt 0) {
        $topSendersBlock = @"
      <details class='email-summary top-senders-summary'>
        <summary>Top senders</summary>
        <div class='top-senders'>
$($TopSenderOptions -join "`n")
        </div>
      </details>
"@
    }
    $Writer.WriteLine(@"
<!doctype html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Purview Email PST Report - $(ConvertTo-HtmlEncodedText $PstItem.Name)</title>
<style>
$css
</style>
</head>
<body>
<div class='page'>
  <header class='hero'>
    <h1>Microsoft Purview eDiscovery Email Report</h1>
    <p>Threaded email conversations grouped by Outlook conversation data when available, with folder and people filters.</p>
    <div class='hero-credit'>By Patrick Bush</div>
  </header>

  <section class='summary-grid' aria-label='Report summary'>
    <div class='summary-card'><div class='label'>PST</div><div class='value'>$(ConvertTo-HtmlEncodedText $PstItem.Name)</div></div>
    <div class='summary-card'><div class='label'>Generated</div><div class='value'>$(ConvertTo-HtmlEncodedText $generated)</div></div>
    <div class='summary-card'><div class='label'>Threads exported</div><div class='value'>$(ConvertTo-HtmlEncodedText $ConversationCount)</div></div>
    <div class='summary-card'><div class='label'>Messages exported</div><div class='value'>$(ConvertTo-HtmlEncodedText $SortedRecords.Count)</div></div>
    <div class='summary-card'><div class='label'>Folders detected</div><div class='value'>$(ConvertTo-HtmlEncodedText $AllFolders.Count)</div></div>
    <div class='summary-card'><div class='label'>People detected</div><div class='value'>$(ConvertTo-HtmlEncodedText $AllParticipants.Count)</div></div>
  </section>

  <div class='review-layout'>
    <aside class='filter-panel folder-filter' aria-label='Conversation filters'>
      <div class='filter-title'>
        <h2>Choose folders and people</h2>
        <span id='resultCount' class='result-count'></span>
      </div>
      <p class='filter-help'>Folders start checked (show all). Uncheck folders to hide them. If none are checked, everything is shown. Type a name or address to narrow by people.</p>
      <details class='email-summary folder-summary' open='open'>
        <summary>Folders</summary>
        <div class='folder-actions'>
          <button type='button' class='secondary' id='selectAllFoldersBtn'>Select all</button>
          <button type='button' class='secondary' id='clearFoldersBtn'>Clear</button>
        </div>
        <div id='folderBox' class='folder-box'>
$($FolderOptions -join "`n")
        </div>
      </details>
      <div class='people-heading'><h3>People</h3></div>
      <div class='people-search-wrap'>
        <label for='peopleSearch'>Search people</label>
        <input id='peopleSearch' type='search' placeholder='Name or email address' autocomplete='off'/>
      </div>
$topSendersBlock
      <div class='email-note'>Message bodies are HTML-encoded for safe review. Folder labels show the leaf name (hover for full path); matching uses the full path. Full paths also appear in each message's Details.</div>
      <details class='folder-summary'>
        <summary>Folder summary</summary>
        <table>
          <thead><tr><th>Folder</th><th>Messages</th></tr></thead>
          <tbody>
$($FolderSummaryRows -join "`n")
          </tbody>
        </table>
      </details>
$readWarningsBlock
      <div class='footer'>Log file: $(ConvertTo-HtmlEncodedText $script:LogPath)<br/>Created by Convert-PurviewTeamsPstToHtml.ps1</div>
    </aside>

    <div id='resizeHandle' class='resize-handle' role='separator' aria-orientation='vertical' aria-label='Resize filter column' tabindex='0' title='Drag to resize the filter column'></div>

    <main id='conversationList' class='conversation-pane'>
      <section class='conversation-toolbar' aria-label='Date filter and conversation sorting'>
        <div class='toolbar-field'>
          <label for='startDateFilter'>Start date</label>
          <input id='startDateFilter' type='date'/>
        </div>
        <div class='toolbar-field'>
          <label for='endDateFilter'>End date</label>
          <input id='endDateFilter' type='date'/>
        </div>
        <div class='toolbar-field'>
          <label for='sortOrder'>Sort conversations</label>
          <select id='sortOrder'>
            <option value='newestFirst' selected='selected'>Newest first</option>
            <option value='oldestFirst'>Oldest first</option>
          </select>
        </div>
      </section>
"@)
}

function Write-EmailConversationHtml {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)][object]$GroupMessages,
        [Parameter(Mandatory = $true)][hashtable]$SenderClasses,
        [Parameter(Mandatory = $true)][ref]$NextSenderIndex
    )

    $GroupMessages = @($GroupMessages)
    if ($GroupMessages.Count -eq 0) { return }

    $palette = @('blue','green','purple','orange','red','teal','pink','brown','slate','indigo')
    $first = $GroupMessages[0]
    $groupParticipants = @(
        $GroupMessages | ForEach-Object { @($_.Participants) } | ForEach-Object { $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
    )
    $groupFolders = @($GroupMessages | ForEach-Object { $_.FolderPath } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    $groupSenders = @($GroupMessages | ForEach-Object { $_.SenderDisplay } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    $participantData = ConvertTo-HtmlEncodedText ($groupParticipants -join '||')
    $folderData = ConvertTo-HtmlEncodedText ($groupFolders -join '||')
    $senderPreview = if ($groupSenders.Count -gt 0) { Format-ParticipantPreview -Participants $groupSenders -MaximumNames 4 -PreferredParticipants @{} } else { '(unknown sender)' }
    $participantPreview = if ($groupParticipants.Count -gt 0) { Format-ParticipantPreview -Participants $groupParticipants -MaximumNames 4 -PreferredParticipants @{} } else { '(unknown participants)' }
    $conversationTitle = Get-NormalizedEmailSubject -Subject $first.Subject
    if ([string]::IsNullOrWhiteSpace($conversationTitle) -or $conversationTitle -eq '(no subject)') {
        if (-not [string]::IsNullOrWhiteSpace([string]$first.ConversationTopic)) {
            $conversationTitle = [string]$first.ConversationTopic
        }
        else {
            $conversationTitle = $participantPreview
        }
    }

    $conversationSortRecord = @($GroupMessages | Sort-Object @{ Expression = { if ($_.SortTime) { ([datetime]$_.SortTime).Ticks } else { 0 } } } | Select-Object -Last 1)
    $conversationSortTime = if ($conversationSortRecord -and $conversationSortRecord[0].SortTime) { ([datetime]$conversationSortRecord[0].SortTime).ToString('o') } else { '' }

    $Writer.WriteLine(@"
<section class='conversation' data-folder='$folderData' data-participants='$participantData' data-msg-count='$($GroupMessages.Count)' data-sort-time='$(ConvertTo-HtmlEncodedText $conversationSortTime)'>
  <div class='conversation-header'>
    <div>
      <h2>$(ConvertTo-HtmlEncodedText $conversationTitle)</h2>
      <div class='conversation-people'>From: $(ConvertTo-HtmlEncodedText $senderPreview)</div>
      <div class='conversation-people'>Participants: $(ConvertTo-HtmlEncodedText $participantPreview)</div>
    </div>
    <div class='conversation-count'>$($GroupMessages.Count) messages</div>
  </div>
  <div class='conversation-messages'>
"@)

    foreach ($record in $GroupMessages) {
        if (-not $SenderClasses.ContainsKey($record.SenderDisplay)) {
            $SenderClasses[$record.SenderDisplay] = $palette[$NextSenderIndex.Value % $palette.Count]
            $NextSenderIndex.Value++
        }
        $senderClass = $SenderClasses[$record.SenderDisplay]
        $timeText = if ($record.SortTime) { ([datetime]$record.SortTime).ToString('MMM d, yyyy h:mm tt') } else { '' }
        $messageDate = if ($record.SortTime) { ([datetime]$record.SortTime).ToString('yyyy-MM-dd') } else { '' }
        $messageSortTime = if ($record.SortTime) { ([datetime]$record.SortTime).ToString('o') } else { '' }
        $bodyHtml = ConvertTo-HtmlBody $record.BodyText
        $sentDisplay = if (Test-MissingDate $record.SentOn) { '' } else { $record.SentOn }
        $receivedDisplay = if (Test-MissingDate $record.ReceivedTime) { '' } else { $record.ReceivedTime }
        $senderLine = if ([string]::IsNullOrWhiteSpace([string]$record.SenderDisplay)) { '(unknown sender)' } else { [string]$record.SenderDisplay }
        $folderLeaf = Get-EmailFolderLeafLabel -FolderPath $record.FolderPath
        $details = @"
<details class='message-details email-details'>
  <summary>Details</summary>
  <div><strong>Folder:</strong> $(ConvertTo-HtmlEncodedText $record.FolderPath)</div>
  <div><strong>Subject:</strong> $(ConvertTo-HtmlEncodedText (Get-NormalizedEmailSubject -Subject $record.Subject))</div>
  <div><strong>From:</strong> $(ConvertTo-HtmlEncodedText $record.SenderName) &lt;$(ConvertTo-HtmlEncodedText $record.SenderEmail)&gt;</div>
  <div><strong>To:</strong> $(ConvertTo-HtmlEncodedText $record.To)</div>
  <div><strong>Cc:</strong> $(ConvertTo-HtmlEncodedText $record.Cc)</div>
  <div><strong>Conversation topic:</strong> $(ConvertTo-HtmlEncodedText $record.ConversationTopic)</div>
  <div><strong>Conversation ID:</strong> <span class='wrap'>$(ConvertTo-HtmlEncodedText $record.ConversationId)</span></div>
  <div><strong>Message class:</strong> $(ConvertTo-HtmlEncodedText $record.MessageClass)</div>
  <div><strong>Sent:</strong> $(ConvertTo-HtmlEncodedText $sentDisplay)</div>
  <div><strong>Received:</strong> $(ConvertTo-HtmlEncodedText $receivedDisplay)</div>
  <div><strong>Entry ID:</strong> <span class='wrap'>$(ConvertTo-HtmlEncodedText $record.EntryId)</span></div>
</details>
"@
        $Writer.WriteLine(@"
<article class='message-card sender-$senderClass' data-sender='$(ConvertTo-HtmlEncodedText $record.SenderDisplay)' data-date='$(ConvertTo-HtmlEncodedText $messageDate)' data-time='$(ConvertTo-HtmlEncodedText $messageSortTime)'>
  <div class='speaker-row'>
    <span class='speaker-block'><span class='speaker-name'>From: $(ConvertTo-HtmlEncodedText $senderLine)</span><span class='speaker-context'>To: $(ConvertTo-HtmlEncodedText $record.To) • Cc: $(ConvertTo-HtmlEncodedText $record.Cc) • Folder: $(ConvertTo-HtmlEncodedText $folderLeaf)</span></span>
    <span class='message-time'>$(ConvertTo-HtmlEncodedText $timeText)</span>
  </div>
  <div class='message-body'>$bodyHtml</div>
  $($record.AttachmentsHtml)
  $details
</article>
"@)
    }

    $Writer.WriteLine(@"
  </div>
</section>
"@)
}

function Write-EmailReportFooter {
    param([Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer)

    $script = Get-EmailReportScript
    $Writer.WriteLine(@"
    </main>
  </div>
</div>
<script>
$script
</script>
</body>
</html>
"@)
}

# Task 4 will replace with full email UX.
function Get-SampleRecord {
    return [pscustomobject]@{
        Teams = @(
            [pscustomobject]@{ SortTime = [datetime]'2024-01-01T09:00:00'; FolderPath = 'SamplePst\TeamsMessagesData'; Subject = 'Sample Teams chat'; MessageClass = 'IPM.Note'; SenderName = 'Torey Page'; SenderEmail = 'torey@example.com'; SenderDisplay = 'Torey Page'; To = 'Linda Artley'; Cc = ''; Participants = @('Linda Artley','Torey Page'); ParticipantsKey = 'Linda Artley || Torey Page'; ConversationKey = "Linda Artley || Torey Page`nSample Teams chat`nSamplePst\TeamsMessagesData"; ConversationTitle = 'Sample Teams chat'; SentOn = [datetime]'2024-01-01T09:00:00'; ReceivedTime = [datetime]'2024-01-01T09:00:05'; CreationTime = [datetime]'2024-01-01T09:00:05'; EntryId = 'sample-entry-1'; BodyText = 'Hello Linda. This sample validates HTML encoding: <script>alert(1)</script>'; AttachmentsHtml = '' }
            [pscustomobject]@{ SortTime = [datetime]'2024-01-01T09:01:00'; FolderPath = 'SamplePst\TeamsMessagesData'; Subject = 'Sample Teams chat'; MessageClass = 'IPM.Note'; SenderName = 'Linda Artley'; SenderEmail = 'linda@example.com'; SenderDisplay = 'Linda Artley'; To = 'Torey Page'; Cc = ''; Participants = @('Linda Artley','Torey Page'); ParticipantsKey = 'Linda Artley || Torey Page'; ConversationKey = "Linda Artley || Torey Page`nSample Teams chat`nSamplePst\TeamsMessagesData"; ConversationTitle = 'Sample Teams chat'; SentOn = [datetime]'2024-01-01T09:01:00'; ReceivedTime = [datetime]'2024-01-01T09:01:05'; CreationTime = [datetime]'2024-01-01T09:01:05'; EntryId = 'sample-entry-2'; BodyText = "Thanks.`nThis message has two lines."; AttachmentsHtml = "<div class='attachments'><strong>Attachments:</strong><table><tbody><tr><td>sample.pdf</td><td>sample.pdf</td><td>1234</td></tr></tbody></table></div>" }
        )
        Email = @(
            [pscustomobject]@{ SortTime = [datetime]'2024-01-01T10:00:00'; FolderPath = 'SamplePst\Inbox'; Subject = 'Budget'; MessageClass = 'IPM.Note'; SenderName = 'Linda Artley'; SenderEmail = 'linda@example.com'; SenderDisplay = 'Linda Artley'; To = 'Torey Page'; Cc = ''; Participants = @('Linda Artley','Torey Page'); ParticipantsKey = 'Linda Artley || Torey Page'; ConversationKey = "id:sample-conversation-1"; ConversationTitle = 'Budget'; ConversationTopic = 'Budget planning'; ConversationId = 'sample-conversation-1'; SentOn = [datetime]'2024-01-01T10:00:00'; ReceivedTime = [datetime]'2024-01-01T10:00:05'; CreationTime = [datetime]'2024-01-01T10:00:05'; EntryId = 'sample-entry-3'; BodyText = 'Budget phoenix attached.'; AttachmentsHtml = '' }
            [pscustomobject]@{ SortTime = [datetime]'2024-01-01T10:05:00'; FolderPath = 'SamplePst\Inbox'; Subject = 'Re: Budget'; MessageClass = 'IPM.Note'; SenderName = 'Torey Page'; SenderEmail = 'torey@example.com'; SenderDisplay = 'Torey Page'; To = 'Linda Artley'; Cc = ''; Participants = @('Linda Artley','Torey Page'); ParticipantsKey = 'Linda Artley || Torey Page'; ConversationKey = "id:sample-conversation-1"; ConversationTitle = 'Budget'; ConversationTopic = 'Budget planning'; ConversationId = 'sample-conversation-1'; SentOn = [datetime]'2024-01-01T10:05:00'; ReceivedTime = [datetime]'2024-01-01T10:05:05'; CreationTime = [datetime]'2024-01-01T10:05:05'; EntryId = 'sample-entry-4'; BodyText = 'Thanks, I will review it.'; AttachmentsHtml = '' }
        )
        Calendar = @(
            [pscustomobject]@{ SortTime = [datetime]'2024-01-08T09:00:00'; StartTime = [datetime]'2024-01-08T09:00:00'; EndTime = [datetime]'2024-01-08T10:00:00'; Subject = 'Weekly planning'; ItemType = 'Appointment'; AllDayEvent = $false; Location = 'Conference Room'; Organizer = 'Torey Page'; RequiredAttendees = 'Linda Artley; Sam Reed'; OptionalAttendees = 'Ava Stone'; IsRecurring = $true; RecurrenceSummary = 'Weekly; every 1; starting 2024-01-08; no end date'; Categories = 'Blue Category'; Sensitivity = 2; FolderPath = 'SamplePst\Calendar'; BodyText = 'Recurring planning agenda.'; Notes = 'Recurring planning agenda.'; MessageClass = 'IPM.Appointment'; EntryId = 'sample-calendar-1'; CreationTime = [datetime]'2024-01-02T08:15:00'; LastModificationTime = [datetime]'2024-01-03T11:20:00'; Attachments = @() }
        )
        Contacts = @(
            [pscustomobject]@{ DisplayName = 'Curriculum Team'; FullName = 'Curriculum Team'; FirstName = ''; MiddleName = ''; LastName = ''; CompanyName = 'Perfection Learning'; JobTitle = 'Distribution List'; Department = 'Curriculum'; Email1 = 'curriculum@example.com'; Email2 = ''; Email3 = ''; BusinessPhone = '555-1000'; HomePhone = ''; MobilePhone = ''; OtherPhone = ''; BusinessAddress = '100 Main St'; HomeAddress = ''; OtherAddress = ''; WebPage = 'https://example.com'; Birthday = $null; Anniversary = $null; Categories = 'Blue Category'; DistributionListMembers = @('Linda Artley', 'Sam Reed'); FolderPath = 'SamplePst\Contacts'; BodyText = 'Shared curriculum contacts.'; Notes = 'Shared curriculum contacts.'; MessageClass = 'IPM.DistList'; EntryId = 'sample-contact-1'; CreationTime = [datetime]'2024-01-03T09:45:00'; LastModificationTime = [datetime]'2024-01-04T10:15:00'; Attachments = @() }
        )
    }
}

function Write-HtmlReport {
    param(
        [Parameter(Mandatory = $true)][object]$Records,
        [Parameter(Mandatory = $true)]$PstItem,
        [Parameter(Mandatory = $true)][string]$ReportPath
    )

    Write-ReportLog 'Preparing HTML report data.'
    Write-ConversionStage -Stage 'PreparingReport'
    $sorted = @($Records | Sort-Object @{ Expression = { if ($_.SortTime) { ([datetime]$_.SortTime).Ticks } else { 0 } } }, ParticipantsKey, Subject, FolderPath)

    $participantSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($record in $sorted) {
        foreach ($participant in @($record.Participants)) {
            if (-not [string]::IsNullOrWhiteSpace($participant)) { [void]$participantSet.Add($participant) }
        }
    }
    $allParticipants = @($participantSet | Sort-Object)
    $likelyParticipants = @($allParticipants | Where-Object { Test-LikelyPersonName $_ })
    $otherParticipants = @($allParticipants | Where-Object { -not (Test-LikelyPersonName $_) })
    $likelyParticipantSet = @{}
    foreach ($participant in $likelyParticipants) { $likelyParticipantSet[$participant] = $true }

    $participantOptions = @(foreach ($participant in $likelyParticipants) { Format-ParticipantOptionHtml -Participant $participant -DefaultParticipants $DefaultConversationParticipants })
    $otherParticipantOptions = @(foreach ($participant in $otherParticipants) { Format-ParticipantOptionHtml -Participant $participant -DefaultParticipants $DefaultConversationParticipants })
    $folderSummaryRows = @($sorted | Group-Object FolderPath | Sort-Object Name | ForEach-Object { '<tr><td>{0}</td><td>{1}</td></tr>' -f (ConvertTo-HtmlEncodedText $_.Name), $_.Count })

    $groups = [ordered]@{}
    foreach ($record in $sorted) {
        if (-not $groups.Contains($record.ConversationKey)) {
            $groups[$record.ConversationKey] = New-Object System.Collections.Generic.List[object]
        }
        [void]$groups[$record.ConversationKey].Add($record)
    }

    $groupCount = [Math]::Max(1, $groups.Count)
    Write-ReportLog "Writing HTML report: 0 of $groupCount conversations."
    Write-ConversionStage -Stage 'WritingReport' -Extra ("Written=0|Total={0}" -f $groupCount)

    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    $writer = [System.IO.StreamWriter]::new($ReportPath, $false, $utf8NoBom)
    try {
        Write-ReportHeader -Writer $writer -PstItem $PstItem -SortedRecords $sorted -AllParticipants $allParticipants -OtherParticipants $otherParticipants -ParticipantOptions $participantOptions -OtherParticipantOptions $otherParticipantOptions -FolderSummaryRows $folderSummaryRows
        $senderClasses = @{}
        $nextSenderIndex = 0
        $writtenGroups = 0
        foreach ($key in $groups.Keys) {
            Write-ConversationHtml -Writer $writer -GroupMessages $groups[$key].ToArray() -PreferredParticipants $likelyParticipantSet -SenderClasses $senderClasses -NextSenderIndex ([ref]$nextSenderIndex)
            $writtenGroups++
            if (($writtenGroups -eq $groupCount) -or ($writtenGroups % 50 -eq 0)) {
                Write-ReportLog "HTML report progress: $writtenGroups of $groupCount conversations."
                Write-ConversionStage -Stage 'WritingReport' -Extra ("Written={0}|Total={1}" -f $writtenGroups, $groupCount)
            }
        }
        Write-ReportLog 'Finalizing HTML report.'
        Write-ConversionStage -Stage 'Finalizing'
        Write-ReportFooter -Writer $writer
    }
    finally {
        $writer.Dispose()
    }
}

# Task 4 will replace with full email UX.
function Write-EmailHtmlReport {
    param(
        [Parameter(Mandatory = $true)][object]$Records,
        [Parameter(Mandatory = $true)]$PstItem,
        [Parameter(Mandatory = $true)][string]$ReportPath
    )

    Write-ReportLog 'Preparing Email HTML report data.'
    Write-ConversionStage -Stage 'PreparingReport'
    $sorted = @($Records | Sort-Object @{ Expression = { if ($_.SortTime) { ([datetime]$_.SortTime).Ticks } else { 0 } } }, FolderPath, Subject, ParticipantsKey)

    $participantSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $folderSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($record in $sorted) {
        foreach ($participant in @($record.Participants)) {
            if (-not [string]::IsNullOrWhiteSpace($participant)) { [void]$participantSet.Add($participant) }
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$record.FolderPath)) { [void]$folderSet.Add([string]$record.FolderPath) }
    }

    $allParticipants = @($participantSet | Sort-Object)
    $allFolders = @($folderSet | Sort-Object)
    # ponytail: message counts (not conversation counts) for folder filter labels
    $folderMessageCounts = @{}
    foreach ($group in @($sorted | Group-Object FolderPath)) {
        $folderMessageCounts[[string]$group.Name] = $group.Count
    }
    $folderOptions = @(foreach ($folder in $allFolders) {
        $count = if ($folderMessageCounts.ContainsKey($folder)) { [int]$folderMessageCounts[$folder] } else { 0 }
        Format-FolderOptionHtml -FolderPath $folder -MessageCount $count
    })
    $folderSummaryRows = @(
        $sorted | Group-Object FolderPath | Sort-Object Name | ForEach-Object {
            $leaf = Get-EmailFolderLeafLabel -FolderPath $_.Name
            '<tr><td title="{0}">{1}</td><td>{2}</td></tr>' -f (ConvertTo-HtmlEncodedText $_.Name), (ConvertTo-HtmlEncodedText $leaf), $_.Count
        }
    )
    $topSenderOptions = @(
        $sorted |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.SenderDisplay) } |
            Group-Object SenderDisplay |
            Sort-Object Count -Descending |
            Select-Object -First 20 |
            ForEach-Object { Format-TopSenderOptionHtml -SenderDisplay $_.Name -MessageCount $_.Count }
    )

    $groups = [ordered]@{}
    foreach ($record in $sorted) {
        $key = Get-EmailThreadKey -Record $record
        if (-not $groups.Contains($key)) {
            $groups[$key] = New-Object System.Collections.Generic.List[object]
        }
        [void]$groups[$key].Add($record)
    }

    $groupCount = [Math]::Max(1, $groups.Count)
    Write-ReportLog "Writing Email HTML report: 0 of $groupCount conversations."
    Write-ConversionStage -Stage 'WritingReport' -Extra ("Written=0|Total={0}" -f $groupCount)

    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    $writer = [System.IO.StreamWriter]::new($ReportPath, $false, $utf8NoBom)
    try {
        Write-EmailReportHeader -Writer $writer -PstItem $PstItem -ConversationCount $groups.Count -SortedRecords $sorted -AllParticipants $allParticipants -AllFolders $allFolders -FolderOptions $folderOptions -FolderSummaryRows $folderSummaryRows -TopSenderOptions $topSenderOptions
        $senderClasses = @{}
        $nextSenderIndex = 0
        $writtenGroups = 0
        foreach ($key in $groups.Keys) {
            Write-EmailConversationHtml -Writer $writer -GroupMessages $groups[$key].ToArray() -SenderClasses $senderClasses -NextSenderIndex ([ref]$nextSenderIndex)
            $writtenGroups++
            if (($writtenGroups -eq $groupCount) -or ($writtenGroups % 50 -eq 0)) {
                Write-ReportLog "Email HTML report progress: $writtenGroups of $groupCount conversations."
                Write-ConversionStage -Stage 'WritingReport' -Extra ("Written={0}|Total={1}" -f $writtenGroups, $groupCount)
            }
        }
        Write-ReportLog 'Finalizing Email HTML report.'
        Write-ConversionStage -Stage 'Finalizing'
        Write-EmailReportFooter -Writer $writer
    }
    finally {
        $writer.Dispose()
    }
}

function ConvertTo-NormalizedFilterText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    return (($text -replace '\s+', ' ').Trim().ToLowerInvariant())
}

function Split-RecordCategories {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return @() }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }

    $items = @($text -split '\s*[;,]\s*' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($items.Count -gt 0) { return $items }
    return @($text.Trim())
}

function ConvertTo-JsonArrayAttributeValue {
    param([AllowNull()][string[]]$Values)

    $safeValues = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($safeValues.Count -eq 0) { return '[]' }

    $jsonItems = @($safeValues | ForEach-Object { ConvertTo-Json -InputObject ([string]$_) -Compress })
    return (ConvertTo-HtmlEncodedText ('[' + ($jsonItems -join ',') + ']'))
}

function Get-RecordValueHtml {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return "<span class='empty'>(none)</span>" }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "<span class='empty'>(none)</span>" }
    return (ConvertTo-HtmlEncodedText $text)
}

function Format-RecordDateTime {
    param([AllowNull()][object]$Value)

    if (Test-MissingDate $Value) { return '' }
    try { return ([datetime]$Value).ToString('MMM d, yyyy h:mm tt') } catch { return [string]$Value }
}

function Format-RecordDate {
    param([AllowNull()][object]$Value)

    if (Test-MissingDate $Value) { return '' }
    try { return ([datetime]$Value).ToString('MMM d, yyyy') } catch { return [string]$Value }
}

function Get-EffectiveFilterEndDate {
    param(
        [AllowNull()][object]$StartTime,
        [AllowNull()][object]$EndTime
    )

    if (Test-MissingDate $EndTime) {
        if (Test-MissingDate $StartTime) { return '' }
        return ([datetime]$StartTime).ToString('yyyy-MM-dd')
    }

    $effectiveEnd = [datetime]$EndTime
    if (-not (Test-MissingDate $StartTime)) {
        $startValue = [datetime]$StartTime
        if ($effectiveEnd -gt $startValue -and $effectiveEnd.TimeOfDay -eq [TimeSpan]::Zero) {
            $effectiveEnd = $effectiveEnd.AddDays(-1)
        }
    }

    return $effectiveEnd.ToString('yyyy-MM-dd')
}

function Get-CalendarSensitivityLabel {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    $intValue = $null
    try { $intValue = [int]$Value } catch { }
    switch ($intValue) {
        1 { return 'Personal (1)' }
        2 { return 'Private (2)' }
        3 { return 'Confidential (3)' }
        0 { return 'Normal (0)' }
        default {
            return [string]$Value
        }
    }
}

function New-DetailFieldHtml {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [AllowNull()][object]$Value
    )

    return "<div class='detail-field'><div class='detail-label'>$(ConvertTo-HtmlEncodedText $Label)</div><div class='detail-value'>$(Get-RecordValueHtml $Value)</div></div>"
}

function New-RecordListHtml {
    param([AllowNull()][string[]]$Values)

    $items = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($items.Count -eq 0) { return "<span class='empty'>(none)</span>" }

    $rows = @($items | ForEach-Object { "<li>$(ConvertTo-HtmlEncodedText $_)</li>" })
    return "<ul class='detail-list'>$($rows -join '')</ul>"
}

function Get-StaticAttachmentMetadataHtml {
    param([AllowNull()][object[]]$Attachments)

    $attachmentRows = @($Attachments)
    if ($attachmentRows.Count -eq 0) {
        return @"
<section class='record-section'>
  <h3>Attachment metadata</h3>
  <div class='detail-value'><span class='empty'>(none)</span></div>
</section>
"@
    }

    $rows = New-Object System.Collections.Generic.List[string]
    foreach ($attachment in $attachmentRows) {
        $rowText = '<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>' -f
            (ConvertTo-HtmlEncodedText (Get-PropSafe -Object $attachment -Name 'FileName' -Default '')),
            (ConvertTo-HtmlEncodedText (Get-PropSafe -Object $attachment -Name 'DisplayName' -Default '')),
            (ConvertTo-HtmlEncodedText (Get-PropSafe -Object $attachment -Name 'Size' -Default ''))
        [void]$rows.Add($rowText)
    }

    return @"
<section class='record-section'>
  <h3>Attachment metadata</h3>
  <table class='meta-table'>
    <thead><tr><th>File name</th><th>Display name</th><th>Size bytes</th></tr></thead>
    <tbody>
$($rows -join "`n")
    </tbody>
  </table>
</section>
"@
}

function Get-StaticRecordReportCss {
    $baseCss = Get-ReportCss
    return $baseCss + @'
@media (min-width: 901px) { .review-layout { grid-template-columns: minmax(260px, 340px) minmax(0, 1fr); } }
.filter-stack { display: grid; gap: 12px; }
.filter-grid { display: grid; grid-template-columns: 1fr; gap: 10px; }
.filter-grid label { display: grid; gap: 5px; font-weight: 700; color: #334155; font-size: .92rem; }
.filter-grid input[type="search"], .filter-grid input[type="date"], .filter-grid select { width: 100%; padding: 10px 11px; border: 1px solid #b9c4d3; border-radius: 10px; font-size: .95rem; background: white; }
.filter-actions { display: flex; gap: 8px; flex-wrap: wrap; }
.record-list { display: grid; gap: 14px; }
.record-card { padding: 16px 18px; content-visibility: auto; contain-intrinsic-size: auto 520px; }
.record-card[hidden] { display: none; }
.record-header { display: flex; justify-content: space-between; gap: 14px; align-items: flex-start; margin-bottom: 12px; }
.record-header h2 { margin: 0 0 4px; font-size: 1.08rem; }
.record-subtitle { color: var(--muted); font-size: .92rem; }
.badge-row { display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
.pill { display: inline-flex; align-items: center; border-radius: 999px; padding: 5px 10px; background: #eef3ff; color: #24408d; font-size: .84rem; font-weight: 800; border: 1px solid #c7d4ef; }
.details-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 10px; }
.detail-field { border: 1px solid #e5eaf2; border-radius: 12px; background: #fbfcff; padding: 10px 12px; }
.detail-label { color: var(--muted); font-size: .8rem; font-weight: 800; text-transform: uppercase; letter-spacing: .02em; margin-bottom: 4px; }
.detail-value { overflow-wrap: anywhere; }
.detail-list { margin: 0; padding-left: 18px; }
.record-section { margin-top: 14px; }
.record-section h3 { margin: 0 0 8px; font-size: .95rem; color: #243047; }
.body-block { background: white; border: 1px solid #e5eaf2; border-radius: 12px; padding: 12px; overflow-wrap: anywhere; }
.meta-table { margin-top: 0; }
.empty { color: var(--muted); font-style: italic; }
@media (max-width: 900px) { .record-header { display: block; } .badge-row { justify-content: flex-start; margin-top: 10px; } }
'@
}

function Get-CalendarReportCss {
    $baseCss = Get-StaticRecordReportCss
    $baseCss = [regex]::Replace($baseCss, '(?m)^\.review-layout \{[^\r\n]*10px[^\r\n]*\}\r?\n?', '')
    return $baseCss + @'
.calendar-shell { display: grid; grid-template-columns: 1fr; gap: 14px; align-items: start; }
.calendar-panel { background: #fff; border: 1px solid #d4dce6; border-radius: 16px; box-shadow: 0 10px 30px rgba(26, 35, 50, .08); overflow: hidden; }
.calendar-agenda, .calendar-detail { min-height: 280px; display: flex; flex-direction: column; }
.calendar-panel-heading { padding: 14px 16px 12px; border-bottom: 1px solid #d4dce6; background: linear-gradient(180deg, #f8fbfd, #fff); }
.calendar-panel-heading h2 { margin: 0 0 4px; font-size: .95rem; font-weight: 700; color: #1a2332; }
.calendar-panel-heading p { margin: 0; color: #5c6b7a; font-size: .75rem; }
.calendar-agenda-count { display: block; margin-top: 8px; color: #0d6e6e; font-size: .75rem; font-weight: 700; }
.calendar-agenda-list { flex: 1; min-height: 0; max-height: 68vh; overflow-y: auto; padding: 8px; }
.calendar-agenda-group { margin: 8px 4px 4px; color: #5c6b7a; font-size: .7rem; font-weight: 700; letter-spacing: .04em; text-transform: uppercase; }
.calendar-agenda-item { display: grid; grid-template-columns: 8px minmax(0, 1fr); gap: 10px; width: 100%; padding: 10px; border: 1px solid transparent; border-radius: 10px; background: transparent; color: #1a2332; text-align: left; white-space: normal; cursor: pointer; font: inherit; }
.calendar-agenda-item:hover { background: #f5f8fb; border-color: #d4dce6; }
.calendar-agenda-item.active { background: #e4f4f3; border-color: #9ecfcd; }
.calendar-agenda-rail { border-radius: 999px; min-height: 28px; align-self: stretch; background: #1f5fbf; }
.calendar-agenda-item.allday .calendar-agenda-rail { background: #8a5a12; }
.calendar-agenda-item.recurring .calendar-agenda-rail { background: #6b3fa0; }
.calendar-agenda-time { color: #5c6b7a; font-size: .7rem; font-weight: 600; }
.calendar-agenda-title { margin-top: 2px; font-size: .82rem; font-weight: 650; line-height: 1.3; }
.calendar-agenda-place { margin-top: 3px; color: #5c6b7a; font-size: .7rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.calendar-agenda-time, .calendar-agenda-title, .calendar-agenda-place { display: block; }
.calendar-agenda-empty { padding: 20px 12px; color: #5c6b7a; text-align: center; }
.calendar-toolbar { display: grid; gap: 12px; padding: 14px 16px; border-bottom: 1px solid #d4dce6; }
.calendar-nav { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
.calendar-nav h2 { min-width: 170px; margin: 0 4px; text-align: center; font-size: 1.45rem; font-weight: 700; color: #1a2332; }
.calendar-nav button { padding: 8px 11px; border: 2px solid #d4dce6; border-radius: 8px; }
.calendar-nav button#calendarToday { padding: 5px 9px; background: #0d6e6e; border-color: #0d6e6e; color: #fff; font-size: .78rem; font-weight: 600; }
.calendar-filter-grid { display: grid; grid-template-columns: repeat(4, minmax(120px, 1fr)); gap: 8px; }
.calendar-filter-grid label { display: grid; gap: 4px; color: #334155; font-size: .76rem; font-weight: 700; }
.calendar-filter-grid input, .calendar-filter-grid select { width: 100%; min-width: 0; padding: 8px 9px; border: 2px solid #b9c4d3; border-radius: 8px; background: white; font: inherit; }
.calendar-filter-actions { display: flex; align-items: end; }
.calendar-filter-actions button { width: 100%; border: 2px solid #d4dce6; border-radius: 8px; }
.calendar-legend { display: flex; flex-wrap: wrap; gap: 14px; padding: 10px 16px 4px; color: #5c6b7a; font-size: .75rem; }
.calendar-legend i { display: inline-block; width: 10px; height: 10px; margin-right: 6px; border-radius: 3px; }
.calendar-month { padding: 8px 10px 14px; }
.calendar-weekdays, .calendar-month-grid { display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); gap: 5px; }
.calendar-weekdays span { padding: 5px 0; color: #5c6b7a; font-size: .7rem; font-weight: 700; letter-spacing: .04em; text-align: center; text-transform: uppercase; }
.calendar-month-grid { min-height: 580px; }
.calendar-day { min-height: 98px; padding: 7px; border: 1px solid #d4dce6; border-radius: 10px; background: #fbfcfe; scroll-margin-top: 100px; transition: box-shadow 180ms ease, background 180ms ease; }
.calendar-day.outside { opacity: .45; background: #f3f5f8; }
.calendar-day.today { border-color: #0d6e6e; box-shadow: inset 0 0 0 1px #0d6e6e; background: #f3fafa; }
.calendar-day.flash { box-shadow: 0 0 0 3px rgba(13, 110, 110, .28); background: #eaf7f6; }
.calendar-day-number { margin-bottom: 2px; color: #5c6b7a; font-size: .75rem; font-weight: 700; }
.calendar-day.today .calendar-day-number { color: #0d6e6e; }
.calendar-event-chip { display: block; width: 100%; margin: 3px 0 0; padding: 4px 6px; border: 0; border-radius: 6px; background: #e8f0fc; color: #1f5fbf; font-size: .7rem; font-weight: 600; line-height: 1.25; overflow: hidden; text-align: left; text-overflow: ellipsis; white-space: nowrap; cursor: pointer; }
.calendar-event-chip.allday { background: #f7efd9; color: #8a5a12; }
.calendar-event-chip.recurring { background: #f1eaf8; color: #6b3fa0; }
.calendar-event-chip.active { outline: 2px solid currentColor; outline-offset: 1px; }
.calendar-more { width: 100%; padding: 2px 4px; border: 0; background: transparent; color: #5c6b7a; font-size: .7rem; text-align: left; cursor: pointer; }
.calendar-detail { max-height: calc(100vh - 16px); }
.calendar-detail-empty { padding: 28px 22px; color: #5c6b7a; min-height: 240px; }
.calendar-detail-empty strong { display: block; margin-bottom: 8px; color: #1a2332; font-size: 1rem; }
.calendar-detail-head { padding: 16px 18px 12px; border-bottom: 1px solid #d4dce6; background: linear-gradient(180deg, #f8fbfd, #fff); }
.calendar-detail-kind { margin-bottom: 8px; font-size: .7rem; font-weight: 700; letter-spacing: .05em; text-transform: uppercase; }
.calendar-detail-head h3 { margin: 0 0 8px; font-family: Georgia, "Times New Roman", serif; font-size: 1.3rem; line-height: 1.2; color: #1a2332; font-weight: 700; }
.calendar-detail-when { color: #5c6b7a; font-size: .88rem; }
.calendar-detail-body { flex: 1; min-height: 0; overflow-y: auto; padding: 14px 18px 20px; }
.calendar-detail-fields { margin: 0; }
.calendar-detail-field { display: grid; grid-template-columns: 100px minmax(0, 1fr); gap: 8px; padding: 8px 0; border-bottom: 1px solid #eef2f6; font-size: .82rem; }
.calendar-detail-field:last-child { border-bottom: 0; }
.calendar-detail-field dt { color: #5c6b7a; font-weight: 500; }
.calendar-detail-field dd { margin: 0; color: #1a2332; overflow-wrap: anywhere; }
.calendar-detail-section { margin-top: 14px; }
.calendar-detail-section h4 { margin: 0 0 8px; font-size: .82rem; color: #243047; }
.calendar-detail-section .body-block, .calendar-detail-section .meta-table { margin-top: 0; }
.calendar-record-store { display: none; }
.calendar-report-footer { margin: 18px 2px 8px; }
@media (min-width: 1181px) { .calendar-shell { grid-template-columns: 240px minmax(0, 1.35fr) 320px; } .calendar-agenda, .calendar-detail { position: sticky; top: 14px; max-height: calc(100vh - 28px); } }
@media (min-width: 821px) and (max-width: 1180px) { .calendar-shell { grid-template-columns: 220px minmax(0, 1fr); } .calendar-detail { grid-column: 1 / -1; max-height: none; } }
@media (max-width: 820px) { .calendar-filter-grid { grid-template-columns: 1fr 1fr; } .calendar-month-grid { min-height: 0; } .calendar-day { min-height: 84px; } }
@media (max-width: 560px) { .calendar-filter-grid { grid-template-columns: 1fr; } .calendar-weekdays, .calendar-month-grid { gap: 2px; } .calendar-day { min-height: 72px; padding: 4px; } .calendar-event-chip { font-size: .64rem; padding: 3px; } .calendar-detail-field { grid-template-columns: 1fr; gap: 2px; } }
'@
}

function Get-CalendarReportScript {
    return @'
(function () {
  const searchInput = document.getElementById('calendarSearch');
  const fromDateInput = document.getElementById('calendarFromDateFilter');
  const toDateInput = document.getElementById('calendarToDateFilter');
  const folderSelect = document.getElementById('calendarFolderFilter');
  const typeSelect = document.getElementById('calendarTypeFilter');
  const allDaySelect = document.getElementById('calendarAllDayFilter');
  const recurringSelect = document.getElementById('calendarRecurringFilter');
  const clearBtn = document.getElementById('calendarClearFiltersBtn');
  const visibleCount = document.getElementById('calendarVisibleCount');
  const agendaCount = document.getElementById('calendarAgendaCount');
  const agendaList = document.getElementById('calendarAgendaList');
  const monthGrid = document.getElementById('calendarMonthGrid');
  const monthLabel = document.getElementById('calendarMonthLabel');
  const detail = document.getElementById('calendarDetail');
  const previousMonth = document.getElementById('calendarPreviousMonth');
  const nextMonth = document.getElementById('calendarNextMonth');
  const todayButton = document.getElementById('calendarToday');
  const cards = Array.from(document.querySelectorAll('.record-card'));
  const totalCount = cards.length;
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const chipLimit = 3;
  let filteredCards = [];
  let currentMonth = null;
  let selectedEntryId = '';
  let flashTimer = null;
  let scheduled = false;

  cards.forEach(card => {
    card._search = (card.dataset.search || '').toLowerCase();
    card._startDate = card.dataset.startDate || '';
    card._endDate = card.dataset.endDate || '';
    card._folder = (card.dataset.folder || '').toLowerCase();
    card._itemType = (card.dataset.itemType || '').toLowerCase();
    card._allDay = (card.dataset.allDay || '').toLowerCase();
    card._recurring = (card.dataset.recurring || '').toLowerCase();
    card._entryId = card.dataset.entryId || '';
    card._subject = card.dataset.subject || '(no subject)';
    card._location = card.dataset.location || '';
    card._folderLabel = card.dataset.folderLabel || '';
    card._startTime = card.dataset.startTime || '';
    card._startDateTime = card.dataset.startDatetime || card._startDate;
    card._itemTypeLabel = card.dataset.itemTypeLabel || card.dataset.itemType || 'Appointment';
  });

  function chronologicalCards(list) {
    return list.slice().sort((a, b) => {
      const left = a._startDateTime || a._startDate || '';
      const right = b._startDateTime || b._startDate || '';
      if (left !== right) return left < right ? -1 : 1;
      return String(a._subject || '').localeCompare(String(b._subject || ''));
    });
  }

  function kindColor(card) {
    const kind = eventClass(card);
    if (kind === 'allday') return '#8a5a12';
    if (kind === 'recurring') return '#6b3fa0';
    return '#1f5fbf';
  }

  function fieldValueMap(card) {
    const values = {};
    card.querySelectorAll('.detail-field').forEach(field => {
      const labelNode = field.querySelector('.detail-label');
      const valueNode = field.querySelector('.detail-value');
      const label = labelNode ? String(labelNode.textContent).trim() : '';
      if (label && valueNode) values[label] = valueNode.innerHTML;
    });
    return values;
  }

  function updateVisibleCount(shown) {
    if (!visibleCount) return;
    visibleCount.textContent = shown + ' of ' + totalCount + ' records shown';
  }

  function parseDateKey(value) {
    const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(value || '');
    return match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])) : null;
  }

  function dateKey(date) {
    return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
  }

  function sameMonth(date, month) {
    return date && month && date.getFullYear() === month.getFullYear() && date.getMonth() === month.getMonth();
  }

  function earliestVisibleMonth() {
    const first = filteredCards.find(card => parseDateKey(card._startDate));
    const date = first ? parseDateKey(first._startDate) : new Date();
    return new Date(date.getFullYear(), date.getMonth(), 1);
  }

  function eventClass(card) {
    if (card._recurring === 'yes') return 'recurring';
    if (card._allDay === 'yes') return 'allday';
    return 'meeting';
  }

  function shortTime(card) {
    return card._allDay === 'yes' ? 'All day' : (card._startTime || 'Time unavailable');
  }

  function createAgendaItem(card) {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'calendar-agenda-item ' + eventClass(card);
    item.setAttribute('data-entry-id', card._entryId);
    item.addEventListener('click', () => selectEvent(card._entryId, 'agenda'));

    const rail = document.createElement('span');
    rail.className = 'calendar-agenda-rail';
    rail.setAttribute('aria-hidden', 'true');
    const content = document.createElement('span');
    const time = document.createElement('span');
    time.className = 'calendar-agenda-time';
    time.textContent = shortTime(card);
    const title = document.createElement('span');
    title.className = 'calendar-agenda-title';
    title.textContent = card._subject;
    const place = document.createElement('span');
    place.className = 'calendar-agenda-place';
    place.textContent = card._location || card._folderLabel || '(no location)';
    content.append(time, title, place);
    item.append(rail, content);
    return item;
  }

  function renderAgenda() {
    agendaList.replaceChildren();
    const agendaCards = chronologicalCards(filteredCards);
    if (agendaCount) agendaCount.textContent = agendaCards.length + ' items';
    if (!agendaCards.length) {
      const empty = document.createElement('div');
      empty.className = 'calendar-agenda-empty';
      empty.textContent = 'No matching scheduled meetings.';
      agendaList.appendChild(empty);
      return;
    }

    let lastDate = '';
    agendaCards.forEach(card => {
      if (card._startDate !== lastDate) {
        const date = parseDateKey(card._startDate);
        const group = document.createElement('div');
        group.className = 'calendar-agenda-group';
        group.textContent = date
          ? weekdayNames[date.getDay()] + ', ' + monthNames[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear()
          : 'Date unavailable';
        agendaList.appendChild(group);
        lastDate = card._startDate;
      }
      agendaList.appendChild(createAgendaItem(card));
    });
    syncSelectionClasses();
  }

  function createEventChip(card) {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'calendar-event-chip ' + eventClass(card);
    chip.setAttribute('data-entry-id', card._entryId);
    chip.textContent = (card._allDay === 'yes' ? '' : shortTime(card) + ' · ') + card._subject;
    chip.title = card._subject;
    chip.addEventListener('click', () => selectEvent(card._entryId, 'month'));
    return chip;
  }

  function renderMonth() {
    monthGrid.replaceChildren();
    monthLabel.textContent = monthNames[currentMonth.getMonth()] + ' ' + currentMonth.getFullYear();
    const first = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), 1);
    const gridStart = new Date(first.getFullYear(), first.getMonth(), 1 - first.getDay());
    const todayKey = dateKey(new Date());

    for (let offset = 0; offset < 42; offset += 1) {
      const date = new Date(gridStart.getFullYear(), gridStart.getMonth(), gridStart.getDate() + offset);
      const key = dateKey(date);
      const day = document.createElement('div');
      day.className = 'calendar-day' + (sameMonth(date, currentMonth) ? '' : ' outside') + (key === todayKey ? ' today' : '');
      day.dataset.date = key;
      const number = document.createElement('div');
      number.className = 'calendar-day-number';
      number.textContent = String(date.getDate());
      day.appendChild(number);

      const dayCards = filteredCards.filter(card => card._startDate === key);
      const shownCards = dayCards.slice(0, chipLimit);
      const selectedCard = dayCards.find(card => card._entryId === selectedEntryId);
      if (selectedCard && shownCards.indexOf(selectedCard) === -1 && shownCards.length === chipLimit) {
        shownCards[chipLimit - 1] = selectedCard;
      }
      shownCards.forEach(card => day.appendChild(createEventChip(card)));
      if (dayCards.length > chipLimit) {
        const more = document.createElement('button');
        more.type = 'button';
        more.className = 'calendar-more';
        more.textContent = '+' + (dayCards.length - chipLimit) + ' more';
        more.addEventListener('click', () => selectEvent(dayCards[chipLimit]._entryId, 'month'));
        day.appendChild(more);
      }
      monthGrid.appendChild(day);
    }
    syncSelectionClasses();
  }

  function syncSelectionClasses() {
    document.querySelectorAll('.calendar-agenda-item, .calendar-event-chip').forEach(item => {
      item.classList.toggle('active', item.dataset.entryId === selectedEntryId);
    });
  }

  function showDetail(card) {
    const values = fieldValueMap(card);
    const date = parseDateKey(card._startDate);
    const dateText = date
      ? weekdayNames[date.getDay()] + ', ' + monthNames[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear()
      : '';
    const startText = values.Start || shortTime(card);
    const endText = values.End || '';
    const whenText = [dateText, startText && endText ? (startText + ' – ' + endText) : (startText || endText)]
      .filter(Boolean)
      .join(' · ');
    const preferredLabels = [
      'Location', 'Organizer', 'Required attendees', 'Optional attendees',
      'Recurrence summary', 'Categories', 'Sensitivity', 'Folder',
      'All-day', 'Created', 'Modified'
    ];
    const fieldHtml = preferredLabels
      .filter(label => {
        if (!Object.prototype.hasOwnProperty.call(values, label)) return false;
        if (label === 'All-day' && card._allDay !== 'yes') return false;
        return true;
      })
      .map(label => "<div class='calendar-detail-field'><dt>" + label + "</dt><dd>" + values[label] + "</dd></div>")
      .join('');
    let extras = '';
    Array.from(card.querySelectorAll('.record-section')).forEach(section => {
      const heading = section.querySelector('h3');
      let title = '';
      if (heading) title = String(heading.textContent).trim();
      if (!title) return;
      const clone = section.cloneNode(true);
      const cloneHeading = clone.querySelector('h3');
      if (cloneHeading) cloneHeading.remove();
      extras += "<section class='calendar-detail-section'><h4>" + title + "</h4>" + clone.innerHTML + "</section>";
    });

    detail.innerHTML =
      "<div class='calendar-detail-head'>" +
        "<div class='calendar-detail-kind' style='color:" + kindColor(card) + "'>" + (card._itemTypeLabel || 'Appointment') + "</div>" +
        "<h3></h3>" +
        "<div class='calendar-detail-when'></div>" +
      "</div>" +
      "<div class='calendar-detail-body'>" +
        "<dl class='calendar-detail-fields'>" + fieldHtml + "</dl>" +
        extras +
      "</div>";
    detail.querySelector('h3').textContent = card._subject;
    detail.querySelector('.calendar-detail-when').textContent = whenText || 'Time unavailable';
  }

  function findByEntryId(selector, entryId) {
    return Array.from(document.querySelectorAll(selector)).find(item => item.dataset.entryId === entryId);
  }

  function selectEvent(entryId, source) {
    const card = cards.find(item => item._entryId === entryId);
    if (!card) return;
    selectedEntryId = entryId;
    const eventDate = parseDateKey(card._startDate);
    if (eventDate && !sameMonth(eventDate, currentMonth)) {
      currentMonth = new Date(eventDate.getFullYear(), eventDate.getMonth(), 1);
    }
    renderMonth();
    showDetail(card);
    syncSelectionClasses();

    if (source === 'agenda') {
      const day = Array.from(monthGrid.querySelectorAll('.calendar-day')).find(cell => cell.dataset.date === card._startDate);
      if (day) {
        day.scrollIntoView({ behavior: 'smooth', block: 'center' });
        day.classList.remove('flash');
        void day.offsetWidth;
        day.classList.add('flash');
        if (flashTimer) clearTimeout(flashTimer);
        flashTimer = setTimeout(() => day.classList.remove('flash'), 900);
      }
    } else {
      const agendaItem = findByEntryId('.calendar-agenda-item', entryId);
      if (agendaItem) agendaItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }

  function matchesDateRange(card, fromDate, toDate) {
    if (fromDate && card._endDate && card._endDate < fromDate) return false;
    if (toDate && card._startDate && card._startDate > toDate) return false;
    return true;
  }

  function applyFilters() {
    const query = searchInput ? (searchInput.value || '').trim().toLowerCase() : '';
    const fromDate = fromDateInput ? (fromDateInput.value || '') : '';
    const toDate = toDateInput ? (toDateInput.value || '') : '';
    const folder = folderSelect ? (folderSelect.value || '').toLowerCase() : '';
    const itemType = typeSelect ? (typeSelect.value || '').toLowerCase() : '';
    const allDay = allDaySelect ? (allDaySelect.value || '') : '';
    const recurring = recurringSelect ? (recurringSelect.value || '') : '';
    filteredCards = cards.filter(card => {
      return (!query || card._search.indexOf(query) !== -1)
        && (!folder || card._folder === folder)
        && (!itemType || card._itemType === itemType)
        && (!allDay || card._allDay === allDay)
        && (!recurring || card._recurring === recurring)
        && matchesDateRange(card, fromDate, toDate);
    });

    cards.forEach(card => { card.hidden = filteredCards.indexOf(card) === -1; });
    updateVisibleCount(filteredCards.length);
    if (!currentMonth || (filteredCards.length && !filteredCards.some(card => sameMonth(parseDateKey(card._startDate), currentMonth)))) {
      currentMonth = earliestVisibleMonth();
    }
    if (selectedEntryId && !filteredCards.some(card => card._entryId === selectedEntryId)) {
      selectedEntryId = '';
      detail.innerHTML = "<div class='calendar-detail-empty'><strong>Select an appointment</strong>Use the agenda or a month chip to view all exported fields.</div>";
    }
    renderAgenda();
    renderMonth();
    if (!selectedEntryId && filteredCards.length) {
      selectEvent(chronologicalCards(filteredCards)[0]._entryId, 'filter');
    }
  }

  function scheduleApply() {
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => {
      scheduled = false;
      applyFilters();
    });
  }

  function resetFilters() {
    if (searchInput) searchInput.value = '';
    if (fromDateInput) fromDateInput.value = '';
    if (toDateInput) toDateInput.value = '';
    if (folderSelect) folderSelect.value = '';
    if (typeSelect) typeSelect.value = '';
    if (allDaySelect) allDaySelect.value = '';
    if (recurringSelect) recurringSelect.value = '';
    applyFilters();
  }

  if (searchInput) searchInput.addEventListener('input', scheduleApply);
  if (fromDateInput) fromDateInput.addEventListener('change', applyFilters);
  if (toDateInput) toDateInput.addEventListener('change', applyFilters);
  if (folderSelect) folderSelect.addEventListener('change', applyFilters);
  if (typeSelect) typeSelect.addEventListener('change', applyFilters);
  if (allDaySelect) allDaySelect.addEventListener('change', applyFilters);
  if (recurringSelect) recurringSelect.addEventListener('change', applyFilters);
  if (clearBtn) clearBtn.addEventListener('click', resetFilters);
  if (previousMonth) previousMonth.addEventListener('click', () => {
    currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() - 1, 1);
    renderAgenda();
    renderMonth();
  });
  if (nextMonth) nextMonth.addEventListener('click', () => {
    currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1);
    renderAgenda();
    renderMonth();
  });
  if (todayButton) todayButton.addEventListener('click', () => {
    const today = new Date();
    currentMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    renderAgenda();
    renderMonth();
  });

  requestAnimationFrame(applyFilters);
})();
'@
}

function Get-ContactsReportScript {
    return @'
(function () {
  const searchInput = document.getElementById('contactsSearch');
  const folderSelect = document.getElementById('contactsFolderFilter');
  const categorySelect = document.getElementById('contactsCategoryFilter');
  const clearBtn = document.getElementById('contactsClearFiltersBtn');
  const visibleCount = document.getElementById('contactsVisibleCount');
  const cards = Array.from(document.querySelectorAll('.record-card'));
  const totalCount = cards.length;
  let scheduled = false;

  function parseList(value) {
    if (!value) return [];
    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? parsed : [];
    } catch (_) {
      return [];
    }
  }

  cards.forEach(card => {
    card._search = (card.dataset.search || '').toLowerCase();
    card._folder = (card.dataset.folder || '').toLowerCase();
    card._categories = parseList(card.dataset.categories).map(v => String(v).toLowerCase());
  });

  function updateVisibleCount(shown) {
    if (!visibleCount) return;
    visibleCount.textContent = shown + ' of ' + totalCount + ' records shown';
  }

  function applyFilters() {
    const query = searchInput ? (searchInput.value || '').trim().toLowerCase() : '';
    const folder = folderSelect ? (folderSelect.value || '').toLowerCase() : '';
    const category = categorySelect ? (categorySelect.value || '').toLowerCase() : '';
    let shown = 0;

    cards.forEach(card => {
      const show = (!query || card._search.indexOf(query) !== -1)
        && (!folder || card._folder === folder)
        && (!category || card._categories.indexOf(category) !== -1);
      card.hidden = !show;
      if (show) shown += 1;
    });

    updateVisibleCount(shown);
  }

  function scheduleApply() {
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => {
      scheduled = false;
      applyFilters();
    });
  }

  function resetFilters() {
    if (searchInput) searchInput.value = '';
    if (folderSelect) folderSelect.value = '';
    if (categorySelect) categorySelect.value = '';
    applyFilters();
  }

  if (searchInput) searchInput.addEventListener('input', scheduleApply);
  if (folderSelect) folderSelect.addEventListener('change', applyFilters);
  if (categorySelect) categorySelect.addEventListener('change', applyFilters);
  if (clearBtn) clearBtn.addEventListener('click', resetFilters);

  requestAnimationFrame(applyFilters);
})();
'@
}

function Write-CalendarReportHeader {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)]$PstItem,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SortedRecords,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$FolderOptions,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$TypeOptions,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss K'
    $css = Get-CalendarReportCss
    $Writer.WriteLine(@"
<!doctype html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Purview Calendar PST Report - $(ConvertTo-HtmlEncodedText $PstItem.Name)</title>
<style>
$css
</style>
</head>
<body>
<div class='page'>
  <header class='hero'>
    <h1>Microsoft Purview eDiscovery Calendar Report</h1>
    <p>Static offline month grid, chronological agenda, and appointment detail with client-side search and filters.</p>
    <div class='hero-credit'>By Patrick Bush</div>
  </header>

  <section class='summary-grid' aria-label='Calendar report summary'>
    <div class='summary-card'><div class='label'>PST</div><div class='value'>$(ConvertTo-HtmlEncodedText $PstItem.Name)</div></div>
    <div class='summary-card'><div class='label'>Generated</div><div class='value'>$(ConvertTo-HtmlEncodedText $generated)</div></div>
    <div class='summary-card'><div class='label'>Calendar records</div><div class='value'>$(ConvertTo-HtmlEncodedText $SortedRecords.Count)</div></div>
    <div class='summary-card'><div class='label'>Folders detected</div><div class='value'>$(ConvertTo-HtmlEncodedText $FolderOptions.Count)</div></div>
    <div class='summary-card'><div class='label'>Item types</div><div class='value'>$(ConvertTo-HtmlEncodedText $TypeOptions.Count)</div></div>
    <div class='summary-card'><div class='label'>Recurring items</div><div class='value'>$(ConvertTo-HtmlEncodedText (@($SortedRecords | Where-Object { $_.IsRecurring }).Count))</div></div>
  </section>

  <div class='calendar-shell'>
    <aside id='calendarAgenda' class='calendar-panel calendar-agenda' aria-label='Calendar agenda'>
      <div class='calendar-panel-heading'>
        <h2>Scheduled meetings</h2>
        <p>All matching meetings in chronological order</p>
        <span id='calendarAgendaCount' class='calendar-agenda-count'></span>
      </div>
      <div id='calendarAgendaList' class='calendar-agenda-list'></div>
    </aside>

    <main class='calendar-panel calendar-month-panel' aria-label='Calendar month'>
      <section class='calendar-toolbar' aria-label='Calendar navigation and filters'>
        <div class='calendar-nav'>
          <button type='button' class='secondary' id='calendarPreviousMonth' aria-label='Previous month'>&lsaquo;</button>
          <h2 id='calendarMonthLabel'>Calendar</h2>
          <button type='button' class='secondary' id='calendarNextMonth' aria-label='Next month'>&rsaquo;</button>
          <button type='button' id='calendarToday'>Today</button>
          <span id='calendarVisibleCount' class='result-count'></span>
        </div>
        <div class='calendar-filter-grid'>
          <label for='calendarSearch'>Search
            <input id='calendarSearch' type='search' placeholder='Subject, people, notes, folder, ID' autocomplete='off'/>
          </label>
          <label for='calendarFromDateFilter'>From date
            <input id='calendarFromDateFilter' type='date'/>
          </label>
          <label for='calendarToDateFilter'>To date
            <input id='calendarToDateFilter' type='date'/>
          </label>
          <label for='calendarFolderFilter'>Folder
            <select id='calendarFolderFilter'>
              <option value=''>All folders</option>
$($FolderOptions -join "`n")
            </select>
          </label>
          <label for='calendarTypeFilter'>Type
            <select id='calendarTypeFilter'>
              <option value=''>All types</option>
$($TypeOptions -join "`n")
            </select>
          </label>
          <label for='calendarAllDayFilter'>All-day
            <select id='calendarAllDayFilter'>
              <option value=''>Any</option>
              <option value='yes'>Yes</option>
              <option value='no'>No</option>
            </select>
          </label>
          <label for='calendarRecurringFilter'>Recurring
            <select id='calendarRecurringFilter'>
              <option value=''>Any</option>
              <option value='yes'>Yes</option>
              <option value='no'>No</option>
            </select>
          </label>
          <div class='calendar-filter-actions'>
            <button type='button' class='secondary' id='calendarClearFiltersBtn'>Clear filters</button>
          </div>
        </div>
      </section>
      <div class='calendar-legend' aria-label='Calendar item legend'>
        <span><i style='background:#1f5fbf'></i>Meeting</span>
        <span><i style='background:#8a5a12'></i>All-day</span>
        <span><i style='background:#6b3fa0'></i>Recurring</span>
      </div>
      <section class='calendar-month'>
        <div class='calendar-weekdays' aria-hidden='true'>
          <span>Sun</span><span>Mon</span><span>Tue</span><span>Wed</span><span>Thu</span><span>Fri</span><span>Sat</span>
        </div>
        <div id='calendarMonthGrid' class='calendar-month-grid'></div>
      </section>
    </main>

    <aside id='calendarDetail' class='calendar-panel calendar-detail' aria-label='Selected calendar record'>
      <div class='calendar-detail-empty'><strong>Select an appointment</strong>Use the agenda or a month chip to view all exported fields.</div>
    </aside>
  </div>
  <div class='footer calendar-report-footer'>Log file: $(ConvertTo-HtmlEncodedText $LogPath)<br/>Created by Convert-PurviewTeamsPstToHtml.ps1</div>
  <section id='calendarRecordList' class='calendar-record-store' aria-hidden='true'>
"@)
}

function Write-ContactsReportHeader {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)]$PstItem,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$SortedRecords,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$FolderOptions,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$CategoryOptions,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss K'
    $css = Get-StaticRecordReportCss
    $Writer.WriteLine(@"
<!doctype html>
<html lang='en'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Purview Contacts PST Report - $(ConvertTo-HtmlEncodedText $PstItem.Name)</title>
<style>
$css
</style>
</head>
<body>
<div class='page'>
  <header class='hero'>
    <h1>Microsoft Purview eDiscovery Contacts Report</h1>
    <p>Static offline review of contacts and distribution lists with client-side search, folder filtering, and category filtering.</p>
    <div class='hero-credit'>By Patrick Bush</div>
  </header>

  <section class='summary-grid' aria-label='Contacts report summary'>
    <div class='summary-card'><div class='label'>PST</div><div class='value'>$(ConvertTo-HtmlEncodedText $PstItem.Name)</div></div>
    <div class='summary-card'><div class='label'>Generated</div><div class='value'>$(ConvertTo-HtmlEncodedText $generated)</div></div>
    <div class='summary-card'><div class='label'>Contact records</div><div class='value'>$(ConvertTo-HtmlEncodedText $SortedRecords.Count)</div></div>
    <div class='summary-card'><div class='label'>Folders detected</div><div class='value'>$(ConvertTo-HtmlEncodedText $FolderOptions.Count)</div></div>
    <div class='summary-card'><div class='label'>Categories detected</div><div class='value'>$(ConvertTo-HtmlEncodedText $CategoryOptions.Count)</div></div>
    <div class='summary-card'><div class='label'>Distribution lists</div><div class='value'>$(ConvertTo-HtmlEncodedText (@($SortedRecords | Where-Object { @($_.DistributionListMembers).Count -gt 0 }).Count))</div></div>
  </section>

  <div class='review-layout'>
    <aside class='filter-panel' aria-label='Contacts filters'>
      <div class='filter-title'>
        <h2>Filter contacts</h2>
        <span id='contactsVisibleCount' class='result-count'></span>
      </div>
      <p class='filter-help'>Search is precomputed from normalized contact names, organizations, addresses, notes, and identifiers so filtering stays fast even on large exports.</p>
      <div class='filter-stack'>
        <div class='filter-grid'>
          <label for='contactsSearch'>Search
            <input id='contactsSearch' type='search' placeholder='Names, organization, email, notes, ID' autocomplete='off'/>
          </label>
          <label for='contactsFolderFilter'>Folder
            <select id='contactsFolderFilter'>
              <option value=''>All folders</option>
$($FolderOptions -join "`n")
            </select>
          </label>
          <label for='contactsCategoryFilter'>Category
            <select id='contactsCategoryFilter'>
              <option value=''>All categories</option>
$($CategoryOptions -join "`n")
            </select>
          </label>
        </div>
        <div class='filter-actions'>
          <button type='button' class='secondary' id='contactsClearFiltersBtn'>Clear filters</button>
        </div>
      </div>
      <div class='footer'>Log file: $(ConvertTo-HtmlEncodedText $LogPath)<br/>Created by Convert-PurviewTeamsPstToHtml.ps1</div>
    </aside>

    <main class='conversation-pane' aria-label='Contact records'>
      <section id='contactsRecordList' class='record-list'>
"@)
}

function Write-StaticRecordReportFooter {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)][string]$ScriptText
    )

    $Writer.WriteLine(@"
      </section>
    </main>
  </div>
</div>
<script>
$ScriptText
</script>
</body>
</html>
"@)
}

function Write-CalendarReportFooter {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)][string]$ScriptText
    )

    $Writer.WriteLine(@"
  </section>
</div>
<script>
$ScriptText
</script>
</body>
</html>
"@)
}

function Write-CalendarRecordHtml {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)]$Record,
        [int]$RecordIndex = 0
    )

    $subject = if ([string]::IsNullOrWhiteSpace([string]$Record.Subject)) { '(no subject)' } else { [string]$Record.Subject }
    $startDisplay = Format-RecordDateTime $Record.StartTime
    $endDisplay = Format-RecordDateTime $Record.EndTime
    $startDate = if (Test-MissingDate $Record.StartTime) { '' } else { ([datetime]$Record.StartTime).ToString('yyyy-MM-dd') }
    $endDate = Get-EffectiveFilterEndDate -StartTime $Record.StartTime -EndTime $Record.EndTime
    $typeText = if ([string]::IsNullOrWhiteSpace([string]$Record.ItemType)) { 'Unknown' } else { [string]$Record.ItemType }
    $allDay = if ($Record.AllDayEvent) { 'yes' } else { 'no' }
    $recurring = if ($Record.IsRecurring) { 'yes' } else { 'no' }
    $sensitivityText = Get-CalendarSensitivityLabel $Record.Sensitivity
    $folderText = [string](Get-PropSafe -Object $Record -Name 'FolderPath' -Default '')
    $entryIdText = [string](Get-PropSafe -Object $Record -Name 'EntryId' -Default '')
    $stableEntryId = if ([string]::IsNullOrWhiteSpace($entryIdText)) { "calendar-record-$RecordIndex" } else { $entryIdText }
    $locationText = [string](Get-PropSafe -Object $Record -Name 'Location' -Default '')
    $startTimeText = if (Test-MissingDate $Record.StartTime) { '' } elseif ($Record.AllDayEvent) { 'All day' } else { ([datetime]$Record.StartTime).ToString('h:mm tt') }
    $startDateTime = if (Test-MissingDate $Record.StartTime) { '' } else { ([datetime]$Record.StartTime).ToString('yyyy-MM-ddTHH:mm:ss') }
    $searchParts = @(
        $subject, $typeText, $startDisplay, $endDisplay, $Record.Location, $Record.Organizer,
        $Record.RequiredAttendees, $Record.OptionalAttendees, $Record.RecurrenceSummary,
        $Record.Categories, $sensitivityText, $folderText, $Record.Notes, $Record.BodyText,
        $Record.MessageClass, $Record.EntryId
    )
    foreach ($attachment in @($Record.Attachments)) {
        $searchParts += @(
            (Get-PropSafe -Object $attachment -Name 'FileName' -Default ''),
            (Get-PropSafe -Object $attachment -Name 'DisplayName' -Default ''),
            (Get-PropSafe -Object $attachment -Name 'Size' -Default '')
        )
    }
    $searchText = ConvertTo-NormalizedFilterText ($searchParts -join ' ')
    $subtitleParts = @($startDisplay, $endDisplay, $folderText | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $subtitle = if ($subtitleParts.Count -gt 0) { $subtitleParts -join ' | ' } else { '' }
    $Writer.WriteLine(@"
<article class='conversation record-card' data-entry-id='$(ConvertTo-HtmlEncodedText $stableEntryId)' data-search='$(ConvertTo-HtmlEncodedText $searchText)' data-start-date='$(ConvertTo-HtmlEncodedText $startDate)' data-start-datetime='$(ConvertTo-HtmlEncodedText $startDateTime)' data-start-time='$(ConvertTo-HtmlEncodedText $startTimeText)' data-end-date='$(ConvertTo-HtmlEncodedText $endDate)' data-folder='$(ConvertTo-HtmlEncodedText (ConvertTo-NormalizedFilterText $folderText))' data-folder-label='$(ConvertTo-HtmlEncodedText $folderText)' data-item-type='$(ConvertTo-HtmlEncodedText (ConvertTo-NormalizedFilterText $typeText))' data-item-type-label='$(ConvertTo-HtmlEncodedText $typeText)' data-all-day='$(ConvertTo-HtmlEncodedText $allDay)' data-recurring='$(ConvertTo-HtmlEncodedText $recurring)' data-subject='$(ConvertTo-HtmlEncodedText $subject)' data-location='$(ConvertTo-HtmlEncodedText $locationText)'>
  <header class='record-header'>
    <div>
      <h2>$(ConvertTo-HtmlEncodedText $subject)</h2>
      <div class='record-subtitle'>$(Get-RecordValueHtml $subtitle)</div>
    </div>
    <div class='badge-row'>
      <span class='pill'>$(ConvertTo-HtmlEncodedText $typeText)</span>
      <span class='pill'>All-day: $(ConvertTo-HtmlEncodedText $allDay)</span>
      <span class='pill'>Recurring: $(ConvertTo-HtmlEncodedText $recurring)</span>
    </div>
  </header>
  <section class='details-grid'>
    $(New-DetailFieldHtml -Label 'Subject' -Value $subject)
    $(New-DetailFieldHtml -Label 'Item type' -Value $typeText)
    $(New-DetailFieldHtml -Label 'Start' -Value $startDisplay)
    $(New-DetailFieldHtml -Label 'End' -Value $endDisplay)
    $(New-DetailFieldHtml -Label 'All-day' -Value $allDay)
    $(New-DetailFieldHtml -Label 'Location' -Value $Record.Location)
    $(New-DetailFieldHtml -Label 'Organizer' -Value $Record.Organizer)
    $(New-DetailFieldHtml -Label 'Required attendees' -Value $Record.RequiredAttendees)
    $(New-DetailFieldHtml -Label 'Optional attendees' -Value $Record.OptionalAttendees)
    $(New-DetailFieldHtml -Label 'Recurrence summary' -Value $Record.RecurrenceSummary)
    $(New-DetailFieldHtml -Label 'Categories' -Value $Record.Categories)
    $(New-DetailFieldHtml -Label 'Sensitivity' -Value $sensitivityText)
    $(New-DetailFieldHtml -Label 'Folder' -Value $folderText)
    $(New-DetailFieldHtml -Label 'Created' -Value (Format-RecordDateTime $Record.CreationTime))
    $(New-DetailFieldHtml -Label 'Modified' -Value (Format-RecordDateTime $Record.LastModificationTime))
    $(New-DetailFieldHtml -Label 'Message class' -Value $Record.MessageClass)
    $(New-DetailFieldHtml -Label 'Entry ID' -Value $Record.EntryId)
  </section>
  $(Get-StaticAttachmentMetadataHtml -Attachments @($Record.Attachments))
  <section class='record-section'>
    <h3>Notes / body</h3>
    <div class='body-block'>$(ConvertTo-HtmlBody $Record.Notes)</div>
  </section>
</article>
"@)
}

function Write-ContactRecordHtml {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)]$Record
    )

    $displayName = if ([string]::IsNullOrWhiteSpace([string]$Record.DisplayName)) { '(no display name)' } else { [string]$Record.DisplayName }
    $folderText = [string](Get-PropSafe -Object $Record -Name 'FolderPath' -Default '')
    $categoryValues = @(Split-RecordCategories $Record.Categories)
    $categoryKeys = @($categoryValues | ForEach-Object { ConvertTo-NormalizedFilterText $_ })
    $searchParts = @(
        $Record.DisplayName, $Record.FullName, $Record.FirstName, $Record.MiddleName, $Record.LastName,
        $Record.CompanyName, $Record.JobTitle, $Record.Department, $Record.Email1, $Record.Email2, $Record.Email3,
        $Record.BusinessPhone, $Record.HomePhone, $Record.MobilePhone, $Record.OtherPhone,
        $Record.BusinessAddress, $Record.HomeAddress, $Record.OtherAddress, $Record.WebPage,
        (Format-RecordDate $Record.Birthday), (Format-RecordDate $Record.Anniversary),
        ($categoryValues -join ' '), (@($Record.DistributionListMembers) -join ' '),
        $folderText, $Record.Notes, $Record.BodyText, $Record.MessageClass, $Record.EntryId
    )
    foreach ($attachment in @($Record.Attachments)) {
        $searchParts += @(
            (Get-PropSafe -Object $attachment -Name 'FileName' -Default ''),
            (Get-PropSafe -Object $attachment -Name 'DisplayName' -Default ''),
            (Get-PropSafe -Object $attachment -Name 'Size' -Default '')
        )
    }
    $searchText = ConvertTo-NormalizedFilterText ($searchParts -join ' ')
    $subtitleParts = @($Record.CompanyName, $Record.JobTitle, $Record.Department | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $subtitle = if ($subtitleParts.Count -gt 0) { $subtitleParts -join ' | ' } else { '' }
    $Writer.WriteLine(@"
<article class='conversation record-card' data-search='$(ConvertTo-HtmlEncodedText $searchText)' data-folder='$(ConvertTo-HtmlEncodedText (ConvertTo-NormalizedFilterText $folderText))' data-categories='$(ConvertTo-JsonArrayAttributeValue $categoryKeys)'>
  <header class='record-header'>
    <div>
      <h2>$(ConvertTo-HtmlEncodedText $displayName)</h2>
      <div class='record-subtitle'>$(Get-RecordValueHtml $subtitle)</div>
    </div>
    <div class='badge-row'>
      <span class='pill'>$(ConvertTo-HtmlEncodedText ([string](Get-PropSafe -Object $Record -Name 'MessageClass' -Default 'Contact')))</span>
      <span class='pill'>$(ConvertTo-HtmlEncodedText $folderText)</span>
    </div>
  </header>
  <section class='details-grid'>
    $(New-DetailFieldHtml -Label 'Display name' -Value $Record.DisplayName)
    $(New-DetailFieldHtml -Label 'Full name' -Value $Record.FullName)
    $(New-DetailFieldHtml -Label 'First name' -Value $Record.FirstName)
    $(New-DetailFieldHtml -Label 'Middle name' -Value $Record.MiddleName)
    $(New-DetailFieldHtml -Label 'Last name' -Value $Record.LastName)
    $(New-DetailFieldHtml -Label 'Organization' -Value $Record.CompanyName)
    $(New-DetailFieldHtml -Label 'Job title' -Value $Record.JobTitle)
    $(New-DetailFieldHtml -Label 'Department' -Value $Record.Department)
    $(New-DetailFieldHtml -Label 'Email 1' -Value $Record.Email1)
    $(New-DetailFieldHtml -Label 'Email 2' -Value $Record.Email2)
    $(New-DetailFieldHtml -Label 'Email 3' -Value $Record.Email3)
    $(New-DetailFieldHtml -Label 'Business phone' -Value $Record.BusinessPhone)
    $(New-DetailFieldHtml -Label 'Home phone' -Value $Record.HomePhone)
    $(New-DetailFieldHtml -Label 'Mobile phone' -Value $Record.MobilePhone)
    $(New-DetailFieldHtml -Label 'Other phone' -Value $Record.OtherPhone)
    $(New-DetailFieldHtml -Label 'Business address' -Value $Record.BusinessAddress)
    $(New-DetailFieldHtml -Label 'Home address' -Value $Record.HomeAddress)
    $(New-DetailFieldHtml -Label 'Other address' -Value $Record.OtherAddress)
    $(New-DetailFieldHtml -Label 'Website' -Value $Record.WebPage)
    $(New-DetailFieldHtml -Label 'Birthday' -Value (Format-RecordDate $Record.Birthday))
    $(New-DetailFieldHtml -Label 'Anniversary' -Value (Format-RecordDate $Record.Anniversary))
    $(New-DetailFieldHtml -Label 'Categories' -Value (($categoryValues -join ', ')))
    $(New-DetailFieldHtml -Label 'Folder' -Value $folderText)
    $(New-DetailFieldHtml -Label 'Created' -Value (Format-RecordDateTime $Record.CreationTime))
    $(New-DetailFieldHtml -Label 'Modified' -Value (Format-RecordDateTime $Record.LastModificationTime))
    $(New-DetailFieldHtml -Label 'Message class' -Value $Record.MessageClass)
    $(New-DetailFieldHtml -Label 'Entry ID' -Value $Record.EntryId)
  </section>
  <section class='record-section'>
    <h3>Distribution list members</h3>
    $(New-RecordListHtml -Values @($Record.DistributionListMembers))
  </section>
  $(Get-StaticAttachmentMetadataHtml -Attachments @($Record.Attachments))
  <section class='record-section'>
    <h3>Notes / body</h3>
    <div class='body-block'>$(ConvertTo-HtmlBody $Record.Notes)</div>
  </section>
</article>
"@)
}

function Write-CalendarHtmlReport {
    param(
        [Parameter(Mandatory = $true)][object]$Records,
        [Parameter(Mandatory = $true)]$PstItem,
        [Parameter(Mandatory = $true)][string]$ReportPath,
        [string]$LogPath = $script:LogPath
    )

    Write-ReportLog 'Preparing Calendar HTML report data.'
    Write-ConversionStage -Stage 'PreparingReport'
    $sorted = @($Records | Sort-Object @{ Expression = { if ($_.StartTime) { ([datetime]$_.StartTime).Ticks } else { 0 } } }, Subject, FolderPath)
    $folderMap = [ordered]@{}
    $typeMap = [ordered]@{}
    foreach ($record in $sorted) {
        $folderText = [string](Get-PropSafe -Object $record -Name 'FolderPath' -Default '')
        $folderKey = ConvertTo-NormalizedFilterText $folderText
        if (-not [string]::IsNullOrWhiteSpace($folderKey) -and -not $folderMap.Contains($folderKey)) { $folderMap[$folderKey] = $folderText }
        $typeText = [string](Get-PropSafe -Object $record -Name 'ItemType' -Default '')
        $typeKey = ConvertTo-NormalizedFilterText $typeText
        if (-not [string]::IsNullOrWhiteSpace($typeKey) -and -not $typeMap.Contains($typeKey)) { $typeMap[$typeKey] = $typeText }
    }
    $folderOptions = @($folderMap.Keys | ForEach-Object { "<option value='$(ConvertTo-HtmlEncodedText $_)'>$(ConvertTo-HtmlEncodedText $folderMap[$_])</option>" })
    $typeOptions = @($typeMap.Keys | ForEach-Object { "<option value='$(ConvertTo-HtmlEncodedText $_)'>$(ConvertTo-HtmlEncodedText $typeMap[$_])</option>" })

    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    $writer = [System.IO.StreamWriter]::new($ReportPath, $false, $utf8NoBom)
    try {
        Write-CalendarReportHeader -Writer $writer -PstItem $PstItem -SortedRecords $sorted -FolderOptions $folderOptions -TypeOptions $typeOptions -LogPath $LogPath
        $writtenCount = 0
        $totalCount = [Math]::Max(1, $sorted.Count)
        foreach ($record in $sorted) {
            Write-CalendarRecordHtml -Writer $writer -Record $record -RecordIndex $writtenCount
            $writtenCount++
            if (($writtenCount -eq $totalCount) -or ($writtenCount % 100 -eq 0)) {
                Write-ReportLog "Calendar HTML report progress: $writtenCount of $totalCount records."
                Write-ConversionStage -Stage 'WritingReport' -Extra ("Written={0}|Total={1}" -f $writtenCount, $totalCount)
            }
        }
        Write-CalendarReportFooter -Writer $writer -ScriptText (Get-CalendarReportScript)
    }
    finally {
        $writer.Dispose()
    }
}

function Write-ContactsHtmlReport {
    param(
        [Parameter(Mandatory = $true)][object]$Records,
        [Parameter(Mandatory = $true)]$PstItem,
        [Parameter(Mandatory = $true)][string]$ReportPath,
        [string]$LogPath = $script:LogPath
    )

    Write-ReportLog 'Preparing Contacts HTML report data.'
    Write-ConversionStage -Stage 'PreparingReport'
    $sorted = @($Records | Sort-Object DisplayName, FullName, CompanyName, FolderPath)
    $folderMap = [ordered]@{}
    $categoryMap = [ordered]@{}
    foreach ($record in $sorted) {
        $folderText = [string](Get-PropSafe -Object $record -Name 'FolderPath' -Default '')
        $folderKey = ConvertTo-NormalizedFilterText $folderText
        if (-not [string]::IsNullOrWhiteSpace($folderKey) -and -not $folderMap.Contains($folderKey)) { $folderMap[$folderKey] = $folderText }

        foreach ($category in @(Split-RecordCategories $record.Categories)) {
            $categoryKey = ConvertTo-NormalizedFilterText $category
            if (-not [string]::IsNullOrWhiteSpace($categoryKey) -and -not $categoryMap.Contains($categoryKey)) { $categoryMap[$categoryKey] = $category }
        }
    }
    $folderOptions = @($folderMap.Keys | ForEach-Object { "<option value='$(ConvertTo-HtmlEncodedText $_)'>$(ConvertTo-HtmlEncodedText $folderMap[$_])</option>" })
    $categoryOptions = @($categoryMap.Keys | ForEach-Object { "<option value='$(ConvertTo-HtmlEncodedText $_)'>$(ConvertTo-HtmlEncodedText $categoryMap[$_])</option>" })

    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    $writer = [System.IO.StreamWriter]::new($ReportPath, $false, $utf8NoBom)
    try {
        Write-ContactsReportHeader -Writer $writer -PstItem $PstItem -SortedRecords $sorted -FolderOptions $folderOptions -CategoryOptions $categoryOptions -LogPath $LogPath
        $writtenCount = 0
        $totalCount = [Math]::Max(1, $sorted.Count)
        foreach ($record in $sorted) {
            Write-ContactRecordHtml -Writer $writer -Record $record
            $writtenCount++
            if (($writtenCount -eq $totalCount) -or ($writtenCount % 100 -eq 0)) {
                Write-ReportLog "Contacts HTML report progress: $writtenCount of $totalCount records."
                Write-ConversionStage -Stage 'WritingReport' -Extra ("Written={0}|Total={1}" -f $writtenCount, $totalCount)
            }
        }
        Write-StaticRecordReportFooter -Writer $writer -ScriptText (Get-ContactsReportScript)
    }
    finally {
        $writer.Dispose()
    }
}

function Resolve-EmailViewerExecutable {
    $candidates = @(
        $env:PURVIEW_EMAIL_VIEWER_PATH,
        (Join-Path $PSScriptRoot '..\EmailReviewViewer\EmailReviewViewer.App.exe'),
        (Join-Path $PSScriptRoot '..\EmailReviewViewer\EmailReviewViewer.App\bin\Release\net8.0-windows\EmailReviewViewer.App.exe'),
        (Join-Path $PSScriptRoot '..\EmailReviewViewer\artifacts\publish\win-x64\EmailReviewViewer.App.exe')
    )
    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and
            (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return [IO.Path]::GetFullPath($candidate)
        }
    }
    throw 'Email report requires EmailReviewViewer.App.exe. Keep the EmailReviewViewer folder beside the converter, or build/publish EmailReviewViewer from source.'
}

function ConvertTo-EmailImportDate {
    param([AllowNull()][object]$Value)
    if (Test-MissingDate $Value) { return $null }
    return ([datetime]$Value).ToUniversalTime().ToString('o')
}

function Write-EmailDatabase {
    param(
        [Parameter(Mandatory = $true)][object]$Records,
        [Parameter(Mandatory = $true)][string]$DatabasePath
    )

    $viewerPath = Resolve-EmailViewerExecutable
    $stagingPath = $DatabasePath + '.staging.ndjson'
    $writer = [IO.StreamWriter]::new($stagingPath, $false, [Text.UTF8Encoding]::new($false))
    try {
        foreach ($record in @($Records)) {
            $preview = @(([string]$record.BodyText -split '\r?\n') |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -First 1)
            $previewText = if ($preview.Count) { [string]$preview[0] } else { '' }
            if ($previewText.Length -gt 500) { $previewText = $previewText.Substring(0, 500) }
            $payload = [ordered]@{
                FolderPath       = [string]$record.FolderPath
                SenderName       = [string]$record.SenderName
                SenderAddress    = [string]$record.SenderEmail
                ToRecipients     = [string]$record.To
                CcRecipients     = [string]$record.Cc
                Subject          = [string]$record.Subject
                SentUtc          = ConvertTo-EmailImportDate $record.SentOn
                ReceivedUtc      = ConvertTo-EmailImportDate $record.ReceivedTime
                Preview          = $previewText
                BodyText         = [string]$record.BodyText
                MessageClass     = [string]$record.MessageClass
                EntryId          = [string]$record.EntryId
                ConversationId   = [string]$record.ConversationId
                ConversationTopic = [string]$record.ConversationTopic
            }
            $writer.WriteLine(($payload | ConvertTo-Json -Compress -Depth 3))
        }
    }
    finally {
        $writer.Dispose()
    }

    Write-ReportLog "Importing Email SQLite database from staging file: $stagingPath"
    Write-ConversionStage -Stage 'ImportingEmailDatabase'
    $psi = [Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = $viewerPath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    foreach ($argument in @('--import', $stagingPath, '--database', $DatabasePath, '--expected-count', [string]@($Records).Count)) {
        [void]$psi.ArgumentList.Add($argument)
    }
    $process = [Diagnostics.Process]::Start($psi)
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    $process.WaitForExit()
    $stdout = $stdoutTask.Result
    $stderr = $stderrTask.Result
    if ($process.ExitCode -ne 0) {
        throw "Email database import failed with exit code $($process.ExitCode). Staging file retained at $stagingPath. $stderr"
    }
    Remove-Item -LiteralPath $stagingPath -Force
    Write-ReportLog "Email database importer completed: $stdout"
}

function Invoke-ReportConversion {
    if (-not $TeamsReport -and -not $EmailReport -and -not $CalendarReport -and -not $ContactsReport) {
        $msg = 'At least one report type must be selected (TeamsReport, EmailReport, CalendarReport, and/or ContactsReport).'
        Write-Output ("CONVERSION_ERROR|{0}ExitCode=2|Message={1}" -f (Get-RunIdField), ($msg -replace '[\r\n|]', ' '))
        throw $msg
    }

    Assert-StaForOutlookCom

    $downloads = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    if ($UseSampleData) {
        $pstItem = [pscustomobject]@{
            FullName = 'Built-in sample data'
            Name = 'Sample-PurviewTeamsExport.pst'
            BaseName = 'Sample-PurviewTeamsExport'
            Length = 0
        }
    }
    else {
        $script:PstPath = Request-PstPath -InitialPath $PstPath
        $pstItem = Get-Item -LiteralPath $script:PstPath
    }

    $baseName = [IO.Path]::GetFileNameWithoutExtension($pstItem.Name) -replace '[^a-zA-Z0-9._-]', '_'
    $displayHtmlPath = Resolve-OutputFilePath -Path $OutputPath -DefaultDirectory $downloads -DefaultFileName "PurviewTeamsPst_ConversationReport_$baseName`_$stamp.html"
    $displayLogPath = Resolve-OutputFilePath -Path $LogPath -DefaultDirectory $downloads -DefaultFileName "PurviewTeamsPst_ConversationReport_$baseName`_$stamp.log"
    $htmlPaths = Get-ReportOutputPaths -DisplayPath $displayHtmlPath -TeamsReport $TeamsReport -EmailReport $EmailReport -CalendarReport $CalendarReport -ContactsReport $ContactsReport
    $logPaths = Get-ReportOutputPaths -DisplayPath $displayLogPath -TeamsReport $TeamsReport -EmailReport $EmailReport -CalendarReport $CalendarReport -ContactsReport $ContactsReport

    $script:TeamsOutputPath = $htmlPaths.TeamsPath
    $script:EmailOutputPath = $htmlPaths.EmailPath
    $script:CalendarOutputPath = $htmlPaths.CalendarPath
    $script:ContactsOutputPath = $htmlPaths.ContactsPath
    $script:TeamsLogPath = $logPaths.TeamsPath
    $script:EmailLogPath = $logPaths.EmailPath
    $script:CalendarLogPath = $logPaths.CalendarPath
    $script:ContactsLogPath = $logPaths.ContactsPath
    $script:OutputPath = @(
        $script:TeamsOutputPath,
        $script:EmailOutputPath,
        $script:CalendarOutputPath,
        $script:ContactsOutputPath
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
    $script:LogPath = @(
        $script:TeamsLogPath,
        $script:EmailLogPath,
        $script:CalendarLogPath,
        $script:ContactsLogPath
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1

    if ($script:TeamsOutputPath) { Assert-OutputPathsSafe -ReportPath $script:TeamsOutputPath -LogPath $script:TeamsLogPath }
    if ($script:EmailOutputPath) { Assert-OutputPathsSafe -ReportPath $script:EmailOutputPath -LogPath $script:EmailLogPath }
    if ($script:CalendarOutputPath) { Assert-OutputPathsSafe -ReportPath $script:CalendarOutputPath -LogPath $script:CalendarLogPath }
    if ($script:ContactsOutputPath) { Assert-OutputPathsSafe -ReportPath $script:ContactsOutputPath -LogPath $script:ContactsLogPath }

    foreach ($pathToPrepare in @($script:TeamsOutputPath, $script:EmailOutputPath, $script:CalendarOutputPath, $script:ContactsOutputPath, $script:TeamsLogPath, $script:EmailLogPath, $script:CalendarLogPath, $script:ContactsLogPath) | Where-Object { $_ } | Sort-Object -Unique) {
        $outDir = Split-Path -LiteralPath $pathToPrepare
        if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = $downloads }
        # New-Item has no -LiteralPath, and -Path treats [ ] as wildcards; use the .NET API so a
        # bracketed output directory is created literally. CreateDirectory is a no-op if it exists.
        if (-not (Test-Path -LiteralPath $outDir)) { [System.IO.Directory]::CreateDirectory($outDir) | Out-Null }
    }

    # Start each conversion with a fresh log so the GUI never reads stale final-stage lines
    # from a previous run and shows cleanup/complete progress too early. Create/truncate mode
    # clears the file; UTF8Encoding($false) writes no BOM so the GUI log-tail parser reads the
    # first token cleanly; AutoFlush keeps lines visible to the live tail immediately.
    $script:LogWriters = @()
    foreach ($logPath in @($script:TeamsLogPath, $script:EmailLogPath, $script:CalendarLogPath, $script:ContactsLogPath) | Where-Object { $_ } | Sort-Object -Unique) {
        $writer = [System.IO.StreamWriter]::new($logPath, $false, [System.Text.UTF8Encoding]::new($false))
        $writer.AutoFlush = $true
        $script:LogWriters += $writer
    }
    $script:LogWriter = if ($script:LogWriters.Count -gt 0) { $script:LogWriters[0] } else { $null }

    if ($KeepPstAttached) {
        Write-Warning 'KeepPstAttached was specified. The PST will remain attached to the current Outlook profile unless you remove it manually.'
    }

    $outlook = $null
    $namespace = $null
    $root = $null
    $teamsRecords = New-Object System.Collections.Generic.List[object]
    $emailRecords = New-Object System.Collections.Generic.List[object]
    $calendarRecords = New-Object System.Collections.Generic.List[object]
    $contactsRecords = New-Object System.Collections.Generic.List[object]
    $weAttached = $false

    try {
        Write-ReportLog "Starting PST to HTML conversion. PST: $($pstItem.FullName)"
        Write-ReportLog "PST size bytes: $($pstItem.Length)"

        if ($UseSampleData) {
            $sample = Get-SampleRecord
            $sampleTeams = @($sample.Teams)
            $sampleEmail = @($sample.Email)
            $sampleCalendar = @($sample.Calendar)
            $sampleContacts = @($sample.Contacts)
            $script:Stats.FoldersScanned = 2
            $script:Stats.ItemsAttempted = $sampleTeams.Count + $sampleEmail.Count + $sampleCalendar.Count + $sampleContacts.Count
            foreach ($record in $sampleTeams) {
                if ($TeamsReport) {
                    [void]$teamsRecords.Add($record)
                    $script:Stats.TeamsItemsExported++
                    $script:Stats.ItemsExported++
                }
                else {
                    $script:Stats.ItemsSkipped++
                }
            }
            foreach ($record in $sampleEmail) {
                if ($EmailReport) {
                    [void]$emailRecords.Add($record)
                    $script:Stats.EmailItemsExported++
                    $script:Stats.ItemsExported++
                }
                else {
                    $script:Stats.ItemsSkipped++
                }
            }
            foreach ($record in $sampleCalendar) {
                if ($CalendarReport) {
                    [void]$calendarRecords.Add($record)
                    $script:Stats.CalendarItemsExported++
                    $script:Stats.ItemsExported++
                }
                else {
                    $script:Stats.ItemsSkipped++
                }
            }
            foreach ($record in $sampleContacts) {
                if ($ContactsReport) {
                    [void]$contactsRecords.Add($record)
                    $script:Stats.ContactsItemsExported++
                    $script:Stats.ItemsExported++
                }
                else {
                    $script:Stats.ItemsSkipped++
                }
            }
        }
        else {
            $outlook = Invoke-OutlookComOperation -Operation 'starting Outlook COM automation' -ScriptBlock { New-Object -ComObject Outlook.Application }
            $namespace = Invoke-OutlookComOperation -Operation 'opening the Outlook MAPI namespace' -ScriptBlock { $outlook.GetNamespace('MAPI') }

            $root = Invoke-OutlookComOperation -Operation 'checking whether the PST is already attached' -ScriptBlock { Find-StoreRootForPst -Namespace $namespace -TargetPath $pstItem.FullName }
            if ($null -ne $root) {
                Write-ReportLog 'PST is already attached to Outlook; reusing the existing store (it will remain attached when finished).'
            }
            else {
                Write-ReportLog 'Attaching PST to Outlook profile temporarily.'
                Invoke-OutlookComOperation -Operation 'attaching the PST to Outlook' -ScriptBlock { $namespace.AddStoreEx($pstItem.FullName, 3) } | Out-Null # 3 = Unicode PST
                $weAttached = $true

                $root = Invoke-OutlookComOperation -Operation 'locating the attached PST in Outlook' -MaxAttempts 10 -ScriptBlock { Wait-StoreRootForPst -Namespace $namespace -TargetPath $pstItem.FullName -TimeoutSeconds 30 }
                if ($null -eq $root) {
                    throw 'Could not locate the attached PST in Outlook stores after waiting up to 30 seconds.'
                }
            }

            $rootName = Get-PropSafe -Object $root -Name 'Name' -Default $pstItem.BaseName
            Read-OutlookFolder -Folder $root -FolderPath $rootName -TeamsRecords $teamsRecords -EmailRecords $emailRecords -CalendarRecords $calendarRecords -ContactsRecords $contactsRecords
        }

        Write-ReportLog "Finished reading PST. Teams items: $($teamsRecords.Count); Email items: $($emailRecords.Count); Calendar items: $($calendarRecords.Count); Contacts items: $($contactsRecords.Count)"
        Write-ConversionStage -Stage 'FinishedReading'
        if ($TeamsReport) {
            Write-ReportLog 'Writing Teams HTML report.'
            Write-HtmlReport -Records $teamsRecords.ToArray() -PstItem $pstItem -ReportPath $script:TeamsOutputPath
            Write-ReportLog "HTML report written to $script:TeamsOutputPath"
        }
        if ($EmailReport) {
            Write-ReportLog 'Writing Email SQLite database.'
            Write-EmailDatabase -Records $emailRecords.ToArray() -DatabasePath $script:EmailOutputPath
            Write-ReportLog "Email SQLite database written to $script:EmailOutputPath"
        }
        if ($CalendarReport) {
            Write-ReportLog 'Writing Calendar HTML report.'
            Write-CalendarHtmlReport -Records $calendarRecords.ToArray() -PstItem $pstItem -ReportPath $script:CalendarOutputPath -LogPath $script:CalendarLogPath
            Write-ReportLog "Calendar HTML report written to $script:CalendarOutputPath"
        }
        if ($ContactsReport) {
            Write-ReportLog 'Writing Contacts HTML report.'
            Write-ContactsHtmlReport -Records $contactsRecords.ToArray() -PstItem $pstItem -ReportPath $script:ContactsOutputPath -LogPath $script:ContactsLogPath
            Write-ReportLog "Contacts HTML report written to $script:ContactsOutputPath"
        }
        Write-ConversionStage -Stage 'ReportWritten'
        Write-Output ("CONVERSION_RESULT|{0}OutputPath={1}|LogPath={2}|ItemsExported={3}|ItemReadFailures={4}|AttachmentReadFailures={5}|SubfolderScanFailures={6}|TeamsOutputPath={7}|EmailOutputPath={8}|CalendarOutputPath={9}|ContactsOutputPath={10}|TeamsLogPath={11}|EmailLogPath={12}|CalendarLogPath={13}|ContactsLogPath={14}|TeamsItemsExported={15}|EmailItemsExported={16}|CalendarItemsExported={17}|ContactsItemsExported={18}" -f (Get-RunIdField), $script:OutputPath, $script:LogPath, $script:Stats.ItemsExported, $script:Stats.ItemReadFailures, $script:Stats.AttachmentReadFailures, $script:Stats.SubfolderScanFailures, $script:TeamsOutputPath, $script:EmailOutputPath, $script:CalendarOutputPath, $script:ContactsOutputPath, $script:TeamsLogPath, $script:EmailLogPath, $script:CalendarLogPath, $script:ContactsLogPath, $script:Stats.TeamsItemsExported, $script:Stats.EmailItemsExported, $script:Stats.CalendarItemsExported, $script:Stats.ContactsItemsExported)
    }
    catch {
        $fatalMessage = if ($_.Exception.Message) { $_.Exception.Message } else { [string]$_.Exception }
        Write-ReportLog "Fatal error: $fatalMessage" 'ERROR'
        try { Write-ConversionError -Message $fatalMessage } catch { }
        throw
    }
    finally {
        # Detach/COM cleanup must never abort a successful conversion. Outlook can die mid-scan
        # (0x800706BA); reports may already be written and CONVERSION_RESULT already emitted.
        try {
            if ($weAttached -and -not $KeepPstAttached -and $null -ne $namespace) {
                Write-ReportLog 'Detaching PST from Outlook profile.'
                Write-ConversionStage -Stage 'Detaching'
                try {
                    [void](Confirm-OutlookPstCleanup -Namespace $namespace -TargetPath $pstItem.FullName -RootFolder $root)
                }
                catch {
                    Write-ReportLog "Could not detach PST automatically. $($_.Exception.Message)" 'WARN'
                    if (Test-PstStoreAttached -Namespace $namespace -TargetPath $pstItem.FullName) {
                        Write-ReportLog "PST is still attached to your Outlook profile: $($pstItem.FullName). Remove it manually in Outlook (File -> Account Settings -> Data Files), then verify your profile is unchanged." 'WARN'
                    }
                }
            }
        }
        catch {
            Write-ReportLog "PST detach cleanup failed unexpectedly. $($_.Exception.Message)" 'WARN'
        }

        try {
            $elapsed = (Get-Date) - $script:Stats.StartedAt
            Write-ReportLog ('Summary: folders={0}; attempted={1}; exported={2}; skipped={3}; itemReadFailures={4}; attachmentReadFailures={5}; subfolderScanFailures={6}; elapsed={7}' -f $script:Stats.FoldersScanned, $script:Stats.ItemsAttempted, $script:Stats.ItemsExported, $script:Stats.ItemsSkipped, $script:Stats.ItemReadFailures, $script:Stats.AttachmentReadFailures, $script:Stats.SubfolderScanFailures, $elapsed)
        }
        catch { }

        try {
            Close-ComObjectSafe $root
            Close-ComObjectSafe $namespace
            Close-ComObjectSafe $outlook
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
        catch { }

        foreach ($writer in @($script:LogWriters)) {
            if ($null -ne $writer) { $writer.Dispose() }
        }
        $script:LogWriters = @()
        $script:LogWriter = $null
    }
}

Invoke-ReportConversion
