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
    [switch]$NoPrompt
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Stats = [ordered]@{
    StartedAt              = Get-Date
    FoldersScanned         = 0
    ItemsAttempted         = 0
    ItemsExported          = 0
    ItemsSkipped           = 0
    ItemReadFailures       = 0
    AttachmentReadFailures = 0
}

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

    $parent = Split-Path -Path $clean -Parent
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
        $parent = Split-Path -Path $pathToCheck -Parent
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

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 's'), $Level, $Message
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    Write-Information $line -InformationAction Continue
}

function Write-ConversionProgress {
    param([Parameter(Mandatory = $true)] [string]$FolderPath)

    $elapsed = (Get-Date) - $script:Stats.StartedAt
    $elapsedSeconds = [Math]::Max(1, [int][Math]::Round($elapsed.TotalSeconds))
    $ratePerMinute = [Math]::Round(($script:Stats.ItemsExported / $elapsedSeconds) * 60, 1)
    $safeFolderPath = ([string]$FolderPath).Replace('\', '/').Replace('|', '/')
    Write-Output ("CONVERSION_PROGRESS|ItemsAttempted={0}|ItemsExported={1}|FoldersScanned={2}|ItemReadFailures={3}|ElapsedSeconds={4}|RatePerMinute={5}|FolderPath={6}" -f $script:Stats.ItemsAttempted, $script:Stats.ItemsExported, $script:Stats.FoldersScanned, $script:Stats.ItemReadFailures, $elapsedSeconds, $ratePerMinute, $safeFolderPath)
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

function ConvertTo-SearchText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    $text = ([string]$Value) -replace '\s+', ' '
    return $text.Trim().ToLowerInvariant()
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
        if ($part -notmatch "^[A-Za-z][A-Za-z'.-]*$") { return $false }
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

    $attachments = $null
    try {
        $attachments = $Item.Attachments
        $count = [int]$attachments.Count
        if ($count -le 0) { return '' }
        $rows = New-Object System.Collections.Generic.List[string]
        for ($i = 1; $i -le $count; $i++) {
            $att = $null
            try {
                $att = $attachments.Item($i)
                [void]$rows.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>' -f (ConvertTo-HtmlEncodedText $att.FileName), (ConvertTo-HtmlEncodedText $att.DisplayName), (ConvertTo-HtmlEncodedText $att.Size)))
            }
            catch {
                $script:Stats.AttachmentReadFailures++
            }
            finally { Close-ComObjectSafe $att }
        }
        return "<div class='attachments'><strong>Attachments:</strong><table><thead><tr><th>File name</th><th>Display name</th><th>Size bytes</th></tr></thead><tbody>$($rows -join "`n")</tbody></table></div>"
    }
    catch {
        $script:Stats.AttachmentReadFailures++
        return ''
    }
    finally {
        Close-ComObjectSafe $attachments
    }
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
    $sortTime = $receivedTime
    if ($null -eq $sortTime -or [string]::IsNullOrWhiteSpace([string]$sortTime)) { $sortTime = $sentOn }
    if ($null -eq $sortTime -or [string]::IsNullOrWhiteSpace([string]$sortTime)) { $sortTime = $creationTime }

    $senderDisplay = ConvertTo-NormalizedPersonName $senderName
    if ([string]::IsNullOrWhiteSpace($senderDisplay)) { $senderDisplay = ConvertTo-NormalizedPersonName $senderEmail }
    if ([string]::IsNullOrWhiteSpace($senderDisplay)) { $senderDisplay = '(unknown sender)' }

    $participants = @(Get-ParticipantName -SenderName $senderName -SenderEmail $senderEmail -To $to -Cc $cc)
    $participantKey = Get-ParticipantKey -Participants $participants
    $conversationSubject = if ([string]::IsNullOrWhiteSpace([string]$subject)) { '(no subject)' } else { [string]$subject }

    $conversationSource = 'ParticipantsSubjectFolder'
    $conversationKey = "$participantKey`n$conversationSubject`n$FolderPath"

    [pscustomobject]@{
        SortTime             = $sortTime
        FolderPath           = $FolderPath
        Subject              = $subject
        ConversationSource   = $conversationSource
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
        BodyText             = [string]$body
        AttachmentsHtml      = Get-AttachmentSummaryHtml -Item $Item
    }
}

function Read-OutlookFolder {
    param(
        [Parameter(Mandatory = $true)] [object]$Folder,
        [Parameter(Mandatory = $true)] [string]$FolderPath,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]]$Records
    )

    $script:Stats.FoldersScanned++
    Write-ReportLog "Scanning folder: $FolderPath"
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

                # Purview Teams messages are usually mail-like items in an Exchange PST. Export any
                # item that has normal message properties instead of requiring one exact MessageClass.
                if ($messageClass -or $body -or $subject) {
                    [void]$Records.Add((Get-MessageRecord -Item $item -FolderPath $FolderPath))
                    $script:Stats.ItemsExported++
                }
                else {
                    $script:Stats.ItemsSkipped++
                }

                if ($Records.Count -gt 0 -and $Records.Count % $LogEvery -eq 0) {
                    Write-ConversionProgress -FolderPath $FolderPath
                    Write-ReportLog "Collected $($Records.Count) message-like items so far."
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
        for ($j = 1; $j -le [int]$subFolders.Count; $j++) {
            $child = $null
            try {
                $child = $subFolders.Item($j)
                $childName = Get-PropSafe -Object $child -Name 'Name' -Default "Folder$j"
                Read-OutlookFolder -Folder $child -FolderPath "$FolderPath\$childName" -Records $Records
            }
            catch {
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
        for ($i = 1; $i -le [int]$stores.Count; $i++) {
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
        $otherPeople = @($groupParticipants | Where-Object {
            $normalized = ConvertTo-NormalizedPersonName $_
            -not [string]::IsNullOrWhiteSpace($normalized) -and $normalized -ne $senderParticipantName
        })
        if ($otherPeople.Count -eq 0 -and $groupParticipants.Count -gt 1) {
            $otherPeople = @($groupParticipants | Where-Object { $_ -ne $record.SenderDisplay })
        }
        $displayOtherPeople = @(Get-ParticipantDisplayOrder -Participants $otherPeople -PreferredParticipants $PreferredParticipants)
        $otherPeoplePreview = Format-ParticipantPreview -Participants $displayOtherPeople -MaximumNames 4 -PreferredParticipants $PreferredParticipants
        $otherPeopleLabel = if ($displayOtherPeople.Count -gt 0) { 'with ' + $otherPeoplePreview } else { 'no other named participants shown' }
        $speakerContext = "Chat: $conversationTitle • $otherPeopleLabel"
        $bodyHtml = ConvertTo-HtmlBody $record.BodyText
        $details = @"
<details class='message-details'>
  <summary>Details</summary>
  <div><strong>Folder:</strong> $(ConvertTo-HtmlEncodedText $record.FolderPath)</div>
  <div><strong>Subject:</strong> $(ConvertTo-HtmlEncodedText $record.Subject)</div>
  <div><strong>Message class:</strong> $(ConvertTo-HtmlEncodedText $record.MessageClass)</div>
  <div><strong>From:</strong> $(ConvertTo-HtmlEncodedText $record.SenderName) &lt;$(ConvertTo-HtmlEncodedText $record.SenderEmail)&gt;</div>
  <div><strong>To:</strong> $(ConvertTo-HtmlEncodedText $record.To)</div>
  <div><strong>Cc:</strong> $(ConvertTo-HtmlEncodedText $record.Cc)</div>
  <div><strong>Sent:</strong> $(ConvertTo-HtmlEncodedText $record.SentOn)</div>
  <div><strong>Received:</strong> $(ConvertTo-HtmlEncodedText $record.ReceivedTime)</div>
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

function Get-SampleRecord {
    $script:Stats.FoldersScanned = 1
    $script:Stats.ItemsAttempted = 3
    $script:Stats.ItemsExported = 3
    return @(
        [pscustomobject]@{ SortTime = [datetime]'2024-01-01T09:00:00'; FolderPath = 'SamplePst\TeamsMessagesData'; Subject = 'Sample Teams chat'; ConversationSource = 'ParticipantsSubjectFolder'; MessageClass = 'IPM.Note'; SenderName = 'Torey Page'; SenderEmail = 'torey@example.com'; SenderDisplay = 'Torey Page'; To = 'Linda Artley'; Cc = ''; Participants = @('Linda Artley','Torey Page'); ParticipantsKey = 'Linda Artley || Torey Page'; ConversationKey = "Linda Artley || Torey Page`nSample Teams chat`nSamplePst\TeamsMessagesData"; ConversationTitle = 'Sample Teams chat'; SentOn = [datetime]'2024-01-01T09:00:00'; ReceivedTime = [datetime]'2024-01-01T09:00:05'; CreationTime = [datetime]'2024-01-01T09:00:05'; EntryId = 'sample-entry-1'; BodyText = 'Hello Linda. This sample validates HTML encoding: <script>alert(1)</script>'; AttachmentsHtml = '' }
        [pscustomobject]@{ SortTime = [datetime]'2024-01-01T09:01:00'; FolderPath = 'SamplePst\TeamsMessagesData'; Subject = 'Sample Teams chat'; ConversationSource = 'ParticipantsSubjectFolder'; MessageClass = 'IPM.Note'; SenderName = 'Linda Artley'; SenderEmail = 'linda@example.com'; SenderDisplay = 'Linda Artley'; To = 'Torey Page'; Cc = ''; Participants = @('Linda Artley','Torey Page'); ParticipantsKey = 'Linda Artley || Torey Page'; ConversationKey = "Linda Artley || Torey Page`nSample Teams chat`nSamplePst\TeamsMessagesData"; ConversationTitle = 'Sample Teams chat'; SentOn = [datetime]'2024-01-01T09:01:00'; ReceivedTime = [datetime]'2024-01-01T09:01:05'; CreationTime = [datetime]'2024-01-01T09:01:05'; EntryId = 'sample-entry-2'; BodyText = "Thanks.`nThis message has two lines."; AttachmentsHtml = "<div class='attachments'><strong>Attachments:</strong><table><tbody><tr><td>sample.pdf</td><td>sample.pdf</td><td>1234</td></tr></tbody></table></div>" }
        [pscustomobject]@{ SortTime = [datetime]'2024-01-01T10:00:00'; FolderPath = 'SamplePst\Other'; Subject = 'Fireflies note'; ConversationSource = 'ParticipantsSubjectFolder'; MessageClass = 'IPM.Note'; SenderName = 'Fireflies.ai Notetaker'; SenderEmail = 'bot@example.com'; SenderDisplay = 'Fireflies.ai Notetaker'; To = 'Torey Page'; Cc = ''; Participants = @('Fireflies.ai Notetaker','Torey Page'); ParticipantsKey = 'Fireflies.ai Notetaker || Torey Page'; ConversationKey = "Fireflies.ai Notetaker || Torey Page`nFireflies note`nSamplePst\Other"; ConversationTitle = 'Fireflies note'; SentOn = [datetime]'2024-01-01T10:00:00'; ReceivedTime = [datetime]'2024-01-01T10:00:05'; CreationTime = [datetime]'2024-01-01T10:00:05'; EntryId = 'sample-entry-3'; BodyText = 'Fireflies should appear in Other detected names / IDs.'; AttachmentsHtml = '' }
    )
}

function Write-HtmlReport {
    param(
        [Parameter(Mandatory = $true)][object]$Records,
        [Parameter(Mandatory = $true)]$PstItem,
        [Parameter(Mandatory = $true)][string]$ReportPath
    )

    Write-ReportLog 'Preparing HTML report data.'
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
            }
        }
        Write-ReportLog 'Finalizing HTML report.'
        Write-ReportFooter -Writer $writer
    }
    finally {
        $writer.Dispose()
    }
}

function Invoke-ReportConversion {
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
    $script:OutputPath = Resolve-OutputFilePath -Path $OutputPath -DefaultDirectory $downloads -DefaultFileName "PurviewTeamsPst_ConversationReport_$baseName`_$stamp.html"
    $script:LogPath = Resolve-OutputFilePath -Path $LogPath -DefaultDirectory $downloads -DefaultFileName "PurviewTeamsPst_ConversationReport_$baseName`_$stamp.log"

    if ([IO.Path]::GetExtension($script:OutputPath) -notin @('.html','.htm')) {
        Write-Warning "OutputPath does not end in .html or .htm: $script:OutputPath"
    }
    if ([IO.Path]::GetExtension($script:LogPath) -notin @('.log','.txt')) {
        Write-Warning "LogPath does not end in .log or .txt: $script:LogPath"
    }

    Assert-OutputPathsSafe -ReportPath $script:OutputPath -LogPath $script:LogPath

    foreach ($pathToPrepare in @($script:OutputPath, $script:LogPath)) {
        $outDir = Split-Path -Path $pathToPrepare -Parent
        if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = $downloads }
        if (-not (Test-Path -LiteralPath $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }
    }

    # Start each conversion with a fresh log so the GUI never reads stale final-stage lines
    # from a previous run and shows cleanup/complete progress too early.
    Set-Content -LiteralPath $script:LogPath -Value '' -Encoding UTF8

    if ($KeepPstAttached) {
        Write-Warning 'KeepPstAttached was specified. The PST will remain attached to the current Outlook profile unless you remove it manually.'
    }

    $outlook = $null
    $namespace = $null
    $root = $null
    $records = New-Object System.Collections.Generic.List[object]
    $attached = $false

    try {
        Write-ReportLog "Starting PST to HTML conversion. PST: $($pstItem.FullName)"
        Write-ReportLog "PST size bytes: $($pstItem.Length)"

        if ($UseSampleData) {
            foreach ($record in (Get-SampleRecord)) { [void]$records.Add($record) }
        }
        else {
            $outlook = Invoke-OutlookComOperation -Operation 'starting Outlook COM automation' -ScriptBlock { New-Object -ComObject Outlook.Application }
            $namespace = Invoke-OutlookComOperation -Operation 'opening the Outlook MAPI namespace' -ScriptBlock { $outlook.GetNamespace('MAPI') }

            Write-ReportLog 'Attaching PST to Outlook profile temporarily.'
            Invoke-OutlookComOperation -Operation 'attaching the PST to Outlook' -ScriptBlock { $namespace.AddStoreEx($pstItem.FullName, 3) } | Out-Null # 3 = Unicode PST
            $attached = $true

            $root = Invoke-OutlookComOperation -Operation 'locating the attached PST in Outlook' -MaxAttempts 10 -ScriptBlock { Wait-StoreRootForPst -Namespace $namespace -TargetPath $pstItem.FullName -TimeoutSeconds 30 }
            if ($null -eq $root) {
                throw 'Could not locate the attached PST in Outlook stores after waiting up to 30 seconds.'
            }

            $rootName = Get-PropSafe -Object $root -Name 'Name' -Default $pstItem.BaseName
            Read-OutlookFolder -Folder $root -FolderPath $rootName -Records $records
        }

        Write-ReportLog "Finished reading PST. Message-like items collected: $($records.Count)"
        Write-HtmlReport -Records $records.ToArray() -PstItem $pstItem -ReportPath $script:OutputPath
        Write-ReportLog "HTML report written to $script:OutputPath"
        Write-Output ("CONVERSION_RESULT|OutputPath={0}|LogPath={1}|ItemsExported={2}|ItemReadFailures={3}|AttachmentReadFailures={4}" -f $script:OutputPath, $script:LogPath, $script:Stats.ItemsExported, $script:Stats.ItemReadFailures, $script:Stats.AttachmentReadFailures)
    }
    catch {
        Write-ReportLog "Fatal error: $($_.Exception.Message)" 'ERROR'
        throw
    }
    finally {
        if ($attached -and -not $KeepPstAttached) {
            try {
                if ($null -ne $root) {
                    Write-ReportLog 'Detaching PST from Outlook profile.'
                    Invoke-OutlookComOperation -Operation 'detaching the PST from Outlook' -ScriptBlock { $namespace.RemoveStore($root) }
                }
            }
            catch {
                Write-ReportLog "Could not detach PST automatically. You may need to remove it from Outlook manually. $($_.Exception.Message)" 'WARN'
            }
        }

        $elapsed = (Get-Date) - $script:Stats.StartedAt
        Write-ReportLog ('Summary: folders={0}; attempted={1}; exported={2}; skipped={3}; itemReadFailures={4}; attachmentReadFailures={5}; elapsed={6}' -f $script:Stats.FoldersScanned, $script:Stats.ItemsAttempted, $script:Stats.ItemsExported, $script:Stats.ItemsSkipped, $script:Stats.ItemReadFailures, $script:Stats.AttachmentReadFailures, $elapsed)

        Close-ComObjectSafe $root
        Close-ComObjectSafe $namespace
        Close-ComObjectSafe $outlook
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Invoke-ReportConversion
