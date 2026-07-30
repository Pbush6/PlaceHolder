[CmdletBinding()]
param(
    [string]$LauncherPath = (Join-Path $PSScriptRoot '..\src\Start-PurviewTeamsPstToHtmlApp.ps1'),
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) 'PurviewTeamsPstToHtmlApp\Verify-StopProcessTreeRegression'),
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$launcherPath = [IO.Path]::GetFullPath($LauncherPath)
$outputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($outputDirectory)

$reportPath = Join-Path $outputDirectory 'verify-stop-process-tree.html'
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $logPath = Join-Path $outputDirectory 'verify-stop-process-tree.log'
} else {
    $logPath = [IO.Path]::GetFullPath($LogPath)
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($logPath))
}
$teamsReportPath = Join-Path $outputDirectory 'verify-stop-process-tree_Teams.html'
$emailReportPath = Join-Path $outputDirectory 'verify-stop-process-tree_Email.db'
$calendarReportPath = Join-Path $outputDirectory 'verify-stop-process-tree_Calendar.html'
$contactsReportPath = Join-Path $outputDirectory 'verify-stop-process-tree_Contacts.html'
$logDirectory = [IO.Path]::GetDirectoryName($logPath)
$logBaseName = [IO.Path]::GetFileNameWithoutExtension($logPath)
$logExtension = [IO.Path]::GetExtension($logPath)
$teamsLogPath = Join-Path $logDirectory "$($logBaseName)_Teams$logExtension"
$emailLogPath = Join-Path $logDirectory "$($logBaseName)_Email$logExtension"
$calendarLogPath = Join-Path $logDirectory "$($logBaseName)_Calendar$logExtension"
$contactsLogPath = Join-Path $logDirectory "$($logBaseName)_Contacts$logExtension"

$sourceText = Get-Content -LiteralPath $launcherPath -Raw
$pattern = [regex]'(?s)\$killer = \[System\.Diagnostics\.Process\]::Start\(\$psi\).*?\[void\]\$killer\.WaitForExit\(5000\).*?if \(\$killer\.HasExited -and \$killer\.ExitCode -eq 0\) \{\s*\$treeKilled = \$true\s*\}'
$sourceContainsExitCodeGate = $pattern.IsMatch($sourceText)
if (-not $sourceContainsExitCodeGate) {
    throw 'Launcher source no longer contains the taskkill exit-code gate in Stop-ProcessTree.'
}

Remove-Item -LiteralPath $reportPath,$logPath,$teamsReportPath,$emailReportPath,$calendarReportPath,$contactsReportPath,$teamsLogPath,$emailLogPath,$calendarLogPath,$contactsLogPath -Force -ErrorAction SilentlyContinue
$launcherOutput = & pwsh -NoProfile -File $launcherPath -NoGui -UseSampleData -OutputPath $reportPath -LogPath $logPath 2>&1
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
    throw "Launcher smoke test failed with exit code $exitCode.`n$($launcherOutput -join [Environment]::NewLine)"
}
foreach ($expectedPath in @($teamsReportPath, $emailReportPath, $calendarReportPath, $contactsReportPath, $teamsLogPath, $emailLogPath, $calendarLogPath, $contactsLogPath)) {
    if (-not (Test-Path -LiteralPath $expectedPath -PathType Leaf)) {
        throw "Expected selected typed output was not created: $expectedPath"
    }
}

$stdoutText = $launcherOutput -join "`n"
$teamsLogText = Get-Content -LiteralPath $teamsLogPath -Raw
$emailLogText = Get-Content -LiteralPath $emailLogPath -Raw
if ($stdoutText -notmatch 'CONVERSION_RESULT\|') {
    throw 'Launcher stdout did not contain a CONVERSION_RESULT line.'
}
if ($teamsLogText -notmatch 'exported=6') {
    throw 'Teams verification log did not confirm exported=6.'
}
if ($emailLogText -notmatch 'exported=6') {
    throw 'Email verification log did not confirm exported=6.'
}
if ($teamsLogText -notmatch 'itemReadFailures=0') {
    throw 'Teams verification log did not confirm itemReadFailures=0.'
}
if ($emailLogText -notmatch 'itemReadFailures=0') {
    throw 'Email verification log did not confirm itemReadFailures=0.'
}

[pscustomobject]@{
    SourceContainsExitCodeGate = $sourceContainsExitCodeGate
    SourceLauncherExitCode = $exitCode
    TeamsReportPath = $teamsReportPath
    EmailReportPath = $emailReportPath
    CalendarReportPath = $calendarReportPath
    ContactsReportPath = $contactsReportPath
    TeamsLogPath = $teamsLogPath
    EmailLogPath = $emailLogPath
    CalendarLogPath = $calendarLogPath
    ContactsLogPath = $contactsLogPath
    StdoutHasResultLine = $true
    LogShowsExported6 = $true
    LogShowsZeroItemReadFailures = $true
} | Format-List
