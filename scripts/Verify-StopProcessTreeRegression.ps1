[CmdletBinding()]
param(
    [string]$LauncherPath = (Join-Path $PSScriptRoot '..\src\Start-PurviewTeamsPstToHtmlApp.ps1'),
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\test-output')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$launcherPath = [IO.Path]::GetFullPath($LauncherPath)
$outputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($outputDirectory)

$reportPath = Join-Path $outputDirectory 'verify-stop-process-tree.html'
$logPath = Join-Path $outputDirectory 'verify-stop-process-tree.log'
$teamsReportPath = Join-Path $outputDirectory 'verify-stop-process-tree_Teams.html'
$emailReportPath = Join-Path $outputDirectory 'verify-stop-process-tree_Email.db'
$teamsLogPath = Join-Path $outputDirectory 'verify-stop-process-tree_Teams.log'
$emailLogPath = Join-Path $outputDirectory 'verify-stop-process-tree_Email.log'

$sourceText = Get-Content -LiteralPath $launcherPath -Raw
$pattern = [regex]'(?s)\$killer = \[System\.Diagnostics\.Process\]::Start\(\$psi\).*?\[void\]\$killer\.WaitForExit\(5000\).*?if \(\$killer\.HasExited -and \$killer\.ExitCode -eq 0\) \{\s*\$treeKilled = \$true\s*\}'
$sourceContainsExitCodeGate = $pattern.IsMatch($sourceText)
if (-not $sourceContainsExitCodeGate) {
    throw 'Launcher source no longer contains the taskkill exit-code gate in Stop-ProcessTree.'
}

Remove-Item -LiteralPath $reportPath,$logPath,$teamsReportPath,$emailReportPath,$teamsLogPath,$emailLogPath -Force -ErrorAction SilentlyContinue
$launcherOutput = & pwsh -NoProfile -File $launcherPath -NoGui -UseSampleData -OutputPath $reportPath -LogPath $logPath 2>&1
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
    throw "Launcher smoke test failed with exit code $exitCode.`n$($launcherOutput -join [Environment]::NewLine)"
}
if (-not (Test-Path -LiteralPath $teamsReportPath)) {
    throw "Expected Teams report was not created: $teamsReportPath"
}
if (-not (Test-Path -LiteralPath $emailReportPath)) {
    throw "Expected Email report was not created: $emailReportPath"
}
if (-not (Test-Path -LiteralPath $teamsLogPath)) {
    throw "Expected Teams log was not created: $teamsLogPath"
}
if (-not (Test-Path -LiteralPath $emailLogPath)) {
    throw "Expected Email log was not created: $emailLogPath"
}

$stdoutText = $launcherOutput -join "`n"
$teamsLogText = Get-Content -LiteralPath $teamsLogPath -Raw
$emailLogText = Get-Content -LiteralPath $emailLogPath -Raw
if ($stdoutText -notmatch 'CONVERSION_RESULT\|') {
    throw 'Launcher stdout did not contain a CONVERSION_RESULT line.'
}
if ($teamsLogText -notmatch 'exported=4') {
    throw 'Teams verification log did not confirm exported=4.'
}
if ($emailLogText -notmatch 'exported=4') {
    throw 'Email verification log did not confirm exported=4.'
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
    TeamsLogPath = $teamsLogPath
    EmailLogPath = $emailLogPath
    StdoutHasResultLine = $true
    LogShowsExported4 = $true
    LogShowsZeroItemReadFailures = $true
} | Format-List
