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

$sourceText = Get-Content -LiteralPath $launcherPath -Raw
$pattern = [regex]'(?s)\$killer = \[System\.Diagnostics\.Process\]::Start\(\$psi\).*?\[void\]\$killer\.WaitForExit\(5000\).*?if \(\$killer\.HasExited -and \$killer\.ExitCode -eq 0\) \{\s*\$treeKilled = \$true\s*\}'
$sourceContainsExitCodeGate = $pattern.IsMatch($sourceText)
if (-not $sourceContainsExitCodeGate) {
    throw 'Launcher source no longer contains the taskkill exit-code gate in Stop-ProcessTree.'
}

Remove-Item -LiteralPath $reportPath,$logPath -Force -ErrorAction SilentlyContinue
$launcherOutput = & pwsh -NoProfile -File $launcherPath -NoGui -UseSampleData -OutputPath $reportPath -LogPath $logPath 2>&1
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
    throw "Launcher smoke test failed with exit code $exitCode.`n$($launcherOutput -join [Environment]::NewLine)"
}
if (-not (Test-Path -LiteralPath $reportPath)) {
    throw "Expected report was not created: $reportPath"
}
if (-not (Test-Path -LiteralPath $logPath)) {
    throw "Expected log was not created: $logPath"
}

$stdoutText = $launcherOutput -join "`n"
$logText = Get-Content -LiteralPath $logPath -Raw
if ($stdoutText -notmatch 'CONVERSION_RESULT\|') {
    throw 'Launcher stdout did not contain a CONVERSION_RESULT line.'
}
if ($logText -notmatch 'exported=3') {
    throw 'Verification log did not confirm exported=3.'
}
if ($logText -notmatch 'itemReadFailures=0') {
    throw 'Verification log did not confirm itemReadFailures=0.'
}

[pscustomobject]@{
    SourceContainsExitCodeGate = $sourceContainsExitCodeGate
    SourceLauncherExitCode = $exitCode
    ReportPath = $reportPath
    LogPath = $logPath
    StdoutHasResultLine = $true
    LogShowsExported3 = $true
    LogShowsZeroItemReadFailures = $true
} | Format-List
