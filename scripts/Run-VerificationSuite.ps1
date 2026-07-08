[CmdletBinding()]
param(
    [string]$RepoRoot = (Join-Path $PSScriptRoot '..')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath($RepoRoot)
$testsPath = Join-Path $repoRoot 'tests'

Import-Module Pester -MinimumVersion 5.0.0 -Force
$config = New-PesterConfiguration
$config.Run.Path = $testsPath
$config.Run.PassThru = $true
$config.Output.Verbosity = 'Detailed'
$config.Should.ErrorAction = 'Stop'

$result = Invoke-Pester -Configuration $config
if ($result.Result -ne 'Passed' -or $result.FailedCount -gt 0) {
    throw "Verification suite failed. Passed=$($result.PassedCount) Failed=$($result.FailedCount) Skipped=$($result.SkippedCount)"
}

Write-Host ("Suite Green: Passed={0} Failed={1} Skipped={2}" -f $result.PassedCount, $result.FailedCount, $result.SkippedCount)
