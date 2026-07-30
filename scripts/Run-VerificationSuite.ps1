[CmdletBinding()]
param(
    [string]$RepoRoot = (Join-Path $PSScriptRoot '..'),
    [switch]$Release,
    [string]$ReleaseZipPath,
    [string]$ExpectedVersion,
    [string]$ConverterInputPath,
    [string]$DebugConverterInputPath,
    [string]$ViewerInputPath,
    [string]$ReleaseExtractionRoot
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
if ($Release) {
    $config.Filter.ExcludeTag = @('Build')
}

$result = Invoke-Pester -Configuration $config
if ($result.Result -ne 'Passed' -or $result.FailedCount -gt 0) {
    throw "Verification suite failed. Passed=$($result.PassedCount) Failed=$($result.FailedCount) Skipped=$($result.SkippedCount)"
}

Write-Host ("Suite Green: Passed={0} Failed={1} Skipped={2}" -f $result.PassedCount, $result.FailedCount, $result.SkippedCount)

if ($Release) {
    if ([string]::IsNullOrWhiteSpace($ReleaseZipPath)) {
        throw 'ReleaseZipPath is required when -Release is specified.'
    }
    if ([string]::IsNullOrWhiteSpace($ExpectedVersion)) {
        throw 'ExpectedVersion is required when -Release is specified.'
    }
    foreach ($requiredInput in @(
        @{ Name = 'ConverterInputPath'; Value = $ConverterInputPath },
        @{ Name = 'DebugConverterInputPath'; Value = $DebugConverterInputPath },
        @{ Name = 'ViewerInputPath'; Value = $ViewerInputPath }
    )) {
        if ([string]::IsNullOrWhiteSpace($requiredInput.Value)) {
            throw "$($requiredInput.Name) is required when -Release is specified."
        }
    }

    $releaseVerifier = Join-Path $repoRoot 'scripts\Test-ReleaseArtifact.ps1'
    $verifyArgs = @{
        ZipPath = $ReleaseZipPath
        ExpectedVersion = $ExpectedVersion
        ConverterInputPath = $ConverterInputPath
        DebugConverterInputPath = $DebugConverterInputPath
        ViewerInputPath = $ViewerInputPath
    }
    if (-not [string]::IsNullOrWhiteSpace($ReleaseExtractionRoot)) {
        $verifyArgs.ExtractionRoot = $ReleaseExtractionRoot
    }
    $releaseResult = & $releaseVerifier @verifyArgs
    $releaseResult | Format-List
    Write-Host "Release Green: ZIP=$($releaseResult.ZipPath) SHA256=$($releaseResult.ZipSHA256)"
}
