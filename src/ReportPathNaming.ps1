# Source of truth; build.ps1 will inline into core/launcher for packaging.
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
