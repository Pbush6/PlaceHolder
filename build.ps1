<#
.SYNOPSIS
Embed the core converter into the launcher and build the Debug/Release EXEs.

.DESCRIPTION
Reproducible build for PurviewTeamsPstToHtmlConverter:
1. Base64-embeds src\Convert-PurviewTeamsPstToHtml.ps1 into the launcher's
   $script:EmbeddedCoreBase64 line (in place, preserving LF endings and no BOM).
2. Builds Debug (console) and Release (-NoConsole) EXEs with PS2EXE.

Base64 contains no quotes/backslashes and the embed uses a MatchEvaluator, so the
launcher's own $-prefixed variables are never re-interpreted (the bug that once
corrupted the embed line).
#>
[CmdletBinding()]
param(
    [string]$Version = '1.2.1.0'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$proj = $PSScriptRoot
$core = Join-Path $proj 'src\Convert-PurviewTeamsPstToHtml.ps1'
$launcher = Join-Path $proj 'src\Start-PurviewTeamsPstToHtmlApp.ps1'
$icon = Join-Path $proj 'assets\PurviewTeamsPstToHtmlConverter_icon.ico'
$buildDir = Join-Path $proj 'build'
$debugOut = Join-Path $buildDir 'PurviewTeamsPstToHtmlConverter_Debug.exe'
$releaseOut = Join-Path $buildDir 'PurviewTeamsPstToHtmlConverter.exe'

foreach ($f in @($core, $launcher, $icon)) {
    if (-not (Test-Path -LiteralPath $f)) { throw "Required file missing: $f" }
}

function Assert-SourceContains {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Snippet,
        [Parameter(Mandatory = $true)][string]$Label
    )
    $text = [IO.File]::ReadAllText($Path)
    if ($text -notmatch [regex]::Escape($Snippet)) {
        throw "Re-sync required before build: $Label missing from $Path."
    }
}

Assert-SourceContains -Path $core -Snippet 'function Get-ReportOutputPaths' -Label 'core helper Get-ReportOutputPaths'
Assert-SourceContains -Path $core -Snippet 'function Get-ItemReportBucket' -Label 'core helper Get-ItemReportBucket'
Assert-SourceContains -Path $core -Snippet 'function Write-CalendarHtmlReport' -Label 'core writer Write-CalendarHtmlReport'
Assert-SourceContains -Path $core -Snippet 'function Write-ContactsHtmlReport' -Label 'core writer Write-ContactsHtmlReport'
Assert-SourceContains -Path $launcher -Snippet 'function Get-ReportOutputPaths' -Label 'launcher helper Get-ReportOutputPaths'
[void][IO.Directory]::CreateDirectory($buildDir)

# --- Embed core ---
$coreBytes = [IO.File]::ReadAllBytes($core)
$b64 = [Convert]::ToBase64String($coreBytes)
$newLine = "`$script:EmbeddedCoreBase64 = '$b64'"

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$text = [IO.File]::ReadAllText($launcher, $utf8NoBom)
$evaluator = [System.Text.RegularExpressions.MatchEvaluator] { param($m) $newLine }
$pattern = "(?m)^\`$script:EmbeddedCoreBase64 = '[^']*'"
$updated = [regex]::Replace($text, $pattern, $evaluator)
if ($updated -eq $text) {
    if ($text -notmatch "(?m)^\`$script:EmbeddedCoreBase64 = ") {
        throw 'Could not find $script:EmbeddedCoreBase64 line in launcher.'
    }
    Write-Host 'Embed line unchanged (core already current).'
}
[IO.File]::WriteAllText($launcher, $updated, $utf8NoBom)
Write-Host "Embedded core ($($coreBytes.Length) bytes -> $($b64.Length) base64 chars)."

# --- Build EXEs ---
Import-Module ps2exe -Force
$common = @{
    InputFile   = $launcher
    IconFile    = $icon
    Description = 'Converts Microsoft Purview PST exports into Teams, Email, Calendar, and Contacts reports'
    Company     = 'Perfection Learning'
    Product     = 'Purview PST Report Converter'
    Copyright   = 'Perfection Learning'
    Version     = $Version
}
Invoke-PS2EXE @common -OutputFile $debugOut -STA -X64 -Title 'Purview Teams PST to HTML Converter (Debug)'
Invoke-PS2EXE @common -OutputFile $releaseOut -NoConsole -NoOutput -STA -X64 -Title 'Purview Teams PST to HTML Converter'

Write-Host "=== built $Version ==="
Get-Item -LiteralPath $debugOut, $releaseOut | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize
