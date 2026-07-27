[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PackageRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$packageRoot = [IO.Path]::GetFullPath($PackageRoot)
if (-not (Test-Path -LiteralPath $packageRoot -PathType Container)) {
    throw "Package folder not found: $packageRoot"
}

$requiredFiles = @(
    'PurviewTeamsPstToHtmlConverter.exe',
    'Tools\PurviewTeamsPstToHtmlConverter_Debug.exe',
    'EmailReviewViewer\EmailReviewViewer.App.exe',
    'README.txt',
    'RELEASE_NOTES.md',
    'Verify-Prerequisites.ps1',
    'checksums.sha256'
)
foreach ($relativePath in $requiredFiles) {
    $path = Join-Path $packageRoot $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required package file missing: $relativePath"
    }
}

$prohibitedExtensions = @('.db', '.db-wal', '.db-shm', '.pst', '.log', '.ndjson', '.pdb')
$prohibitedDirectoryNames = @('.git', '.cursor', '.superpowers', 'bin', 'obj')
$files = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -File)
$directories = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -Directory)

$badFiles = @($files | Where-Object {
    $name = $_.Name.ToLowerInvariant()
    $extension = $_.Extension.ToLowerInvariant()
    $prohibitedExtensions -contains $extension -or
    $name.EndsWith('.db-wal') -or
    $name.EndsWith('.db-shm')
})
if ($badFiles.Count -gt 0) {
    throw "Prohibited package file(s): $($badFiles.FullName -join '; ')"
}

$badDirectories = @($directories | Where-Object {
    $prohibitedDirectoryNames -contains $_.Name.ToLowerInvariant()
})
if ($badDirectories.Count -gt 0) {
    throw "Prohibited package directorie(s): $($badDirectories.FullName -join '; ')"
}

$checksumPath = Join-Path $packageRoot 'checksums.sha256'
$checksumLines = @(
    Get-Content -LiteralPath $checksumPath |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
)
$expectedChecksumCount = $files.Count - 1
if ($checksumLines.Count -ne $expectedChecksumCount) {
    throw "Checksum entry count mismatch. Expected $expectedChecksumCount, found $($checksumLines.Count)."
}

foreach ($line in $checksumLines) {
    if ($line -notmatch '^([0-9A-Fa-f]{64}) \*(.+)$') {
        throw "Invalid checksum line: $line"
    }
    $expectedHash = $Matches[1].ToUpperInvariant()
    $relativePath = $Matches[2].Replace('/', '\')
    $filePath = Join-Path $packageRoot $relativePath
    if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
        throw "Checksum references missing file: $relativePath"
    }
    $actualHash = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash
    if ($actualHash -ne $expectedHash) {
        throw "Checksum mismatch: $relativePath"
    }
}

$totalBytes = [int64]$(
    $sum = ($files | Measure-Object -Property Length -Sum).Sum
    if ($null -eq $sum) { 0 } else { $sum }
)

[pscustomobject]@{
    PackageRoot = $packageRoot
    FileCount = $files.Count
    TotalBytes = $totalBytes
    ChecksumsVerified = $checksumLines.Count
}
