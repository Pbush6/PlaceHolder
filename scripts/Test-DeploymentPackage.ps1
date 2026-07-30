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

$manifestEntries = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($line in $checksumLines) {
    if ($line -notmatch '^([0-9A-Fa-f]{64}) \*(.+)$') {
        throw "Invalid checksum line: $line"
    }
    $expectedHash = $Matches[1].ToUpperInvariant()
    $relativePath = $Matches[2].Replace('/', '\')
    if ([string]::IsNullOrWhiteSpace($relativePath)) {
        throw 'Checksum path cannot be empty.'
    }
    if ([IO.Path]::IsPathRooted($relativePath) -or $relativePath -match '^[A-Za-z]:') {
        throw "Manifest contains an absolute checksum path: $relativePath"
    }
    if (@($relativePath -split '[\\/]' | Where-Object { $_ -eq '..' }).Count -gt 0) {
        throw "Manifest checksum path contains parent traversal: $relativePath"
    }
    $resolvedPath = [IO.Path]::GetFullPath((Join-Path $packageRoot $relativePath))
    $packagePrefix = $packageRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolvedPath.StartsWith($packagePrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Manifest checksum path escapes the package root: $relativePath"
    }
    if ($manifestEntries.ContainsKey($relativePath)) {
        throw "Duplicate checksum entry: $relativePath"
    }
    $manifestEntries.Add($relativePath, $expectedHash)
}

$actualFiles = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($file in @($files | Where-Object { $_.FullName -ne $checksumPath })) {
    $relativePath = $file.FullName.Substring($packageRoot.Length).TrimStart('\', '/')
    $actualFiles.Add($relativePath, $file.FullName)
}

$missingEntries = @($actualFiles.Keys | Where-Object { -not $manifestEntries.ContainsKey($_) } | Sort-Object)
if ($missingEntries.Count -gt 0) {
    throw "Missing checksum entry for package file(s): $($missingEntries -join '; ')"
}
$missingFiles = @($manifestEntries.Keys | Where-Object { -not $actualFiles.ContainsKey($_) } | Sort-Object)
if ($missingFiles.Count -gt 0) {
    throw "Checksum entry references missing package file(s): $($missingFiles -join '; ')"
}

foreach ($relativePath in $manifestEntries.Keys) {
    $actualHash = (Get-FileHash -LiteralPath $actualFiles[$relativePath] -Algorithm SHA256).Hash
    if ($actualHash -ne $manifestEntries[$relativePath]) {
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
    ChecksumsVerified = $manifestEntries.Count
}
