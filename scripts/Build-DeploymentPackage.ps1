<#
.SYNOPSIS
Builds the portable Windows x64 deployment ZIP.

.DESCRIPTION
Builds both converter executables, publishes Email Review Viewer as a
self-contained single file, assembles a clean release folder, writes SHA-256
checksums, and creates a stable-order ZIP with normalized entry timestamps.
#>
[CmdletBinding()]
param(
    [string]$OutputRoot = 'C:\Users\pbush\OneDrive - Perfection Learning\Documents\AI\Cursor Output\PurviewTeamsPstToHtmlApp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$readmePath = Join-Path $repoRoot 'README.md'
$buildScriptPath = Join-Path $repoRoot 'build.ps1'
$viewerProjectPath = Join-Path $repoRoot 'EmailReviewViewer\EmailReviewViewer.App\EmailReviewViewer.App.csproj'
$converterPath = Join-Path $repoRoot 'build\PurviewTeamsPstToHtmlConverter.exe'
$debugConverterPath = Join-Path $repoRoot 'build\PurviewTeamsPstToHtmlConverter_Debug.exe'

function Resolve-DotnetSdkPath {
    $command = Get-Command dotnet -ErrorAction SilentlyContinue
    $candidates = @(
        (Join-Path $env:USERPROFILE '.dotnet\dotnet.exe'),
        $(if ($command) { $command.Source } else { $null })
    ) | Where-Object {
        $_ -and (Test-Path -LiteralPath $_ -PathType Leaf)
    } | Select-Object -Unique

    foreach ($candidate in $candidates) {
        $sdks = @(& $candidate --list-sdks 2>$null)
        if ($LASTEXITCODE -eq 0 -and $sdks.Count -gt 0) {
            return $candidate
        }
    }
    throw 'A .NET SDK is required, but no dotnet executable with an installed SDK was found.'
}

foreach ($path in @($readmePath, $buildScriptPath, $viewerProjectPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required source file missing: $path"
    }
}

$readmeText = [IO.File]::ReadAllText($readmePath)
$buildText = [IO.File]::ReadAllText($buildScriptPath)
[xml]$viewerProject = [IO.File]::ReadAllText($viewerProjectPath)

$readmeMatch = [regex]::Match($readmeText, '(?m)^\*\*Current version:\*\*\s*(\d+\.\d+\.\d+\.\d+)\s*$')
$buildMatch = [regex]::Match($buildText, "(?m)\[string\]\`$Version\s*=\s*'(\d+\.\d+\.\d+\.\d+)'")
if (-not $readmeMatch.Success -or -not $buildMatch.Success) {
    throw 'Could not determine the repository version from README.md and build.ps1.'
}

$version = $readmeMatch.Groups[1].Value
$viewerVersion = [string]$viewerProject.Project.PropertyGroup.Version
if ($buildMatch.Groups[1].Value -ne $version -or $viewerVersion -ne $version) {
    throw "Version mismatch. README=$version build=$($buildMatch.Groups[1].Value) viewer=$viewerVersion"
}

$pwsh = (Get-Command pwsh -ErrorAction Stop).Source
$dotnet = Resolve-DotnetSdkPath
if (-not (Get-Module -ListAvailable ps2exe)) {
    throw 'PS2EXE is required. Install-Module ps2exe -Scope CurrentUser'
}
if (-not [Environment]::Is64BitOperatingSystem) {
    throw 'Packaging requires 64-bit Windows.'
}

$outputRoot = [IO.Path]::GetFullPath($OutputRoot)
$releasesRoot = Join-Path $outputRoot 'Releases'
$packageName = "PurviewTeamsPstToHtmlApp-$version-win-x64"
$packageRoot = Join-Path $releasesRoot $packageName
$zipPath = Join-Path $releasesRoot "$packageName.zip"
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) "PurviewTeamsPstToHtmlApp-package-$PID"
$viewerPublishRoot = Join-Path $tempRoot 'viewer'

if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
}
[void][IO.Directory]::CreateDirectory($viewerPublishRoot)
[void][IO.Directory]::CreateDirectory($releasesRoot)

try {
    Write-Host "Building converter $version..."
    & $pwsh -NoProfile -ExecutionPolicy Bypass -File $buildScriptPath -Version $version
    if ($LASTEXITCODE -ne 0) { throw "Converter build failed with exit code $LASTEXITCODE." }

    foreach ($path in @($converterPath, $debugConverterPath)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Expected converter build output missing: $path"
        }
    }

    Write-Host 'Publishing self-contained single-file Email Review Viewer...'
    & $dotnet publish $viewerProjectPath `
        -c Release `
        -r win-x64 `
        --self-contained true `
        -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:DebugType=None `
        -p:DebugSymbols=false `
        "-p:Version=$version" `
        -o $viewerPublishRoot
    if ($LASTEXITCODE -ne 0) { throw "Viewer publish failed with exit code $LASTEXITCODE." }

    $publishedViewerPath = Join-Path $viewerPublishRoot 'EmailReviewViewer.App.exe'
    if (-not (Test-Path -LiteralPath $publishedViewerPath -PathType Leaf)) {
        throw "Published viewer executable missing: $publishedViewerPath"
    }

    if (Test-Path -LiteralPath $packageRoot) {
        Remove-Item -LiteralPath $packageRoot -Recurse -Force
    }
    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }

    $toolsRoot = Join-Path $packageRoot 'Tools'
    $viewerRoot = Join-Path $packageRoot 'EmailReviewViewer'
    [void][IO.Directory]::CreateDirectory($toolsRoot)
    [void][IO.Directory]::CreateDirectory($viewerRoot)

    Copy-Item -LiteralPath $converterPath -Destination (Join-Path $packageRoot 'PurviewTeamsPstToHtmlConverter.exe')
    Copy-Item -LiteralPath $debugConverterPath -Destination (Join-Path $toolsRoot 'PurviewTeamsPstToHtmlConverter_Debug.exe')
    Copy-Item -LiteralPath $publishedViewerPath -Destination (Join-Path $viewerRoot 'EmailReviewViewer.App.exe')

    $quickStart = @"
Purview PST Report Converter $version
Portable Windows x64 internal release

REQUIREMENTS
- Windows 10 or Windows 11, x64
- Classic Microsoft Outlook installed and registered for COM
- PowerShell 7 (the "pwsh" command) for the converter
- No .NET Desktop Runtime is required for Email Review Viewer

QUICK START
1. Extract the entire ZIP to a local folder. Keep the folder structure intact.
2. Close Outlook before conversion.
3. Run PurviewTeamsPstToHtmlConverter.exe.
4. Choose the Purview PST, output location, and Teams, Email, or both reports.
5. Start the conversion. The converter temporarily attaches the PST to Outlook.

OUTPUTS
- Teams report: <name>_Teams.html
- Email database: <name>_Email.db
- Logs: <name>_Teams.log and/or <name>_Email.log beside the outputs
- The Email Review Viewer opens the Email database automatically after conversion.

To review an existing Email database, run
EmailReviewViewer\EmailReviewViewer.App.exe and choose File > Open Database.

PRIVACY
PST contents and generated reports stay on this computer. The tools do not
upload review data or require an online service.

TROUBLESHOOTING
- Windows SmartScreen may warn because this internal build is unsigned. Confirm
  the file came from Perfection Learning before choosing More info > Run anyway.
- If "pwsh" is missing, install PowerShell 7 and reopen the converter.
- If Outlook COM is missing, install/repair classic Outlook and make it the
  registered desktop Outlook application. New Outlook alone is not sufficient.
- Close Outlook before retrying PST attach/detach errors.
- Conversion logs are written beside the selected report output as _Teams.log
  and/or _Email.log. The debug converter is under Tools\ for support use.
- Run Verify-Prerequisites.ps1 for a read-only prerequisite check.

This release is portable and unsigned. Do not separate the converter from the
EmailReviewViewer folder.
"@
    [IO.File]::WriteAllText(
        (Join-Path $packageRoot 'README.txt'),
        $quickStart,
        [Text.UTF8Encoding]::new($false)
    )

    $releaseNotes = @"
# Purview PST Report Converter $version

## Deployment
- Portable Windows x64 ZIP for internal use.
- Converter built as Release/no-console; support Debug build is under `Tools`.
- Email Review Viewer published self-contained as a single file. The target
  computer does not need the .NET 8 Desktop Runtime.
- Viewer SQLite native components are embedded and extracted by the .NET
  single-file host to the current user's temporary application cache.

## Product changes
- Packaged converter locates and opens the bundled Email Review Viewer.
- Viewer supports validated File > Open Database switching.
- Email review includes SQLite FTS5 search, folder counts, paging, sorting, and
  on-demand full-message detail.

## Prerequisites and caveats
- Converter requires Windows 10/11 x64, PowerShell 7, and classic Outlook COM.
- Close Outlook before conversion.
- Executables are not code-signed; Windows SmartScreen may show an internal-app
  warning. Verify `checksums.sha256` when transferring this package.
"@
    [IO.File]::WriteAllText(
        (Join-Path $packageRoot 'RELEASE_NOTES.md'),
        $releaseNotes,
        [Text.UTF8Encoding]::new($false)
    )

    $prerequisiteScript = @'
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-CurrentDirectoryAclWrite {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    $acl = Get-Acl -LiteralPath $PSScriptRoot
    $writeRights = [Security.AccessControl.FileSystemRights]'Write, Modify, FullControl, CreateFiles, CreateDirectories'
    $allowed = $false
    foreach ($rule in $acl.Access) {
        $applies = $identity.User -eq $rule.IdentityReference
        if (-not $applies) {
            try {
                $sid = $rule.IdentityReference.Translate([Security.Principal.SecurityIdentifier])
                $applies = $principal.IsInRole($sid)
            } catch {
                $applies = $false
            }
        }
        if (-not $applies -or -not ($rule.FileSystemRights -band $writeRights)) { continue }
        if ($rule.AccessControlType -eq [Security.AccessControl.AccessControlType]::Deny) { return $false }
        $allowed = $true
    }
    return $allowed
}

$checks = @(
    [pscustomobject]@{
        Requirement = 'Windows 10/11 x64'
        Passed = $env:OS -eq 'Windows_NT' -and [Environment]::Is64BitOperatingSystem
        Detail = [Environment]::OSVersion.VersionString
    },
    [pscustomobject]@{
        Requirement = 'PowerShell 7 (pwsh)'
        Passed = [bool](Get-Command pwsh -ErrorAction SilentlyContinue)
        Detail = $(if (Get-Command pwsh -ErrorAction SilentlyContinue) { (Get-Command pwsh).Source } else { 'Not found' })
    },
    [pscustomobject]@{
        Requirement = 'Classic Outlook COM'
        Passed = [bool](Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction SilentlyContinue)
        Detail = $(if (Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction SilentlyContinue) { 'Registered' } else { 'Not registered' })
    },
    [pscustomobject]@{
        Requirement = 'Package folder ACL permits writes'
        Passed = Test-CurrentDirectoryAclWrite
        Detail = $PSScriptRoot
    }
)

$checks | Select-Object @{Name='Status';Expression={if ($_.Passed) {'PASS'} else {'FAIL'}}}, Requirement, Detail | Format-Table -AutoSize
if ($checks.Passed -contains $false) {
    Write-Host 'One or more prerequisites need attention. See README.txt.' -ForegroundColor Yellow
    exit 1
}
Write-Host 'All prerequisites passed.' -ForegroundColor Green
'@
    [IO.File]::WriteAllText(
        (Join-Path $packageRoot 'Verify-Prerequisites.ps1'),
        $prerequisiteScript,
        [Text.UTF8Encoding]::new($false)
    )

    $prohibitedExtensions = @('.db', '.db-wal', '.db-shm', '.pst', '.log', '.ndjson', '.pdb')
    $prohibitedDirectoryNames = @('.git', '.cursor', '.superpowers', 'bin', 'obj')
    $packageFiles = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -File)
    $badFiles = @($packageFiles | Where-Object {
        $name = $_.Name.ToLowerInvariant()
        $extension = $_.Extension.ToLowerInvariant()
        $prohibitedExtensions -contains $extension -or
        $name.EndsWith('.db-wal') -or
        $name.EndsWith('.db-shm')
    })
    $badDirectories = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -Directory | Where-Object {
        $prohibitedDirectoryNames -contains $_.Name.ToLowerInvariant()
    })
    if ($badFiles.Count -gt 0 -or $badDirectories.Count -gt 0) {
        throw "Prohibited deployment artifacts found: $(@($badFiles.FullName) + @($badDirectories.FullName) -join '; ')"
    }

    $checksumPath = Join-Path $packageRoot 'checksums.sha256'
    $checksumLines = foreach ($file in ($packageFiles | Sort-Object FullName)) {
        $relativePath = $file.FullName.Substring($packageRoot.Length + 1).Replace('\', '/')
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        "$hash *$relativePath"
    }
    [IO.File]::WriteAllLines($checksumPath, $checksumLines, [Text.UTF8Encoding]::new($false))

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $fixedTimestamp = [DateTimeOffset]::new(2026, 7, 14, 0, 0, 0, [TimeSpan]::Zero)
    $zipStream = [IO.File]::Open($zipPath, [IO.FileMode]::CreateNew)
    try {
        $archive = [IO.Compression.ZipArchive]::new($zipStream, [IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            foreach ($file in (Get-ChildItem -LiteralPath $packageRoot -Recurse -File | Sort-Object FullName)) {
                $relativePath = $file.FullName.Substring($packageRoot.Length + 1).Replace('\', '/')
                $entry = $archive.CreateEntry("$packageName/$relativePath", [IO.Compression.CompressionLevel]::Optimal)
                $entry.LastWriteTime = $fixedTimestamp
                $inputStream = [IO.File]::OpenRead($file.FullName)
                $entryStream = $entry.Open()
                try {
                    $inputStream.CopyTo($entryStream)
                } finally {
                    $entryStream.Dispose()
                    $inputStream.Dispose()
                }
            }
        } finally {
            $archive.Dispose()
        }
    } finally {
        $zipStream.Dispose()
    }

    $finalFiles = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -File)
    $totalBytes = [int64]$(
        $sum = ($finalFiles | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $sum) { 0 } else { $sum }
    )
    $zip = Get-Item -LiteralPath $zipPath
    $zipHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash

    Write-Host ''
    Write-Host "Release folder: $packageRoot"
    Write-Host "ZIP:            $zipPath"
    Write-Host "Version:        $version"
    Write-Host "Files:          $($finalFiles.Count)"
    Write-Host ("Folder bytes:   {0:N0}" -f $totalBytes)
    Write-Host ("ZIP bytes:      {0:N0}" -f $zip.Length)
    Write-Host "ZIP SHA-256:    $zipHash"
} finally {
    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}
