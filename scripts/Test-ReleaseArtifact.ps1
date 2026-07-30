[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ZipPath,
    [Parameter(Mandatory = $true)][string]$ExpectedVersion,
    [Parameter(Mandatory = $true)][string]$ConverterInputPath,
    [Parameter(Mandatory = $true)][string]$DebugConverterInputPath,
    [Parameter(Mandatory = $true)][string]$ViewerInputPath,
    [string]$ExtractionRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-FileExists {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Label)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label not found: $Path"
    }
}

function Quote-NativeArgument {
    param([Parameter(Mandatory = $true)][string]$Value)
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Initialize-VerifierOwnedDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$RequestedPath,
        [Parameter(Mandatory = $true)][string[]]$ProtectedPaths
    )

    if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
        throw 'ExtractionRoot cannot be empty.'
    }
    if ($RequestedPath -match '(^|[\\/])\.\.([\\/]|$)') {
        throw "ExtractionRoot contains parent traversal: $RequestedPath"
    }

    $fullPath = [IO.Path]::GetFullPath($RequestedPath)
    $pathRoot = [IO.Path]::GetPathRoot($fullPath)
    if ($fullPath.TrimEnd('\', '/') -eq $pathRoot.TrimEnd('\', '/')) {
        throw "ExtractionRoot cannot be a filesystem root: $fullPath"
    }
    $candidateCanonical = $fullPath.TrimEnd('\', '/')
    $candidateBounded = $candidateCanonical + [IO.Path]::DirectorySeparatorChar
    foreach ($protectedPath in $ProtectedPaths) {
        $protectedCanonical = [IO.Path]::GetFullPath($protectedPath).TrimEnd('\', '/')
        $protectedBounded = $protectedCanonical + [IO.Path]::DirectorySeparatorChar
        if (
            $candidateCanonical.Equals($protectedCanonical, [StringComparison]::OrdinalIgnoreCase) -or
            $candidateBounded.StartsWith($protectedBounded, [StringComparison]::OrdinalIgnoreCase) -or
            $protectedBounded.StartsWith($candidateBounded, [StringComparison]::OrdinalIgnoreCase)
        ) {
            throw "ExtractionRoot overlaps protected path '$protectedCanonical': $fullPath"
        }
    }

    $markerName = '.purview-release-verifier-owned.json'
    $markerPath = Join-Path $fullPath $markerName
    if (Test-Path -LiteralPath $fullPath) {
        if (-not (Test-Path -LiteralPath $fullPath -PathType Container)) {
            throw "ExtractionRoot exists but is not a directory: $fullPath"
        }
        if (-not (Test-Path -LiteralPath $markerPath -PathType Leaf)) {
            throw "Existing ExtractionRoot is not verifier-owned and will not be deleted: $fullPath"
        }
        try {
            $marker = [IO.File]::ReadAllText($markerPath) | ConvertFrom-Json
        }
        catch {
            throw "Existing ExtractionRoot has an invalid verifier ownership marker: $fullPath"
        }
        if (
            $marker.Owner -ne 'PurviewTeamsPstToHtmlApp.Test-ReleaseArtifact' -or
            -not ([string]$marker.Path).Equals($fullPath, [StringComparison]::OrdinalIgnoreCase)
        ) {
            throw "Existing ExtractionRoot has an invalid verifier ownership marker: $fullPath"
        }
        Remove-Item -LiteralPath $fullPath -Recurse -Force
    }

    [void][IO.Directory]::CreateDirectory($fullPath)
    [ordered]@{
        Owner = 'PurviewTeamsPstToHtmlApp.Test-ReleaseArtifact'
        Path = $fullPath
        CreatedUtc = [datetime]::UtcNow.ToString('o')
    } | ConvertTo-Json | Set-Content -LiteralPath $markerPath -Encoding utf8NoBOM
    return $fullPath
}

function Get-DescendantProcessRecords {
    param([Parameter(Mandatory = $true)][int]$RootProcessId)

    $allProcesses = @(Get-CimInstance Win32_Process)
    $descendants = [Collections.Generic.List[object]]::new()
    $pendingParents = [Collections.Generic.Queue[int]]::new()
    $seen = [Collections.Generic.HashSet[int]]::new()
    $pendingParents.Enqueue($RootProcessId)
    [void]$seen.Add($RootProcessId)
    while ($pendingParents.Count -gt 0) {
        $parentId = $pendingParents.Dequeue()
        foreach ($child in @($allProcesses | Where-Object { [int]$_.ParentProcessId -eq $parentId })) {
            $childId = [int]$child.ProcessId
            if ($seen.Add($childId)) {
                $descendants.Add([pscustomobject]@{
                    ProcessId = $childId
                    ParentProcessId = [int]$child.ParentProcessId
                    Name = [string]$child.Name
                    CreationDate = [string]$child.CreationDate
                    CommandLine = [string]$child.CommandLine
                })
                $pendingParents.Enqueue($childId)
            }
        }
    }
    return @($descendants)
}

function Get-RunningOwnedProcesses {
    param([Parameter(Mandatory = $true)][object[]]$OwnedProcessRecords)

    if ($OwnedProcessRecords.Count -eq 0) { return @() }
    $currentById = @{}
    foreach ($process in @(Get-CimInstance Win32_Process)) {
        $currentById[[int]$process.ProcessId] = $process
    }
    return @(
        foreach ($owned in $OwnedProcessRecords) {
            $current = $currentById[[int]$owned.ProcessId]
            if (
                $null -ne $current -and
                [string]$current.CreationDate -eq [string]$owned.CreationDate -and
                [string]$current.Name -eq [string]$owned.Name
            ) {
                $current
            }
        }
    )
}

function Invoke-PackagedNoGuiSmoke {
    param(
        [Parameter(Mandatory = $true)][string]$ExecutablePath,
        [Parameter(Mandatory = $true)][string]$WorkingDirectory,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [switch]$CaptureOutput
    )

    if (Test-Path -LiteralPath $OutputDirectory) {
        throw "NoGui output directory unexpectedly exists in fresh verifier root: $OutputDirectory"
    }
    [void][IO.Directory]::CreateDirectory($OutputDirectory)
    $outputPath = Join-Path $OutputDirectory 'release-smoke.html'
    $logPath = Join-Path $OutputDirectory 'release-smoke.log'

    $psi = [Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = $ExecutablePath
    $psi.WorkingDirectory = $WorkingDirectory
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.Arguments = "-NoGui -UseSampleData -OutputPath $(Quote-NativeArgument $outputPath) -LogPath $(Quote-NativeArgument $logPath)"
    if ($CaptureOutput) {
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
    }

    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $psi
    $stdoutTask = $null
    $stderrTask = $null
    $stdout = ''
    $stderr = ''
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $started = $false
    $exitCode = $null
    $ownedProcessRecords = [Collections.Generic.Dictionary[int, object]]::new()
    $detectedOwnedOrphans = @()
    $remainingOwnedProcesses = @()
    try {
        if (-not $process.Start()) { throw "NoGui process did not start: $ExecutablePath" }
        $started = $true
        if ($CaptureOutput) {
            $stdoutTask = $process.StandardOutput.ReadToEndAsync()
            $stderrTask = $process.StandardError.ReadToEndAsync()
        }
        $deadline = [datetime]::UtcNow.AddSeconds(120)
        do {
            foreach ($owned in @(Get-DescendantProcessRecords -RootProcessId $process.Id)) {
                if (-not $ownedProcessRecords.ContainsKey([int]$owned.ProcessId)) {
                    $ownedProcessRecords.Add([int]$owned.ProcessId, $owned)
                }
            }
            if ($process.WaitForExit(250)) { break }
        } while ([datetime]::UtcNow -lt $deadline)
        if (-not $process.HasExited) {
            & taskkill.exe /PID $process.Id /T /F | Out-Null
            [void]$process.WaitForExit(10000)
            throw "NoGui process timed out: $ExecutablePath"
        }
        if ($CaptureOutput) {
            if ($stdoutTask.Wait(10000)) { $stdout = $stdoutTask.Result }
            if ($stderrTask.Wait(10000)) { $stderr = $stderrTask.Result }
        }
        if ($process.ExitCode -ne 0) {
            throw "NoGui process failed with exit code $($process.ExitCode): $ExecutablePath`n$stderr"
        }
        $exitCode = $process.ExitCode
    }
    finally {
        if ($started -and -not $process.HasExited) {
            & taskkill.exe /PID $process.Id /T /F | Out-Null
            [void]$process.WaitForExit(10000)
        }
        $detectedOwnedOrphans = @(Get-RunningOwnedProcesses -OwnedProcessRecords @($ownedProcessRecords.Values))
        foreach ($owned in $detectedOwnedOrphans) {
            Stop-Process -Id ([int]$owned.ProcessId) -Force -ErrorAction SilentlyContinue
        }
        $remainingOwnedProcesses = @(Get-RunningOwnedProcesses -OwnedProcessRecords @($ownedProcessRecords.Values))
        $stopwatch.Stop()
        $process.Dispose()
        if ($detectedOwnedOrphans.Count -gt 0) {
            throw "NoGui smoke left tool-owned descendant process(es): $($detectedOwnedOrphans.ProcessId -join ', ')"
        }
    }

    $requiredArtifacts = @(
        'release-smoke_Teams.html',
        'release-smoke_Email.db',
        'release-smoke_Calendar.html',
        'release-smoke_Contacts.html',
        'release-smoke_Teams.log',
        'release-smoke_Email.log',
        'release-smoke_Calendar.log',
        'release-smoke_Contacts.log'
    )
    foreach ($name in $requiredArtifacts) {
        Assert-FileExists -Path (Join-Path $OutputDirectory $name) -Label 'NoGui artifact'
    }
    $teamsLog = [IO.File]::ReadAllText((Join-Path $OutputDirectory 'release-smoke_Teams.log'))
    if ($teamsLog -notmatch 'exported=6' -or $teamsLog -notmatch 'itemReadFailures=0') {
        throw "NoGui smoke summary failed: $ExecutablePath"
    }
    if ($CaptureOutput) {
        foreach ($expected in @(
            'CONVERSION_RESULT|',
            'ItemsExported=6',
            'TeamsItemsExported=2',
            'EmailItemsExported=2',
            'CalendarItemsExported=1',
            'ContactsItemsExported=1'
        )) {
            if ($stdout -notmatch [regex]::Escape($expected)) {
                throw "Debug NoGui stdout missing '$expected': $ExecutablePath"
            }
        }
    }

    $result = [pscustomobject]@{
        ExecutablePath = $ExecutablePath
        ExitCode = $exitCode
        ElapsedMilliseconds = $stopwatch.ElapsedMilliseconds
        ArtifactCount = $requiredArtifacts.Count
        CapturedMachineResult = [bool]$CaptureOutput
        OwnedProcessIds = @($ownedProcessRecords.Keys | Sort-Object)
        OwnedPowerShellDescendants = @(
            $ownedProcessRecords.Values |
                Where-Object { $_.Name -in @('powershell.exe', 'pwsh.exe') }
        ).Count
        RemainingOwnedProcesses = @($remainingOwnedProcesses).Count
    }
    return $result
}

$zipPath = [IO.Path]::GetFullPath($ZipPath)
$converterInputPath = [IO.Path]::GetFullPath($ConverterInputPath)
$debugConverterInputPath = [IO.Path]::GetFullPath($DebugConverterInputPath)
$viewerInputPath = [IO.Path]::GetFullPath($ViewerInputPath)
Assert-FileExists -Path $zipPath -Label 'Release ZIP'
Assert-FileExists -Path $converterInputPath -Label 'Converter build input'
Assert-FileExists -Path $debugConverterInputPath -Label 'Debug converter build input'
Assert-FileExists -Path $viewerInputPath -Label 'Email Reviewer build input'

if ($PSBoundParameters.ContainsKey('ExtractionRoot') -and [string]::IsNullOrWhiteSpace($ExtractionRoot)) {
    throw 'ExtractionRoot cannot be empty.'
}
if ([string]::IsNullOrWhiteSpace($ExtractionRoot)) {
    $ExtractionRoot = Join-Path ([IO.Path]::GetTempPath()) "PurviewTeamsPstToHtmlApp-release-verify-$PID"
}
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$currentDirectory = [IO.Path]::GetFullPath((Get-Location).ProviderPath)
$extractionRoot = Initialize-VerifierOwnedDirectory `
    -RequestedPath $ExtractionRoot `
    -ProtectedPaths @($repoRoot, $currentDirectory)
Expand-Archive -LiteralPath $zipPath -DestinationPath $extractionRoot

$packageName = "PurviewTeamsPstToHtmlApp-$ExpectedVersion-win-x64"
$packageRoot = Join-Path $extractionRoot $packageName
if (-not (Test-Path -LiteralPath $packageRoot -PathType Container)) {
    throw "Expected package root missing after extraction: $packageRoot"
}

$packageVerifier = Join-Path $PSScriptRoot 'Test-DeploymentPackage.ps1'
$packageResult = & $packageVerifier -PackageRoot $packageRoot

$artifactBindings = @(
    [pscustomobject]@{
        Name = 'Release converter'
        InputPath = $converterInputPath
        PackagePath = Join-Path $packageRoot 'PurviewTeamsPstToHtmlConverter.exe'
    },
    [pscustomobject]@{
        Name = 'Debug converter'
        InputPath = $debugConverterInputPath
        PackagePath = Join-Path $packageRoot 'Tools\PurviewTeamsPstToHtmlConverter_Debug.exe'
    },
    [pscustomobject]@{
        Name = 'Email Reviewer'
        InputPath = $viewerInputPath
        PackagePath = Join-Path $packageRoot 'EmailReviewViewer\EmailReviewViewer.App.exe'
    }
)

foreach ($binding in $artifactBindings) {
    Assert-FileExists -Path $binding.PackagePath -Label $binding.Name
    $file = Get-Item -LiteralPath $binding.PackagePath
    if ($file.VersionInfo.FileVersion -ne $ExpectedVersion) {
        throw "$($binding.Name) file version mismatch. Expected=$ExpectedVersion Actual=$($file.VersionInfo.FileVersion)"
    }
    if ($file.VersionInfo.ProductVersion -notlike "$ExpectedVersion*") {
        throw "$($binding.Name) product version mismatch. Expected=$ExpectedVersion Actual=$($file.VersionInfo.ProductVersion)"
    }
    $binding | Add-Member -NotePropertyName InputSHA256 -NotePropertyValue ((Get-FileHash -LiteralPath $binding.InputPath -Algorithm SHA256).Hash)
    $binding | Add-Member -NotePropertyName PackageSHA256 -NotePropertyValue ((Get-FileHash -LiteralPath $binding.PackagePath -Algorithm SHA256).Hash)
    if ($binding.InputSHA256 -ne $binding.PackageSHA256) {
        throw "$($binding.Name) is not byte-for-byte identical to its build input."
    }
}

$releaseSmoke = Invoke-PackagedNoGuiSmoke `
    -ExecutablePath (Join-Path $packageRoot 'PurviewTeamsPstToHtmlConverter.exe') `
    -WorkingDirectory $packageRoot `
    -OutputDirectory (Join-Path $extractionRoot 'release-nogui-smoke')

$debugSmokeRoot = Join-Path $extractionRoot 'debug-nogui-package'
[void][IO.Directory]::CreateDirectory($debugSmokeRoot)
$debugSmokeExe = Join-Path $debugSmokeRoot 'PurviewTeamsPstToHtmlConverter_Debug.exe'
Copy-Item -LiteralPath (Join-Path $packageRoot 'Tools\PurviewTeamsPstToHtmlConverter_Debug.exe') -Destination $debugSmokeExe
Copy-Item -LiteralPath (Join-Path $packageRoot 'EmailReviewViewer') -Destination (Join-Path $debugSmokeRoot 'EmailReviewViewer') -Recurse
$debugSmoke = Invoke-PackagedNoGuiSmoke `
    -ExecutablePath $debugSmokeExe `
    -WorkingDirectory $debugSmokeRoot `
    -OutputDirectory (Join-Path $extractionRoot 'debug-nogui-smoke') `
    -CaptureOutput

[pscustomobject]@{
    ZipPath = $zipPath
    ZipBytes = (Get-Item -LiteralPath $zipPath).Length
    ZipSHA256 = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash
    ExpectedVersion = $ExpectedVersion
    ExtractionRoot = $extractionRoot
    PackageFiles = $packageResult.FileCount
    ManifestChecksumsVerified = $packageResult.ChecksumsVerified
    ArtifactBindings = $artifactBindings
    ReleaseNoGui = $releaseSmoke
    DebugNoGui = $debugSmoke
    RemainingOwnedProcesses = $releaseSmoke.RemainingOwnedProcesses + $debugSmoke.RemainingOwnedProcesses
}
