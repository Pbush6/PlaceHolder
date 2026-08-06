Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Release artifact verification contracts' {
    BeforeAll {
        $script:repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
        $script:verifierPath = Join-Path $script:repoRoot 'scripts\Test-ReleaseArtifact.ps1'
        $script:packageVerifierPath = Join-Path $script:repoRoot 'scripts\Test-DeploymentPackage.ps1'
        $script:suitePath = Join-Path $script:repoRoot 'scripts\Run-VerificationSuite.ps1'
        $script:packageBuilderPath = Join-Path $script:repoRoot 'scripts\Build-DeploymentPackage.ps1'
        $script:stopTreePath = Join-Path $script:repoRoot 'scripts\Verify-StopProcessTreeRegression.ps1'
        $script:contractsPath = Join-Path $script:repoRoot 'docs\DATA_CONTRACTS.md'

        function New-TestPackage {
            param([Parameter(Mandatory = $true)][string]$Root)
            [void][IO.Directory]::CreateDirectory((Join-Path $Root 'Tools'))
            [void][IO.Directory]::CreateDirectory((Join-Path $Root 'EmailReviewViewer'))
            $content = [ordered]@{
                'PurviewTeamsPstToHtmlConverter.exe' = 'release'
                'Tools\PurviewTeamsPstToHtmlConverter_Debug.exe' = 'debug'
                'EmailReviewViewer\EmailReviewViewer.App.exe' = 'viewer'
                'README.txt' = 'readme'
                'RELEASE_NOTES.md' = 'notes'
                'Verify-Prerequisites.ps1' = 'prerequisites'
            }
            foreach ($entry in $content.GetEnumerator()) {
                [IO.File]::WriteAllText((Join-Path $Root $entry.Key), $entry.Value)
            }
            $lines = foreach ($relativePath in $content.Keys) {
                $hash = (Get-FileHash -LiteralPath (Join-Path $Root $relativePath) -Algorithm SHA256).Hash
                "$hash *$($relativePath.Replace('\', '/'))"
            }
            [IO.File]::WriteAllLines((Join-Path $Root 'checksums.sha256'), $lines)
        }

        function Invoke-UnsafeExtractionProbe {
            param(
                [AllowEmptyString()][string]$ExtractionRoot,
                [string]$WorkingDirectory
            )
            $fixtureRoot = Join-Path $TestDrive 'extraction-fixture'
            [void][IO.Directory]::CreateDirectory($fixtureRoot)
            $zipPath = Join-Path $fixtureRoot 'fixture.zip'
            $inputPath = Join-Path $fixtureRoot 'input.bin'
            [IO.File]::WriteAllText($zipPath, 'not a zip')
            [IO.File]::WriteAllText($inputPath, 'input')
            $pwshArguments = @()
            if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
                $pwshArguments += @('-WorkingDirectory', $WorkingDirectory)
            }
            $pwshArguments += @(
                '-NoProfile', '-File', $script:verifierPath,
                '-ZipPath', $zipPath,
                '-ExpectedVersion', '1.3.0.0',
                '-ConverterInputPath', $inputPath,
                '-DebugConverterInputPath', $inputPath,
                '-ViewerInputPath', $inputPath,
                '-ExtractionRoot', $ExtractionRoot
            )
            $output = @(
                & pwsh @pwshArguments 2>&1
            )
            return [pscustomobject]@{
                ExitCode = $LASTEXITCODE
                Output = $output -join "`n"
            }
        }
    }

    It 'provides a dedicated verifier with explicit artifact and build-input parameters' {
        (Test-Path -LiteralPath $script:verifierPath -PathType Leaf) | Should -BeTrue
        $text = Get-Content -LiteralPath $script:verifierPath -Raw
        foreach ($parameter in @(
            'ZipPath',
            'ExpectedVersion',
            'ConverterInputPath',
            'DebugConverterInputPath',
            'ViewerInputPath'
        )) {
            $text | Should -Match "(?s)\[Parameter\(Mandatory\s*=\s*\`$true\)\]\s*\[string\]\`$$parameter"
        }
        $text | Should -Match 'Expand-Archive'
        $text | Should -Match 'Get-FileHash'
        $text | Should -Match 'FileVersion'
        $text | Should -Match 'ProductVersion'
        $text | Should -Match 'Test-DeploymentPackage\.ps1'
        (Get-Content -LiteralPath $script:packageBuilderPath -Raw) | Should -Match 'BuildInputs'
    }

    It 'fails when the explicitly requested release ZIP is missing' {
        $missingZip = Join-Path $TestDrive 'missing-release.zip'
        $output = @(
            & pwsh -NoProfile -File $script:verifierPath `
                -ZipPath $missingZip `
                -ExpectedVersion '1.3.0.0' `
                -ConverterInputPath $script:verifierPath `
                -DebugConverterInputPath $script:verifierPath `
                -ViewerInputPath $script:verifierPath 2>&1
        )
        $LASTEXITCODE | Should -Not -Be 0
        ($output -join "`n") | Should -Match 'Release ZIP not found'
    }

    It 'requires an explicit ZIP when canonical verification runs in release mode' {
        $text = Get-Content -LiteralPath $script:suitePath -Raw
        $text | Should -Match '\[switch\]\$Release'
        $text | Should -Match '\[string\]\$ReleaseZipPath'
        $text | Should -Match 'ReleaseZipPath is required'
        $text | Should -Match 'Test-ReleaseArtifact\.ps1'
    }

    It 'rejects an existing unowned extraction directory without deleting its sentinel' {
        $unowned = Join-Path $TestDrive 'unowned-extraction'
        [void][IO.Directory]::CreateDirectory($unowned)
        $sentinel = Join-Path $unowned 'do-not-delete.txt'
        [IO.File]::WriteAllText($sentinel, 'sentinel')

        $result = Invoke-UnsafeExtractionProbe -ExtractionRoot $unowned

        $result.ExitCode | Should -Not -Be 0
        $result.Output | Should -Match 'not verifier-owned'
        (Test-Path -LiteralPath $sentinel -PathType Leaf) | Should -BeTrue
        [IO.File]::ReadAllText($sentinel) | Should -Be 'sentinel'
    }

    It 'rejects extraction paths containing parent traversal' {
        $unsafePath = Join-Path $TestDrive 'nested\..\escape'
        $result = Invoke-UnsafeExtractionProbe -ExtractionRoot $unsafePath

        $result.ExitCode | Should -Not -Be 0
        $result.Output | Should -Match 'parent traversal'
    }

    It 'rejects an empty extraction path' {
        $result = Invoke-UnsafeExtractionProbe -ExtractionRoot ''

        $result.ExitCode | Should -Not -Be 0
        $result.Output | Should -Match 'cannot be empty'
    }

    It 'rejects a filesystem root without deleting an existing root sentinel' {
        $filesystemRoot = [IO.Path]::GetPathRoot($TestDrive)
        $sentinel = Join-Path $filesystemRoot 'Windows'
        (Test-Path -LiteralPath $sentinel -PathType Container) | Should -BeTrue

        $result = Invoke-UnsafeExtractionProbe -ExtractionRoot $filesystemRoot

        $result.ExitCode | Should -Not -Be 0
        $result.Output | Should -Match 'filesystem root'
        (Test-Path -LiteralPath $sentinel -PathType Container) | Should -BeTrue
    }

    It 'rejects the repository directory without changing its README sentinel' {
        $sentinel = Join-Path $script:repoRoot 'README.md'
        $beforeHash = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash

        $result = Invoke-UnsafeExtractionProbe -ExtractionRoot $script:repoRoot

        $result.ExitCode | Should -Not -Be 0
        $result.Output | Should -Match 'overlaps protected path'
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -Be $beforeHash
    }

    It 'rejects a marker-bearing child inside the repository before deleting its sentinel' {
        $child = Join-Path $script:repoRoot '.release-verifier-containment-child'
        $sentinel = Join-Path $child 'do-not-delete.txt'
        try {
            [void][IO.Directory]::CreateDirectory($child)
            [IO.File]::WriteAllText($sentinel, 'child sentinel')
            $mixedPath = $child.ToUpperInvariant().TrimEnd('\') + '\'
            [ordered]@{
                Owner = 'PurviewTeamsPstToHtmlApp.Test-ReleaseArtifact'
                Path = $mixedPath
                CreatedUtc = [datetime]::UtcNow.ToString('o')
            } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $child '.purview-release-verifier-owned.json')

            $result = Invoke-UnsafeExtractionProbe -ExtractionRoot $mixedPath

            $result.ExitCode | Should -Not -Be 0
            $result.Output | Should -Match 'overlaps protected path'
            (Test-Path -LiteralPath $sentinel -PathType Leaf) | Should -BeTrue
            [IO.File]::ReadAllText($sentinel) | Should -Be 'child sentinel'
        }
        finally {
            if (Test-Path -LiteralPath $child) {
                Remove-Item -LiteralPath $child -Recurse -Force
            }
        }
    }

    It 'rejects a marker-bearing ancestor of the current directory before deleting its sentinel' {
        $ancestor = Join-Path $TestDrive 'protected-ancestor'
        $workingDirectory = Join-Path $ancestor 'nested-current'
        $sentinel = Join-Path $ancestor 'do-not-delete.txt'
        $requestedAncestor = $ancestor.TrimEnd('\') + '\'
        [void][IO.Directory]::CreateDirectory($workingDirectory)
        [IO.File]::WriteAllText($sentinel, 'ancestor sentinel')
        [ordered]@{
            Owner = 'PurviewTeamsPstToHtmlApp.Test-ReleaseArtifact'
            Path = $requestedAncestor
            CreatedUtc = [datetime]::UtcNow.ToString('o')
        } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $ancestor '.purview-release-verifier-owned.json')

        $result = Invoke-UnsafeExtractionProbe `
            -ExtractionRoot $requestedAncestor `
            -WorkingDirectory $workingDirectory

        $result.ExitCode | Should -Not -Be 0
        $result.Output | Should -Match 'overlaps protected path'
        (Test-Path -LiteralPath $sentinel -PathType Leaf) | Should -BeTrue
        [IO.File]::ReadAllText($sentinel) | Should -Be 'ancestor sentinel'
    }

    It 'scopes orphan detection to process descendants owned by each smoke run' {
        $text = Get-Content -LiteralPath $script:verifierPath -Raw
        $text | Should -Match 'ParentProcessId'
        $text | Should -Match 'OwnedProcess'
        $text | Should -Match 'powershell|pwsh'
        $text | Should -Not -Match "Get-Process -Name 'PurviewTeamsPstToHtmlConverter'"
    }

    It 'disposes every launched smoke process and reports a measured survivor count' {
        $text = Get-Content -LiteralPath $script:verifierPath -Raw
        $smokeMatch = [regex]::Match($text, '(?s)function Invoke-PackagedNoGuiSmoke \{(.+?)\r?\n\}')
        $smokeMatch.Success | Should -BeTrue
        $body = $smokeMatch.Groups[1].Value

        $body | Should -Match '(?s)finally\s*\{.*\$process\.Dispose\(\)'
        $body | Should -Not -Match 'RemainingOwnedProcesses\s*=\s*0'
        $body | Should -Match 'RemainingOwnedProcesses\s*=\s*@\(\$remainingOwnedProcesses\)\.Count'
        $body | Should -Match 'Get-RunningOwnedProcesses'
    }

    It 'accepts a complete one-to-one package checksum manifest' {
        $packageRoot = Join-Path $TestDrive 'valid-package'
        New-TestPackage -Root $packageRoot

        $result = & $script:packageVerifierPath -PackageRoot $packageRoot

        $result.FileCount | Should -Be 7
        $result.ChecksumsVerified | Should -Be 6
    }

    It 'rejects a checksum entry containing parent traversal' {
        $packageRoot = Join-Path $TestDrive 'traversal-package'
        New-TestPackage -Root $packageRoot
        $outsidePath = Join-Path $TestDrive 'outside.txt'
        [IO.File]::WriteAllText($outsidePath, 'readme')
        $lines = @(Get-Content -LiteralPath (Join-Path $packageRoot 'checksums.sha256'))
        $outsideHash = (Get-FileHash -LiteralPath $outsidePath -Algorithm SHA256).Hash
        $lines = @($lines | Where-Object { $_ -notmatch 'README\.txt$' })
        $lines += "$outsideHash *../outside.txt"
        [IO.File]::WriteAllLines((Join-Path $packageRoot 'checksums.sha256'), $lines)

        { & $script:packageVerifierPath -PackageRoot $packageRoot } |
            Should -Throw '*parent traversal*'
    }

    It 'rejects absolute checksum paths' {
        $packageRoot = Join-Path $TestDrive 'absolute-package'
        New-TestPackage -Root $packageRoot
        $outsidePath = Join-Path $TestDrive 'absolute.txt'
        [IO.File]::WriteAllText($outsidePath, 'readme')
        $lines = @(Get-Content -LiteralPath (Join-Path $packageRoot 'checksums.sha256'))
        $outsideHash = (Get-FileHash -LiteralPath $outsidePath -Algorithm SHA256).Hash
        $lines = @($lines | Where-Object { $_ -notmatch 'README\.txt$' })
        $lines += "$outsideHash *$outsidePath"
        [IO.File]::WriteAllLines((Join-Path $packageRoot 'checksums.sha256'), $lines)

        { & $script:packageVerifierPath -PackageRoot $packageRoot } |
            Should -Throw '*absolute checksum path*'
    }

    It 'rejects case-insensitive duplicate entries that hide a missing file' {
        $packageRoot = Join-Path $TestDrive 'duplicate-package'
        New-TestPackage -Root $packageRoot
        $lines = @(Get-Content -LiteralPath (Join-Path $packageRoot 'checksums.sha256'))
        $duplicate = @($lines | Where-Object { $_ -match 'PurviewTeamsPstToHtmlConverter\.exe$' })[0]
        $lines = @($lines | Where-Object { $_ -notmatch 'README\.txt$' })
        $lines += $duplicate.ToUpperInvariant()
        [IO.File]::WriteAllLines((Join-Path $packageRoot 'checksums.sha256'), $lines)

        { & $script:packageVerifierPath -PackageRoot $packageRoot } |
            Should -Throw '*Duplicate checksum entry*'
    }

    It 'rejects package files missing from the checksum manifest' {
        $packageRoot = Join-Path $TestDrive 'missing-entry-package'
        New-TestPackage -Root $packageRoot
        $lines = @(
            Get-Content -LiteralPath (Join-Path $packageRoot 'checksums.sha256') |
                Where-Object { $_ -notmatch 'README\.txt$' }
        )
        [IO.File]::WriteAllLines((Join-Path $packageRoot 'checksums.sha256'), $lines)

        { & $script:packageVerifierPath -PackageRoot $packageRoot } |
            Should -Throw '*Missing checksum entry*README.txt*'
    }

    It 'rejects extra unmanifested package files' {
        $packageRoot = Join-Path $TestDrive 'extra-file-package'
        New-TestPackage -Root $packageRoot
        [IO.File]::WriteAllText((Join-Path $packageRoot 'unexpected.txt'), 'extra')

        { & $script:packageVerifierPath -PackageRoot $packageRoot } |
            Should -Throw '*Missing checksum entry*unexpected.txt*'
    }

    It 'keeps stop-process-tree outputs outside the repository by default' {
        $text = Get-Content -LiteralPath $script:stopTreePath -Raw
        $text | Should -Not -Match "\\.\\.[\\/]test-output"
        $text | Should -Match '\[string\]\$OutputDirectory'
        $text | Should -Match '\[string\]\$LogPath'
        (Test-Path -LiteralPath (Join-Path $script:repoRoot 'test-output\verify-stop-process-tree.log')) | Should -BeFalse
        (Test-Path -LiteralPath (Join-Path $script:repoRoot 'test-output\verify-stop-process-tree.html')) | Should -BeFalse
    }

    It 'documents completed 1.3.0.0 version alignment without stale drift language' {
        $text = Get-Content -LiteralPath $script:contractsPath -Raw
        $text | Should -Match '\*\*Version:\*\*\s*1\.3\.0\.0'
        $text | Should -Match '1\.3\.0\.0'
        $text | Should -Not -Match 'Current drift to resolve before release'
    }
}
