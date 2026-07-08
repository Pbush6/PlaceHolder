Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Purview Teams PST to HTML verification suite' {
    BeforeAll {
        $script:repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
        $script:launcherPath = Join-Path $script:repoRoot 'src\Start-PurviewTeamsPstToHtmlApp.ps1'
        $script:corePath = Join-Path $script:repoRoot 'src\Convert-PurviewTeamsPstToHtml.ps1'
        $script:buildScriptPath = Join-Path $script:repoRoot 'build.ps1'
        $script:regressionScriptPath = Join-Path $script:repoRoot 'scripts\Verify-StopProcessTreeRegression.ps1'
        $script:buildDir = Join-Path $script:repoRoot 'build'
        $script:invokePwshScriptCapture = {
            param(
                [Parameter(Mandatory = $true)][string]$FilePath,
                [string[]]$Arguments = @()
            )

            $mergedOutput = (& pwsh -NoProfile -File $FilePath @Arguments 2>&1 | Out-String)
            $mergedOutput = $mergedOutput -replace "`e\[[0-9;]*m", ''
            [pscustomobject]@{
                ExitCode = $LASTEXITCODE
                StdOut = $mergedOutput
                StdErr = ''
            }
        }
    }

    It 'parses the launcher, core, and regression script cleanly' {
        foreach ($path in @($script:launcherPath, $script:corePath, $script:regressionScriptPath)) {
            $null = $tokens = $errors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors) | Out-Null
            $errors | Should -BeNullOrEmpty -Because "parser errors in $path"
        }
    }

    It 'keeps the embedded launcher core in sync with the source core' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw
        $match = [regex]::Match($launcherText, '(?m)^\$script:EmbeddedCoreBase64 = ''([^'']*)''')
        $match.Success | Should -BeTrue -Because 'launcher must embed the core converter'

        $embeddedBytes = [Convert]::FromBase64String($match.Groups[1].Value)
        $embeddedHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($embeddedBytes))
        $coreHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([IO.File]::ReadAllBytes($script:corePath)))
        $embeddedHash | Should -Be $coreHash
    }

    It 'runs the core sample path successfully' {
        $reportPath = Join-Path $TestDrive 'core-sample.html'
        $logPath = Join-Path $TestDrive 'core-sample.log'

        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @('-UseSampleData', '-NoPrompt', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath $reportPath) | Should -BeTrue
        (Test-Path -LiteralPath $logPath) | Should -BeTrue
        $result.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $result.StdOut | Should -Match 'SubfolderScanFailures='
        (Get-Content -LiteralPath $logPath -Raw) | Should -Match 'exported=3'
        (Get-Content -LiteralPath $logPath -Raw) | Should -Match 'itemReadFailures=0'
    }

    It 'runs the launcher sample path successfully' {
        $reportPath = Join-Path $TestDrive 'launcher-sample.html'
        $logPath = Join-Path $TestDrive 'launcher-sample.log'

        $result = & $script:invokePwshScriptCapture -FilePath $script:launcherPath -Arguments @('-NoGui', '-UseSampleData', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath $reportPath) | Should -BeTrue
        (Test-Path -LiteralPath $logPath) | Should -BeTrue
        $result.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $result.StdOut | Should -Match 'SubfolderScanFailures='
        (Get-Content -LiteralPath $logPath -Raw) | Should -Match 'exported=3'
        (Get-Content -LiteralPath $logPath -Raw) | Should -Match 'itemReadFailures=0'
    }

    It 'passes the permanent stop-process-tree regression script' {
        $outputDir = Join-Path $TestDrive 'regression-output'
        $result = & $script:invokePwshScriptCapture -FilePath $script:regressionScriptPath -Arguments @('-OutputDirectory', $outputDir)

        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'SourceContainsExitCodeGate\s+:\s+True'
        $result.StdOut | Should -Match 'SourceLauncherExitCode\s+:\s+0'
        $result.StdOut | Should -Match 'LogShowsExported3\s+:\s+True'
        $result.StdOut | Should -Match 'LogShowsZeroItemReadFailures\s+:\s+True'
    }

    It 'validates data contracts against schemas' {
        $validateScript = Join-Path $script:repoRoot 'scripts\Validate-DataContracts.ps1'
        $result = & $script:invokePwshScriptCapture -FilePath $validateScript
        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'CONTRACT_VALIDATION_PASSED'
    }

    It 'requires a delivered report file on the NoGui success path' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw
        $match = [regex]::Match($launcherText, '(?s)function Start-NoGuiMode \{(.+?)\r?\n\}')
        $match.Success | Should -BeTrue
        $body = $match.Groups[1].Value
        $body | Should -Match 'Test-Path -LiteralPath \$reportPath'
        $body | Should -Match 'report file was not found'
    }

    It 'builds debug and release executables' {
        $version = '1.0.29.1'
        $result = & $script:invokePwshScriptCapture -FilePath $script:buildScriptPath -Arguments @('-Version', $version)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath (Join-Path $script:buildDir 'PurviewTeamsPstToHtmlConverter_Debug.exe')) | Should -BeTrue
        (Test-Path -LiteralPath (Join-Path $script:buildDir 'PurviewTeamsPstToHtmlConverter.exe')) | Should -BeTrue
        $result.StdOut | Should -Match 'built 1.0.29.1'
    }
}
