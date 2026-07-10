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

    It 'passes the Teams and Email report flags through the launcher' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw
        $launcherText | Should -Match '\[bool\]\$TeamsReport = \$true'
        $launcherText | Should -Match '\[bool\]\$EmailReport = \$true'
        $launcherText | Should -Match '(?s)ConvertTo-ArgumentList .* -TeamsReport:\$TeamsReport -EmailReport:\$EmailReport'
        $launcherText | Should -Match '(?s)Invoke-EmbeddedConversion .* -TeamsReport:\$TeamsReport -EmailReport:\$EmailReport'
        $launcherText | Should -Match '(?s)Start-EmbeddedConversionProcess .* -TeamsReport:\$teamsCheck\.Checked -EmailReport:\$emailCheck\.Checked'
    }

    It 'runs the core sample path successfully with both reports' {
        $reportPath = Join-Path $TestDrive 'dual.html'
        $logPath = Join-Path $TestDrive 'dual.log'
        $teamsLogPath = Join-Path $TestDrive 'dual_Teams.log'
        $emailLogPath = Join-Path $TestDrive 'dual_Email.log'
        $teamsReportPath = Join-Path $TestDrive 'dual_Teams.html'
        $emailReportPath = Join-Path $TestDrive 'dual_Email.html'

        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @('-UseSampleData', '-NoPrompt', '-TeamsReport:$true', '-EmailReport:$true', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath $teamsReportPath) | Should -BeTrue
        (Test-Path -LiteralPath $emailReportPath) | Should -BeTrue
        (Test-Path -LiteralPath $teamsLogPath) | Should -BeTrue
        (Test-Path -LiteralPath $emailLogPath) | Should -BeTrue
        $result.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $result.StdOut | Should -Match 'TeamsOutputPath=.*_Teams\.html'
        $result.StdOut | Should -Match 'EmailOutputPath=.*_Email\.html'
        $result.StdOut | Should -Match 'SubfolderScanFailures='
        (Get-Content -LiteralPath $teamsLogPath -Raw) | Should -Match 'exported=4'
        (Get-Content -LiteralPath $emailLogPath -Raw) | Should -Match 'exported=4'
        (Get-Content -LiteralPath $teamsLogPath -Raw) | Should -Match 'itemReadFailures=0'
        (Get-Content -LiteralPath $emailLogPath -Raw) | Should -Match 'itemReadFailures=0'
        (Get-Content -LiteralPath $emailReportPath -Raw) | Should -Match 'Purview Email'
        (Get-Content -LiteralPath $emailReportPath -Raw) | Should -Match 'data-folder'
        (Get-Content -LiteralPath $emailReportPath -Raw) | Should -Match 'folder-filter'
        (Get-Content -LiteralPath $emailReportPath -Raw) | Should -Match 'participantMatchMode'
        (Get-Content -LiteralPath $emailReportPath -Raw) | Should -Match 'conversation-toolbar'
        ([regex]::Matches((Get-Content -LiteralPath $emailReportPath -Raw), "<section class='conversation'")).Count | Should -Be 1
    }

    It 'runs the launcher sample path successfully' {
        $reportPath = Join-Path $TestDrive 'launcher-sample.html'
        $logPath = Join-Path $TestDrive 'launcher-sample.log'
        $teamsReportPath = Join-Path $TestDrive 'launcher-sample_Teams.html'
        $emailReportPath = Join-Path $TestDrive 'launcher-sample_Email.html'
        $teamsLogPath = Join-Path $TestDrive 'launcher-sample_Teams.log'
        $emailLogPath = Join-Path $TestDrive 'launcher-sample_Email.log'

        $result = & $script:invokePwshScriptCapture -FilePath $script:launcherPath -Arguments @('-NoGui', '-UseSampleData', '-TeamsReport:$true', '-EmailReport:$true', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath $teamsReportPath) | Should -BeTrue
        (Test-Path -LiteralPath $emailReportPath) | Should -BeTrue
        (Test-Path -LiteralPath $teamsLogPath) | Should -BeTrue
        (Test-Path -LiteralPath $emailLogPath) | Should -BeTrue
        $result.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $result.StdOut | Should -Match 'SubfolderScanFailures='
        (Get-Content -LiteralPath $teamsLogPath -Raw) | Should -Match 'exported=4'
        (Get-Content -LiteralPath $emailLogPath -Raw) | Should -Match 'exported=4'
        (Get-Content -LiteralPath $teamsLogPath -Raw) | Should -Match 'itemReadFailures=0'
        (Get-Content -LiteralPath $emailLogPath -Raw) | Should -Match 'itemReadFailures=0'
    }

    It 'passes the permanent stop-process-tree regression script' {
        $outputDir = Join-Path $TestDrive 'regression-output'
        $result = & $script:invokePwshScriptCapture -FilePath $script:regressionScriptPath -Arguments @('-OutputDirectory', $outputDir)

        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'SourceContainsExitCodeGate\s+:\s+True'
        $result.StdOut | Should -Match 'SourceLauncherExitCode\s+:\s+0'
        $result.StdOut | Should -Match 'LogShowsExported4\s+:\s+True'
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
        $body | Should -Match 'Get-ConversionPathsFromResultLine'
        $body | Should -Match 'lastReportPaths'
        $body | Should -Match 'one or more report files'
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
