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
        $script:viewerExe = Join-Path $script:repoRoot 'EmailReviewViewer\EmailReviewViewer.App\bin\Release\net8.0-windows\EmailReviewViewer.App.exe'
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

    It 'passes all four report flags through the launcher with default-true public parameters' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw
        $launcherText | Should -Match '\[bool\]\$TeamsReport = \$true'
        $launcherText | Should -Match '\[bool\]\$EmailReport = \$true'
        $launcherText | Should -Match '\[bool\]\$CalendarReport = \$true'
        $launcherText | Should -Match '\[bool\]\$ContactsReport = \$true'
        $launcherText | Should -Match '(?s)ConvertTo-ArgumentList .* -TeamsReport:\$TeamsReport -EmailReport:\$EmailReport -CalendarReport:\$CalendarReport -ContactsReport:\$ContactsReport'
        $launcherText | Should -Match '(?s)Invoke-EmbeddedConversion .* -TeamsReport:\$TeamsReport -EmailReport:\$EmailReport -CalendarReport:\$CalendarReport -ContactsReport:\$ContactsReport'
        $launcherText | Should -Match '(?s)Start-EmbeddedConversionProcess .* -TeamsReport:\$true -EmailReport:\$true -CalendarReport:\$true -ContactsReport:\$true'
    }

    It 'runs the core sample path successfully with all four reports by default' {
        $reportPath = Join-Path $TestDrive 'all-four.html'
        $logPath = Join-Path $TestDrive 'all-four.log'
        $teamsLogPath = Join-Path $TestDrive 'all-four_Teams.log'
        $emailLogPath = Join-Path $TestDrive 'all-four_Email.log'
        $calendarLogPath = Join-Path $TestDrive 'all-four_Calendar.log'
        $contactsLogPath = Join-Path $TestDrive 'all-four_Contacts.log'
        $teamsReportPath = Join-Path $TestDrive 'all-four_Teams.html'
        $emailReportPath = Join-Path $TestDrive 'all-four_Email.db'
        $calendarReportPath = Join-Path $TestDrive 'all-four_Calendar.html'
        $contactsReportPath = Join-Path $TestDrive 'all-four_Contacts.html'

        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @('-UseSampleData', '-NoPrompt', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        foreach ($path in @($teamsReportPath, $emailReportPath, $calendarReportPath, $contactsReportPath, $teamsLogPath, $emailLogPath, $calendarLogPath, $contactsLogPath)) {
            (Test-Path -LiteralPath $path) | Should -BeTrue -Because "selected typed output must exist: $path"
        }
        $result.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $result.StdOut | Should -Match 'ItemsExported=6'
        $result.StdOut | Should -Match 'TeamsItemsExported=2'
        $result.StdOut | Should -Match 'EmailItemsExported=2'
        $result.StdOut | Should -Match 'CalendarItemsExported=1'
        $result.StdOut | Should -Match 'ContactsItemsExported=1'
        $result.StdOut | Should -Match 'TeamsOutputPath=.*_Teams\.html'
        $result.StdOut | Should -Match 'EmailOutputPath=.*_Email\.db'
        $result.StdOut | Should -Match 'CalendarOutputPath=.*_Calendar\.html'
        $result.StdOut | Should -Match 'ContactsOutputPath=.*_Contacts\.html'
        $result.StdOut | Should -Match 'TeamsLogPath=.*_Teams\.log'
        $result.StdOut | Should -Match 'EmailLogPath=.*_Email\.log'
        $result.StdOut | Should -Match 'CalendarLogPath=.*_Calendar\.log'
        $result.StdOut | Should -Match 'ContactsLogPath=.*_Contacts\.log'
        $result.StdOut | Should -Match 'SubfolderScanFailures='
        (Get-Content -LiteralPath $teamsLogPath -Raw) | Should -Match 'exported=6'
        (Get-Content -LiteralPath $emailLogPath -Raw) | Should -Match 'exported=6'
        (Get-Content -LiteralPath $teamsLogPath -Raw) | Should -Match 'itemReadFailures=0'
        (Get-Content -LiteralPath $emailLogPath -Raw) | Should -Match 'itemReadFailures=0'
        (Test-Path -LiteralPath (Join-Path $TestDrive 'all-four_Email.html')) | Should -BeFalse
        $benchmark = (& $script:viewerExe --benchmark $emailReportPath | Out-String)
        $benchmark | Should -Match '"bodyLoadedInListQuery": false'
        $benchmark | Should -Match '"totalMatches": 1'
    }

    It 'preserves legacy core two-report calls when new flags are omitted' {
        $reportPath = Join-Path $TestDrive 'legacy-core.html'
        $logPath = Join-Path $TestDrive 'legacy-core.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @(
            '-UseSampleData', '-NoPrompt', '-TeamsReport:$true', '-EmailReport:$true',
            '-OutputPath', $reportPath, '-LogPath', $logPath
        )

        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'ItemsExported=4'
        (Test-Path -LiteralPath (Join-Path $TestDrive 'legacy-core_Teams.html')) | Should -BeTrue
        (Test-Path -LiteralPath (Join-Path $TestDrive 'legacy-core_Email.db')) | Should -BeTrue
        (Test-Path -LiteralPath (Join-Path $TestDrive 'legacy-core_Calendar.html')) | Should -BeFalse
        (Test-Path -LiteralPath (Join-Path $TestDrive 'legacy-core_Contacts.html')) | Should -BeFalse
    }

    It 'preserves legacy launcher Teams-only calls when new flags are omitted' {
        $reportPath = Join-Path $TestDrive 'legacy-launcher.html'
        $logPath = Join-Path $TestDrive 'legacy-launcher.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:launcherPath -Arguments @(
            '-NoGui', '-UseSampleData', '-TeamsReport:$true', '-EmailReport:$false',
            '-OutputPath', $reportPath, '-LogPath', $logPath
        )

        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'ItemsExported=2'
        (Test-Path -LiteralPath (Join-Path $TestDrive 'legacy-launcher_Teams.html')) | Should -BeTrue
        (Test-Path -LiteralPath (Join-Path $TestDrive 'legacy-launcher_Email.db')) | Should -BeFalse
        (Test-Path -LiteralPath (Join-Path $TestDrive 'legacy-launcher_Calendar.html')) | Should -BeFalse
        (Test-Path -LiteralPath (Join-Path $TestDrive 'legacy-launcher_Contacts.html')) | Should -BeFalse
    }

    It 'runs the core sample path successfully with Teams only' {
        $reportPath = Join-Path $TestDrive 'teams-only.html'
        $logPath = Join-Path $TestDrive 'teams-only.log'
        $teamsReportPath = Join-Path $TestDrive 'teams-only_Teams.html'
        $teamsLogPath = Join-Path $TestDrive 'teams-only_Teams.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @('-UseSampleData', '-NoPrompt', '-TeamsReport:$true', '-EmailReport:$false', '-CalendarReport:$false', '-ContactsReport:$false', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath $teamsReportPath) | Should -BeTrue
        (Test-Path -LiteralPath $teamsLogPath) | Should -BeTrue
        $result.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $result.StdOut | Should -Match 'ItemsExported=2'
        $result.StdOut | Should -Match 'TeamsItemsExported=2'
        $result.StdOut | Should -Match 'TeamsOutputPath=.*_Teams\.html'
        $result.StdOut | Should -Match 'TeamsLogPath=.*_Teams\.log'
        (Test-Path -LiteralPath (Join-Path $TestDrive 'teams-only_Email.db')) | Should -BeFalse
        (Test-Path -LiteralPath (Join-Path $TestDrive 'teams-only_Email.log')) | Should -BeFalse
    }

    It 'runs the core sample path successfully with Email only' {
        $reportPath = Join-Path $TestDrive 'email-only.html'
        $logPath = Join-Path $TestDrive 'email-only.log'
        $emailReportPath = Join-Path $TestDrive 'email-only_Email.db'
        $emailLogPath = Join-Path $TestDrive 'email-only_Email.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @('-UseSampleData', '-NoPrompt', '-TeamsReport:$false', '-EmailReport:$true', '-CalendarReport:$false', '-ContactsReport:$false', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath $emailReportPath) | Should -BeTrue
        (Test-Path -LiteralPath $emailLogPath) | Should -BeTrue
        $result.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $result.StdOut | Should -Match 'ItemsExported=2'
        $result.StdOut | Should -Match 'EmailItemsExported=2'
        $result.StdOut | Should -Match 'EmailOutputPath=.*_Email\.db'
        $result.StdOut | Should -Match 'EmailLogPath=.*_Email\.log'
        (Test-Path -LiteralPath (Join-Path $TestDrive 'email-only_Teams.html')) | Should -BeFalse
        (Test-Path -LiteralPath (Join-Path $TestDrive 'email-only_Teams.log')) | Should -BeFalse
    }

    It 'runs the core sample path successfully with Calendar only' {
        $reportPath = Join-Path $TestDrive 'calendar-only.html'
        $logPath = Join-Path $TestDrive 'calendar-only.log'
        $calendarReportPath = Join-Path $TestDrive 'calendar-only_Calendar.html'
        $calendarLogPath = Join-Path $TestDrive 'calendar-only_Calendar.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @('-UseSampleData', '-NoPrompt', '-TeamsReport:$false', '-EmailReport:$false', '-CalendarReport:$true', '-ContactsReport:$false', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath $calendarReportPath) | Should -BeTrue
        (Test-Path -LiteralPath $calendarLogPath) | Should -BeTrue
        $result.StdOut | Should -Match 'ItemsExported=1'
        $result.StdOut | Should -Match 'CalendarItemsExported=1'
        $result.StdOut | Should -Match 'CalendarOutputPath=.*_Calendar\.html'
        $result.StdOut | Should -Match 'CalendarLogPath=.*_Calendar\.log'
        (Test-Path -LiteralPath (Join-Path $TestDrive 'calendar-only_Contacts.html')) | Should -BeFalse
    }

    It 'runs the core sample path successfully with Contacts only' {
        $reportPath = Join-Path $TestDrive 'contacts-only.html'
        $logPath = Join-Path $TestDrive 'contacts-only.log'
        $contactsReportPath = Join-Path $TestDrive 'contacts-only_Contacts.html'
        $contactsLogPath = Join-Path $TestDrive 'contacts-only_Contacts.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @('-UseSampleData', '-NoPrompt', '-TeamsReport:$false', '-EmailReport:$false', '-CalendarReport:$false', '-ContactsReport:$true', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath $contactsReportPath) | Should -BeTrue
        (Test-Path -LiteralPath $contactsLogPath) | Should -BeTrue
        $result.StdOut | Should -Match 'ItemsExported=1'
        $result.StdOut | Should -Match 'ContactsItemsExported=1'
        $result.StdOut | Should -Match 'ContactsOutputPath=.*_Contacts\.html'
        $result.StdOut | Should -Match 'ContactsLogPath=.*_Contacts\.log'
        (Test-Path -LiteralPath (Join-Path $TestDrive 'contacts-only_Calendar.html')) | Should -BeFalse
    }

    It 'runs the launcher sample path successfully with all four reports by default' {
        $reportPath = Join-Path $TestDrive 'launcher-sample.html'
        $logPath = Join-Path $TestDrive 'launcher-sample.log'
        $teamsReportPath = Join-Path $TestDrive 'launcher-sample_Teams.html'
        $emailReportPath = Join-Path $TestDrive 'launcher-sample_Email.db'
        $calendarReportPath = Join-Path $TestDrive 'launcher-sample_Calendar.html'
        $contactsReportPath = Join-Path $TestDrive 'launcher-sample_Contacts.html'
        $teamsLogPath = Join-Path $TestDrive 'launcher-sample_Teams.log'
        $emailLogPath = Join-Path $TestDrive 'launcher-sample_Email.log'
        $calendarLogPath = Join-Path $TestDrive 'launcher-sample_Calendar.log'
        $contactsLogPath = Join-Path $TestDrive 'launcher-sample_Contacts.log'

        $result = & $script:invokePwshScriptCapture -FilePath $script:launcherPath -Arguments @('-NoGui', '-UseSampleData', '-OutputPath', $reportPath, '-LogPath', $logPath)

        $result.ExitCode | Should -Be 0
        foreach ($path in @($teamsReportPath, $emailReportPath, $calendarReportPath, $contactsReportPath, $teamsLogPath, $emailLogPath, $calendarLogPath, $contactsLogPath)) {
            (Test-Path -LiteralPath $path) | Should -BeTrue -Because "launcher selected typed output must exist: $path"
        }
        $result.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $result.StdOut | Should -Match 'ItemsExported=6'
        $result.StdOut | Should -Match 'TeamsItemsExported=2'
        $result.StdOut | Should -Match 'EmailItemsExported=2'
        $result.StdOut | Should -Match 'CalendarItemsExported=1'
        $result.StdOut | Should -Match 'ContactsItemsExported=1'
        $result.StdOut | Should -Match 'SubfolderScanFailures='
        (Get-Content -LiteralPath $teamsLogPath -Raw) | Should -Match 'exported=6'
        (Get-Content -LiteralPath $emailLogPath -Raw) | Should -Match 'exported=6'
        (Get-Content -LiteralPath $teamsLogPath -Raw) | Should -Match 'itemReadFailures=0'
        (Get-Content -LiteralPath $emailLogPath -Raw) | Should -Match 'itemReadFailures=0'
    }

    It 'runs launcher Calendar-only and Contacts-only sample paths' {
        foreach ($case in @(
            @{ Name = 'launcher-calendar-only'; Calendar = '$true'; Contacts = '$false'; Suffix = 'Calendar'; CountField = 'CalendarItemsExported=1' },
            @{ Name = 'launcher-contacts-only'; Calendar = '$false'; Contacts = '$true'; Suffix = 'Contacts'; CountField = 'ContactsItemsExported=1' }
        )) {
            $reportPath = Join-Path $TestDrive ($case.Name + '.html')
            $logPath = Join-Path $TestDrive ($case.Name + '.log')
            $expectedReportPath = Join-Path $TestDrive ($case.Name + '_' + $case.Suffix + '.html')
            $expectedLogPath = Join-Path $TestDrive ($case.Name + '_' + $case.Suffix + '.log')
            $result = & $script:invokePwshScriptCapture -FilePath $script:launcherPath -Arguments @(
                '-NoGui', '-UseSampleData', '-TeamsReport:$false', '-EmailReport:$false',
                ('-CalendarReport:' + $case.Calendar), ('-ContactsReport:' + $case.Contacts),
                '-OutputPath', $reportPath, '-LogPath', $logPath
            )

            $result.ExitCode | Should -Be 0
            (Test-Path -LiteralPath $expectedReportPath -PathType Leaf) | Should -BeTrue
            (Test-Path -LiteralPath $expectedLogPath -PathType Leaf) | Should -BeTrue
            $result.StdOut | Should -Match $case.CountField
        }
    }

    It 'rejects all four report flags false in core and launcher CLI modes' {
        $reportPath = Join-Path $TestDrive 'none-selected.html'
        $logPath = Join-Path $TestDrive 'none-selected.log'
        $flags = @('-UseSampleData', '-TeamsReport:$false', '-EmailReport:$false', '-CalendarReport:$false', '-ContactsReport:$false', '-OutputPath', $reportPath, '-LogPath', $logPath)
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments (@('-NoPrompt') + $flags)

        $result.ExitCode | Should -Not -Be 0
        $result.StdOut | Should -Match 'CONVERSION_ERROR'
        $result.StdOut | Should -Match 'At least one report type must be selected'

        $launcherResult = & $script:invokePwshScriptCapture -FilePath $script:launcherPath -Arguments (@('-NoGui') + $flags)
        $launcherResult.ExitCode | Should -Not -Be 0
        $launcherResult.StdOut | Should -Match 'At least one report type must be selected'
    }

    It 'passes the permanent stop-process-tree regression script' {
        $outputDir = Join-Path $TestDrive 'regression-output'
        $result = & $script:invokePwshScriptCapture -FilePath $script:regressionScriptPath -Arguments @('-OutputDirectory', $outputDir)

        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'SourceContainsExitCodeGate\s+:\s+True'
        $result.StdOut | Should -Match 'SourceLauncherExitCode\s+:\s+0'
        $result.StdOut | Should -Match 'LogShowsExported6\s+:\s+True'
        $result.StdOut | Should -Match 'LogShowsZeroItemReadFailures\s+:\s+True'
    }

    It 'validates data contracts against schemas' {
        $validateScript = Join-Path $script:repoRoot 'scripts\Validate-DataContracts.ps1'
        $result = & $script:invokePwshScriptCapture -FilePath $validateScript
        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'CONTRACT_VALIDATION_PASSED'
    }

    It 'requires selected typed report and log artifacts on the NoGui success path' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw
        $match = [regex]::Match($launcherText, '(?s)function Start-NoGuiMode \{(.+?)\r?\n\}')
        $match.Success | Should -BeTrue
        $body = $match.Groups[1].Value
        $body | Should -Match 'Assert-SelectedConversionArtifacts'
        $body | Should -Match 'lastReportPaths'
        $body | Should -Match 'lastLogPaths'
        $body | Should -Not -Match 'Get-FallbackWritePaths'
    }

    It 'GUI success gate captures RESULT during progress drain and never requires dual DisplayPath' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw
        $launcherText | Should -Match 'capturedResultLine'
        $launcherText | Should -Match 'Assert-SelectedConversionArtifacts'
        $launcherText | Should -Match "TeamsOutputPath', 'EmailOutputPath', 'CalendarOutputPath', 'ContactsOutputPath"
        # Progress drain must capture RESULT/ERROR before ConvertFrom-ProgressStdoutLine discards them
        $launcherText | Should -Match '(?s)while \(\$queue\.TryDequeue\(\[ref\]\$item\)\).*CONVERSION_RESULT\|.*ConvertFrom-ProgressStdoutLine'
    }

    It 'detach finally-block never lets COM verify-throw abort a successful conversion' {
        $coreText = Get-Content -LiteralPath $script:corePath -Raw
        $coreText | Should -Match '(?s)function Test-PstStoreAttached \{.*?catch \{.*?return \$true'
        $coreText | Should -Match 'PST detach cleanup failed unexpectedly'
        $coreText | Should -Match '(?s)finally \{.*?PST detach cleanup failed unexpectedly.*?Summary: folders='
    }

    It 'builds debug and release executables' -Tag 'Build' {
        $version = '1.3.0.0'
        $result = & $script:invokePwshScriptCapture -FilePath $script:buildScriptPath -Arguments @('-Version', $version)

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath (Join-Path $script:buildDir 'PurviewTeamsPstToHtmlConverter_Debug.exe')) | Should -BeTrue
        (Test-Path -LiteralPath (Join-Path $script:buildDir 'PurviewTeamsPstToHtmlConverter.exe')) | Should -BeTrue
        $result.StdOut | Should -Match 'built 1.3.0.0'
    }

    It 'release no-console executable completes NoGui sample mode without interaction' -Tag 'Build' {
        $releaseExe = Join-Path $script:buildDir 'PurviewTeamsPstToHtmlConverter.exe'
        $outputPath = Join-Path $TestDrive 'release-nogui.html'
        $logPath = Join-Path $TestDrive 'release-nogui.log'
        $psi = [Diagnostics.ProcessStartInfo]::new()
        $psi.FileName = $releaseExe
        $psi.WorkingDirectory = $script:repoRoot
        $psi.UseShellExecute = $false
        $psi.Arguments = "-NoGui -UseSampleData -TeamsReport:`$true -EmailReport:`$false -CalendarReport:`$true -ContactsReport:`$true -OutputPath `"$outputPath`" -LogPath `"$logPath`""
        $process = [Diagnostics.Process]::new()
        $process.StartInfo = $psi
        $completed = $false
        try {
            $process.Start() | Should -BeTrue
            $completed = $process.WaitForExit(60000)
            $completed | Should -BeTrue -Because 'NoGui mode must not wait for a PS2EXE GUI output dialog'
            $process.ExitCode | Should -Be 0
        }
        finally {
            if (-not $process.HasExited) {
                & taskkill.exe /PID $process.Id /T /F | Out-Null
                [void]$process.WaitForExit(10000)
            }
            $process.Dispose()
        }

        foreach ($name in @(
            'release-nogui_Teams.html',
            'release-nogui_Calendar.html',
            'release-nogui_Contacts.html',
            'release-nogui_Teams.log',
            'release-nogui_Calendar.log',
            'release-nogui_Contacts.log'
        )) {
            (Test-Path -LiteralPath (Join-Path $TestDrive $name) -PathType Leaf) |
                Should -BeTrue -Because "NoGui output must exist: $name"
        }
    }
}
