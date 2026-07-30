Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Report launching' {
    BeforeAll {
        $script:launcherPath = Join-Path $PSScriptRoot '..\src\Start-PurviewTeamsPstToHtmlApp.ps1'
        $tokens = $null
        $errors = $null
        $script:launcherAst = [Management.Automation.Language.Parser]::ParseFile(
            $script:launcherPath,
            [ref]$tokens,
            [ref]$errors)
        $errors | Should -BeNullOrEmpty

        function Import-LauncherFunction([string]$Name) {
            $functionAst = $script:launcherAst.Find({
                param($node)
                $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq $Name
            }, $true)
            $functionAst | Should -Not -BeNullOrEmpty
            $definition = $functionAst.Extent.Text -replace
                ('^function\s+' + [regex]::Escape($Name)),
                ('function script:' + $Name)
            Invoke-Expression $definition
        }

        Import-LauncherFunction 'ConvertTo-NativeArgumentString'
        Import-LauncherFunction 'ConvertTo-ArgumentList'
        Import-LauncherFunction 'Split-PipeDelimitedStdoutLine'
        Import-LauncherFunction 'ConvertFrom-ResultLine'
        Import-LauncherFunction 'Get-WritePathsFromResultFields'
        Import-LauncherFunction 'Get-ConversionPathsFromResultLine'
        Import-LauncherFunction 'Get-ReportPathBaseName'
        Import-LauncherFunction 'Get-ReportOutputPaths'
        Import-LauncherFunction 'Get-FallbackLogPaths'
        Import-LauncherFunction 'Get-StdoutErrorDetail'
        Import-LauncherFunction 'Get-OutlookCleanupWarningFromLog'
        Import-LauncherFunction 'Assert-SelectedConversionArtifacts'
        Import-LauncherFunction 'Get-CurrentExecutablePath'
        Import-LauncherFunction 'Resolve-EmailViewerPath'
    }

    It 'resolves the viewer beside the packaged converter process' {
        $converterDirectory = Join-Path $TestDrive 'Delivered Folder'
        $viewerDirectory = Join-Path $converterDirectory 'EmailReviewViewer'
        New-Item -ItemType Directory -Path $viewerDirectory -Force | Out-Null
        $converterPath = Join-Path $converterDirectory 'PurviewTeamsPstToHtmlConverter.exe'
        $viewerPath = Join-Path $viewerDirectory 'EmailReviewViewer.App.exe'
        Set-Content -LiteralPath $converterPath -Value ''
        Set-Content -LiteralPath $viewerPath -Value ''

        Resolve-EmailViewerPath -ProcessPath $converterPath -ScriptDirectory '' -AppBaseDirectory '' |
            Should -Be ([IO.Path]::GetFullPath($viewerPath))
    }

    It 'resolves the viewer from a source launcher tree' {
        $sourceDirectory = Join-Path $TestDrive 'repo\src'
        $viewerDirectory = Join-Path $TestDrive 'repo\EmailReviewViewer\EmailReviewViewer.App\bin\Release\net8.0-windows'
        New-Item -ItemType Directory -Path $sourceDirectory, $viewerDirectory -Force | Out-Null
        $viewerPath = Join-Path $viewerDirectory 'EmailReviewViewer.App.exe'
        Set-Content -LiteralPath $viewerPath -Value ''

        Resolve-EmailViewerPath -ProcessPath 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' `
            -ScriptDirectory $sourceDirectory -AppBaseDirectory 'C:\Windows\System32\WindowsPowerShell\v1.0\' |
            Should -Be ([IO.Path]::GetFullPath($viewerPath))
    }

    It 'quotes database paths for Windows native process arguments' {
        $databasePath = 'C:\Review files\O''Brien (Final)\mail "Q4".db'

        ConvertTo-NativeArgumentString -Argument @($databasePath) |
            Should -Be '"C:\Review files\O''Brien (Final)\mail \"Q4\".db"'
    }

    It 'uses the Windows PowerShell compatible Arguments property' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw

        $launcherText | Should -Not -Match '\.ArgumentList\.Add'
        $launcherText | Should -Match '\$psi\.Arguments\s*=\s*ConvertTo-NativeArgumentString'
    }

    It 'forwards all four explicit report flags to the child process' {
        $arguments = ConvertTo-ArgumentList -CorePath 'C:\Temp\core.ps1' -PstPath '' `
            -OutputPath 'C:\Temp\report.html' -LogPath 'C:\Temp\report.log' `
            -DefaultConversationParticipants @() -TeamsReport:$false -EmailReport:$true `
            -CalendarReport:$false -ContactsReport:$true

        $arguments | Should -Contain '-TeamsReport:$false'
        $arguments | Should -Contain '-EmailReport:$true'
        $arguments | Should -Contain '-CalendarReport:$false'
        $arguments | Should -Contain '-ContactsReport:$true'
    }

    It 'accepts all four typed result paths and exported counts' {
        $line = 'CONVERSION_RESULT|OutputPath=C:\Reports\case.html|LogPath=C:\Reports\case.log|ItemsExported=6|TeamsOutputPath=C:\Reports\case_Teams.html|EmailOutputPath=C:\Reports\case_Email.db|CalendarOutputPath=C:\Reports\case_Calendar.html|ContactsOutputPath=C:\Reports\case_Contacts.html|TeamsLogPath=C:\Reports\case_Teams.log|EmailLogPath=C:\Reports\case_Email.log|CalendarLogPath=C:\Reports\case_Calendar.log|ContactsLogPath=C:\Reports\case_Contacts.log|TeamsItemsExported=2|EmailItemsExported=2|CalendarItemsExported=1|ContactsItemsExported=1'
        $parsed = ConvertFrom-ResultLine -ResultLine $line
        $paths = Get-WritePathsFromResultFields -Fields $parsed

        $parsed['CalendarOutputPath'] | Should -Be 'C:\Reports\case_Calendar.html'
        $parsed['ContactsOutputPath'] | Should -Be 'C:\Reports\case_Contacts.html'
        $parsed['CalendarLogPath'] | Should -Be 'C:\Reports\case_Calendar.log'
        $parsed['ContactsLogPath'] | Should -Be 'C:\Reports\case_Contacts.log'
        $parsed['CalendarItemsExported'] | Should -Be '1'
        $parsed['ContactsItemsExported'] | Should -Be '1'
        @($paths.ReportPaths).Count | Should -Be 4
        @($paths.LogPaths).Count | Should -Be 4
    }

    It 'reads NoGui cleanup warnings and fatal details from typed result logs' {
        $displayLog = Join-Path $TestDrive 'case.log'
        $teamsLog = Join-Path $TestDrive 'case_Teams.log'
        $contactsLog = Join-Path $TestDrive 'case_Contacts.log'
        Set-Content -LiteralPath $displayLog -Value 'display path is not a write target'
        Set-Content -LiteralPath $teamsLog -Value '2026-07-29T12:00:00 [WARN] PST is still attached to your Outlook profile: Sample PST'
        Set-Content -LiteralPath $contactsLog -Value '2026-07-29T12:00:01 [ERROR] Fatal error: typed-log-only detail'
        $line = "CONVERSION_RESULT|LogPath=$displayLog|TeamsLogPath=$teamsLog|ContactsLogPath=$contactsLog"

        $paths = Get-ConversionPathsFromResultLine -ResultLine $line

        Get-OutlookCleanupWarningFromLog -LogPath $paths.LogPaths |
            Should -Match 'PST is still attached'
        Get-StdoutErrorDetail -ErrorLine $null -LogPath $paths.LogPaths |
            Should -Match 'typed-log-only detail'
    }

    It 'resolves GUI diagnostics from typed result logs and typed fallbacks' {
        $functionAst = $script:launcherAst.Find({
            param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Get-GuiDiagnosticLogPaths'
        }, $true)
        $functionAst | Should -Not -BeNullOrEmpty
        $definition = $functionAst.Extent.Text -replace
            '^function\s+Get-GuiDiagnosticLogPaths',
            'function script:Get-GuiDiagnosticLogPaths'
        Invoke-Expression $definition

        $typedLine = 'CONVERSION_RESULT|LogPath=C:\Reports\case.log|TeamsLogPath=C:\Reports\case_Teams.log|ContactsLogPath=C:\Reports\case_Contacts.log'
        $typed = Get-GuiDiagnosticLogPaths -ResultLine $typedLine -DisplayLogPath 'C:\Reports\case.log' `
            -TeamsReport:$true -EmailReport:$false -CalendarReport:$false -ContactsReport:$true
        $fallback = Get-GuiDiagnosticLogPaths -ResultLine $null -DisplayLogPath 'C:\Reports\case.log' `
            -TeamsReport:$false -EmailReport:$false -CalendarReport:$true -ContactsReport:$true

        $typed | Should -Be @('C:\Reports\case_Teams.log', 'C:\Reports\case_Contacts.log')
        $typed | Should -Not -Contain 'C:\Reports\case.log'
        $fallback | Should -Be @('C:\Reports\case_Calendar.log', 'C:\Reports\case_Contacts.log')
    }

    It 'uses GUI diagnostic parity without changing the NoGui result-only path' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw
        $launcherText | Should -Match '(?s)if \(\$proc\.ExitCode -ne 0\).*Get-GuiDiagnosticLogPaths.*Get-StdoutErrorDetail[^\r\n]*-LogPath \$diagnosticLogPaths'
        $launcherText | Should -Match 'Get-OutlookCleanupWarningFromLog -LogPath \$script:lastLogPaths'
        $launcherText | Should -Not -Match 'Get-OutlookCleanupWarningFromLog -LogPath \(\$script:lastLogPaths \| Select-Object -First 1\)'
        $launcherText | Should -Match '(?s)function Start-NoGuiMode \{.*Get-ConversionPathsFromResultLine -ResultLine \$result\.ResultLine'
    }

    It 'wires four default-selected GUI report checkboxes and disables them while running' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw

        foreach ($name in @('teamsCheck', 'emailCheck', 'calendarCheck', 'contactsCheck')) {
            $launcherText | Should -Match ('\${0}\.Checked\s*=\s*\$true' -f $name)
            $launcherText | Should -Match ('\${0}\.Enabled\s*=\s*\$Enabled' -f $name)
            $launcherText | Should -Match ('\${0}\.Add_CheckedChanged\(' -f $name)
        }
        $launcherText | Should -Match 'At least one box must be checked \(Teams, Email, Calendar, and/or Contacts report\)\.'
    }

    It 'gates typed outputs before launching each selected report once' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw

        $launcherText | Should -Match '(?s)Get-ReportOutputPaths .*CalendarReport \$calendarCheck\.Checked .*ContactsReport \$contactsCheck\.Checked'
        $launcherText | Should -Match '\$logOutputPaths\s*=\s*Get-ReportOutputPaths'
        $launcherText | Should -Match '\$existingWritePaths'
        $launcherText | Should -Match 'selected report/log file\(s\)'
        $launcherText | Should -Match '(?s)if \(\$proc\.ExitCode -ne 0\).*Assert-SelectedConversionArtifacts.*foreach \(\$path in @\(\$script:lastReportPaths \| Select-Object -Unique\)\).*Open-GeneratedReport'
    }

    It 'rejects a selected typed output that is missing on disk' {
        $outputPath = Join-Path $TestDrive 'missing-calendar.html'
        $logPath = Join-Path $TestDrive 'calendar.log'
        Set-Content -LiteralPath $logPath -Value 'log'
        $fields = @{
            CalendarOutputPath = $outputPath
            CalendarLogPath = $logPath
        }

        {
            Assert-SelectedConversionArtifacts -Fields $fields `
                -TeamsReport:$false -EmailReport:$false -CalendarReport:$true -ContactsReport:$false
        } | Should -Throw '*selected report output file was not found*CalendarOutputPath*'
    }

    It 'rejects a selected typed log that is missing on disk' {
        $outputPath = Join-Path $TestDrive 'contacts.html'
        $logPath = Join-Path $TestDrive 'missing-contacts.log'
        Set-Content -LiteralPath $outputPath -Value '<html></html>'
        $fields = @{
            ContactsOutputPath = $outputPath
            ContactsLogPath = $logPath
        }

        {
            Assert-SelectedConversionArtifacts -Fields $fields `
                -TeamsReport:$false -EmailReport:$false -CalendarReport:$false -ContactsReport:$true
        } | Should -Throw '*selected report log file was not found*ContactsLogPath*'
    }

    It 'returns every selected typed output and log when all are present' {
        $fields = @{}
        foreach ($report in @('Teams', 'Email', 'Calendar', 'Contacts')) {
            $outputExtension = if ($report -eq 'Email') { '.db' } else { '.html' }
            $outputPath = Join-Path $TestDrive ($report + $outputExtension)
            $logPath = Join-Path $TestDrive ($report + '.log')
            Set-Content -LiteralPath $outputPath -Value 'output'
            Set-Content -LiteralPath $logPath -Value 'log'
            $fields[$report + 'OutputPath'] = $outputPath
            $fields[$report + 'LogPath'] = $logPath
        }

        $artifacts = Assert-SelectedConversionArtifacts -Fields $fields `
            -TeamsReport:$true -EmailReport:$true -CalendarReport:$true -ContactsReport:$true

        @($artifacts.ReportPaths).Count | Should -Be 4
        @($artifacts.LogPaths).Count | Should -Be 4
    }

    It 'rejects a partial result when all four reports were selected' {
        $fields = @{}
        foreach ($report in @('Teams', 'Email', 'Calendar', 'Contacts')) {
            $outputExtension = if ($report -eq 'Email') { '.db' } else { '.html' }
            $outputPath = Join-Path $TestDrive ('partial-' + $report + $outputExtension)
            $logPath = Join-Path $TestDrive ('partial-' + $report + '.log')
            Set-Content -LiteralPath $outputPath -Value 'output'
            Set-Content -LiteralPath $logPath -Value 'log'
            $fields[$report + 'OutputPath'] = $outputPath
            $fields[$report + 'LogPath'] = $logPath
        }
        $fields.Remove('ContactsOutputPath')

        {
            Assert-SelectedConversionArtifacts -Fields $fields `
                -TeamsReport:$true -EmailReport:$true -CalendarReport:$true -ContactsReport:$true
        } | Should -Throw '*missing selected report output path field ContactsOutputPath*'
    }
}
