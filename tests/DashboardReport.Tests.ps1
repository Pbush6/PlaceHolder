Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Conversion dashboard report' {
    BeforeAll {
        $script:corePath = Join-Path $PSScriptRoot '..\src\Convert-PurviewTeamsPstToHtml.ps1'
        $tokens = $null
        $errors = $null
        $script:coreAst = [Management.Automation.Language.Parser]::ParseFile(
            $script:corePath,
            [ref]$tokens,
            [ref]$errors)
        $errors | Should -BeNullOrEmpty

        $scriptSource = Get-Content -LiteralPath $script:corePath -Raw
        $scriptSource = [regex]::Replace($scriptSource, '\r?\nInvoke-ReportConversion\s*$', "`r`n")
        . ([scriptblock]::Create($scriptSource))

        function script:Reset-DashboardTestState {
            $script:Stats = [ordered]@{
                StartedAt               = Get-Date
                FoldersScanned          = 3
                ItemsAttempted          = 6
                ItemsExported           = 6
                ItemsSkipped            = 0
                TeamsItemsExported      = 2
                EmailItemsExported      = 2
                CalendarItemsExported   = 1
                ContactsItemsExported   = 1
                ItemReadFailures        = 0
                AttachmentReadFailures  = 0
                SubfolderScanFailures   = 0
            }
            $script:LogPath = Join-Path $TestDrive 'dashboard-test.log'
            $script:LogWriter = $null
            $script:LogWriters = @()
            $script:RunId = ''
            $script:TeamsOutputPath = $null
            $script:EmailOutputPath = $null
            $script:CalendarOutputPath = $null
            $script:ContactsOutputPath = $null
            $script:TeamsLogPath = $null
            $script:EmailLogPath = $null
            $script:CalendarLogPath = $null
            $script:ContactsLogPath = $null
        }

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
            }
        }
    }

    BeforeEach {
        Reset-DashboardTestState
    }

    It 'defines the dashboard writer, path helper, and email launch helper writer' {
        foreach ($name in @('Write-DashboardHtmlReport', 'Get-DashboardOutputPath', 'Write-EmailReportLaunchHelper', 'Get-DashboardReportEntries')) {
            $functionAst = $script:coreAst.Find({
                param($node)
                $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq $name
            }, $true)
            $functionAst | Should -Not -BeNullOrEmpty -Because "the dashboard phase requires $name"
        }
    }

    It 'derives the dashboard path from any typed report path' {
        Get-DashboardOutputPath -DisplayPath 'C:\out\LArtley Messages.html' |
            Should -Be 'C:\out\LArtley Messages_Dashboard.html'
        Get-DashboardOutputPath -DisplayPath 'C:\out\LArtley Messages_Calendar.html' |
            Should -Be 'C:\out\LArtley Messages_Dashboard.html'
        Get-DashboardOutputPath -DisplayPath 'C:\out\LArtley Messages_Dashboard.html' |
            Should -Be 'C:\out\LArtley Messages_Dashboard.html'
    }

    It 'writes one card per produced report and opens Email through the launch helper' {
        $baseOutputPath = Join-Path $TestDrive 'dashboard-all.html'
        $baseLogPath = Join-Path $TestDrive 'dashboard-all.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @(
            '-UseSampleData',
            '-NoPrompt',
            '-OutputPath', $baseOutputPath,
            '-LogPath', $baseLogPath
        )

        $dashboardPath = Join-Path $TestDrive 'dashboard-all_Dashboard.html'
        $helperPath = Join-Path $TestDrive 'Open-EmailReport.cmd'

        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'DashboardOutputPath=.*dashboard-all_Dashboard\.html'
        (Test-Path -LiteralPath $dashboardPath -PathType Leaf) | Should -BeTrue
        (Test-Path -LiteralPath $helperPath -PathType Leaf) | Should -BeTrue

        $dashboardHtml = Get-Content -LiteralPath $dashboardPath -Raw
        foreach ($reportKey in @('teams', 'email', 'calendar', 'contacts')) {
            $dashboardHtml | Should -Match ("data-report='{0}'" -f $reportKey)
        }
        $dashboardHtml | Should -Match 'dashboard-all_Teams\.html'
        $dashboardHtml | Should -Match 'dashboard-all_Calendar\.html'
        $dashboardHtml | Should -Match 'dashboard-all_Contacts\.html'
        $dashboardHtml | Should -Match 'Open-EmailReport\.cmd'
        $dashboardHtml | Should -Match 'dashboard-all_Email\.db'
        $dashboardHtml | Should -Match "class='dashboard-open'"
        ([regex]::Matches($dashboardHtml, "class='dashboard-open'[^>]*target='_blank'")).Count |
            Should -Be 4 -Because 'reports open in a new tab so the dashboard stays available'

        $helperText = Get-Content -LiteralPath $helperPath -Raw
        $helperText | Should -Match 'dashboard-all_Email\.db'
        $helperText | Should -Match 'EmailReviewViewer\.App\.exe'
    }

    It 'omits cards for reports that were not produced and skips the helper without Email' {
        $subsetDirectory = Join-Path $TestDrive 'subset'
        New-Item -ItemType Directory -Path $subsetDirectory -Force | Out-Null
        $baseOutputPath = Join-Path $subsetDirectory 'dashboard-subset.html'
        $baseLogPath = Join-Path $subsetDirectory 'dashboard-subset.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @(
            '-UseSampleData',
            '-NoPrompt',
            '-TeamsReport:$false',
            '-EmailReport:$false',
            '-CalendarReport:$true',
            '-ContactsReport:$false',
            '-OutputPath', $baseOutputPath,
            '-LogPath', $baseLogPath
        )

        $dashboardPath = Join-Path $subsetDirectory 'dashboard-subset_Dashboard.html'
        $helperPath = Join-Path $subsetDirectory 'Open-EmailReport.cmd'

        $result.ExitCode | Should -Be 0
        (Test-Path -LiteralPath $dashboardPath -PathType Leaf) | Should -BeTrue
        (Test-Path -LiteralPath $helperPath -PathType Leaf) | Should -BeFalse

        $dashboardHtml = Get-Content -LiteralPath $dashboardPath -Raw
        $dashboardHtml | Should -Match "data-report='calendar'"
        $dashboardHtml | Should -Not -Match "data-report='teams'"
        $dashboardHtml | Should -Not -Match "data-report='email'"
        $dashboardHtml | Should -Not -Match "data-report='contacts'"
    }

    It 'reports exported counts and read warnings from run statistics' {
        $script:TeamsOutputPath = Join-Path $TestDrive 'counts_Teams.html'
        $script:TeamsLogPath = Join-Path $TestDrive 'counts_Teams.log'
        $script:Stats.ItemReadFailures = 4
        $script:Stats.AttachmentReadFailures = 7
        $dashboardPath = Join-Path $TestDrive 'counts_Dashboard.html'
        $pstItem = [pscustomobject]@{ Name = 'Case.pst'; FullName = 'C:\Case.pst' }

        Write-DashboardHtmlReport -PstItem $pstItem -ReportPath $dashboardPath -Entries (Get-DashboardReportEntries)

        $html = Get-Content -LiteralPath $dashboardPath -Raw
        $html | Should -Match "data-report='teams'"
        $html | Should -Match 'data-item-count=.2.'
        $html | Should -Match 'Items: 4'
        $html | Should -Match 'Attachments: 7'
    }

    It 'HTML-encodes hostile PST and report names' {
        $hostileDirectory = Join-Path $TestDrive 'hostile'
        New-Item -ItemType Directory -Path $hostileDirectory -Force | Out-Null
        $script:CalendarOutputPath = Join-Path $hostileDirectory "<script>alert('x')</script>_Calendar.html"
        $script:CalendarLogPath = Join-Path $hostileDirectory 'safe_Calendar.log'
        $dashboardPath = Join-Path $hostileDirectory 'hostile_Dashboard.html'
        $pstItem = [pscustomobject]@{ Name = "<img src=x onerror=alert('pst')>"; FullName = 'C:\hostile.pst' }

        Write-DashboardHtmlReport -PstItem $pstItem -ReportPath $dashboardPath -Entries (Get-DashboardReportEntries)

        $html = Get-Content -LiteralPath $dashboardPath -Raw
        $html | Should -Not -Match '<img src=x'
        $html | Should -Not -Match "<script>alert\('x'\)"
        $html | Should -Match '&lt;img src=x'
    }
}
