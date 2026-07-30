Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Dot-source only the helper region after Task 1 implements it in a tiny shared file:
# Create: src/ReportPathNaming.ps1 (dot-sourced by launcher + core + tests)

Describe 'Report path naming' {
    BeforeAll {
        . (Join-Path $PSScriptRoot '..\src\ReportPathNaming.ps1')
    }
    It 'strips report suffixes from base' {
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Teams.html' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Email.html' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Email.db' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Calendar.html' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Contacts.html' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages.html' | Should -Be 'LArtley Messages'
    }

    It 'keeps Get-ReportPathBaseName textually synchronized across all sources' {
        $paths = @(
            (Join-Path $PSScriptRoot '..\src\ReportPathNaming.ps1'),
            (Join-Path $PSScriptRoot '..\src\Convert-PurviewTeamsPstToHtml.ps1'),
            (Join-Path $PSScriptRoot '..\src\Start-PurviewTeamsPstToHtmlApp.ps1')
        )
        $definitions = foreach ($path in $paths) {
            $tokens = $null
            $errors = $null
            $ast = [Management.Automation.Language.Parser]::ParseFile(
                [IO.Path]::GetFullPath($path),
                [ref]$tokens,
                [ref]$errors)
            $errors | Should -BeNullOrEmpty
            $functionAst = $ast.Find({
                param($node)
                $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq 'Get-ReportPathBaseName'
            }, $true)
            $functionAst | Should -Not -BeNullOrEmpty
            $functionAst.Extent.Text -replace "`r`n", "`n"
        }

        $definitions[1] | Should -BeExactly $definitions[0]
        $definitions[2] | Should -BeExactly $definitions[0]
    }

    It 'multiple selected -> display base and write typed sibling outputs' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.html' -TeamsReport $true -EmailReport $true -CalendarReport $true -ContactsReport $true
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages.html'
        $r.TeamsPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.EmailPath | Should -Be 'C:\out\LArtley Messages_Email.db'
        $r.CalendarPath | Should -Be 'C:\out\LArtley Messages_Calendar.html'
        $r.ContactsPath | Should -Be 'C:\out\LArtley Messages_Contacts.html'
    }

    It 'Teams only -> display and write _Teams' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.html' -TeamsReport $true -EmailReport $false -CalendarReport $false -ContactsReport $false
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.TeamsPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.EmailPath | Should -BeNullOrEmpty
        $r.CalendarPath | Should -BeNullOrEmpty
        $r.ContactsPath | Should -BeNullOrEmpty
    }

    It 'Email only -> display and write _Email' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages_Teams.html' -TeamsReport $false -EmailReport $true -CalendarReport $false -ContactsReport $false
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Email.db'
        $r.EmailPath | Should -Be 'C:\out\LArtley Messages_Email.db'
        $r.TeamsPath | Should -BeNullOrEmpty
        $r.CalendarPath | Should -BeNullOrEmpty
        $r.ContactsPath | Should -BeNullOrEmpty
    }

    It 'keeps new path flags optional and false for legacy direct helper callers' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\Legacy.html' -TeamsReport $true -EmailReport $false

        $r.DisplayPath | Should -Be 'C:\out\Legacy_Teams.html'
        $r.TeamsPath | Should -Be 'C:\out\Legacy_Teams.html'
        $r.EmailPath | Should -BeNullOrEmpty
        $r.CalendarPath | Should -BeNullOrEmpty
        $r.ContactsPath | Should -BeNullOrEmpty
    }

    It 'Calendar only -> display and write _Calendar' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages_Email.db' -TeamsReport $false -EmailReport $false -CalendarReport $true -ContactsReport $false
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Calendar.html'
        $r.CalendarPath | Should -Be 'C:\out\LArtley Messages_Calendar.html'
        $r.TeamsPath | Should -BeNullOrEmpty
        $r.EmailPath | Should -BeNullOrEmpty
        $r.ContactsPath | Should -BeNullOrEmpty
    }

    It 'Contacts only -> display and write _Contacts' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages_Calendar.html' -TeamsReport $false -EmailReport $false -CalendarReport $false -ContactsReport $true
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Contacts.html'
        $r.ContactsPath | Should -Be 'C:\out\LArtley Messages_Contacts.html'
        $r.TeamsPath | Should -BeNullOrEmpty
        $r.EmailPath | Should -BeNullOrEmpty
        $r.CalendarPath | Should -BeNullOrEmpty
    }

    It 'none selected throws' {
        { Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.html' -TeamsReport $false -EmailReport $false -CalendarReport $false -ContactsReport $false } |
            Should -Throw 'At least one report type must be selected.'
    }

    It 'log paths keep their original extension' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.txt' -TeamsReport $true -EmailReport $false -CalendarReport $true -ContactsReport $false
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages.txt'
        $r.TeamsPath | Should -Be 'C:\out\LArtley Messages_Teams.txt'
        $r.CalendarPath | Should -Be 'C:\out\LArtley Messages_Calendar.txt'
        $r.EmailPath | Should -BeNullOrEmpty
        $r.ContactsPath | Should -BeNullOrEmpty
    }

    It 'typed result fields prefer all report-specific write targets' {
        $fields = @{
            OutputPath = 'C:\out\LArtley Messages.html'
            LogPath = 'C:\out\LArtley Messages.log'
            TeamsOutputPath = 'C:\out\LArtley Messages_Teams.html'
            EmailOutputPath = 'C:\out\LArtley Messages_Email.db'
            CalendarOutputPath = 'C:\out\LArtley Messages_Calendar.html'
            ContactsOutputPath = 'C:\out\LArtley Messages_Contacts.html'
            TeamsLogPath = 'C:\out\LArtley Messages_Teams.log'
            EmailLogPath = 'C:\out\LArtley Messages_Email.log'
            CalendarLogPath = 'C:\out\LArtley Messages_Calendar.log'
            ContactsLogPath = 'C:\out\LArtley Messages_Contacts.log'
        }
        $w = Get-WritePathsFromResultFields -Fields $fields
        $w.ReportPaths | Should -Be @(
            'C:\out\LArtley Messages_Teams.html',
            'C:\out\LArtley Messages_Email.db',
            'C:\out\LArtley Messages_Calendar.html',
            'C:\out\LArtley Messages_Contacts.html'
        )
        $w.LogPaths | Should -Be @(
            'C:\out\LArtley Messages_Teams.log',
            'C:\out\LArtley Messages_Email.log',
            'C:\out\LArtley Messages_Calendar.log',
            'C:\out\LArtley Messages_Contacts.log'
        )
        $w.ReportPaths | Should -Not -Contain 'C:\out\LArtley Messages.html'
        $w.LogPaths | Should -Not -Contain 'C:\out\LArtley Messages.log'
    }

    It 'single-mode result fields fall back to OutputPath when typed paths absent' {
        $fields = @{
            OutputPath = 'C:\out\LArtley Messages_Teams.html'
            LogPath = 'C:\out\LArtley Messages_Teams.log'
        }
        $w = Get-WritePathsFromResultFields -Fields $fields
        $w.ReportPaths | Should -Be @('C:\out\LArtley Messages_Teams.html')
        $w.LogPaths | Should -Be @('C:\out\LArtley Messages_Teams.log')
    }
}
