Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Dot-source only the helper region after Task 1 implements it in a tiny shared file:
# Create: src/ReportPathNaming.ps1 (dot-sourced by launcher + core + tests)

Describe 'Report path naming' {
    BeforeAll {
        . (Join-Path $PSScriptRoot '..\src\ReportPathNaming.ps1')
    }
    It 'strips Teams and Email suffixes from base' {
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Teams.html' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Email.html' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages_Email.db' | Should -Be 'LArtley Messages'
        Get-ReportPathBaseName 'C:\out\LArtley Messages.html' | Should -Be 'LArtley Messages'
    }

    It 'both selected -> display base, write two sibling files' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.html' -TeamsReport $true -EmailReport $true
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages.html'
        $r.TeamsPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.EmailPath | Should -Be 'C:\out\LArtley Messages_Email.db'
    }

    It 'Teams only -> display and write _Teams' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.html' -TeamsReport $true -EmailReport $false
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.TeamsPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.EmailPath | Should -BeNullOrEmpty
    }

    It 'Email only -> display and write _Email' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages_Teams.html' -TeamsReport $false -EmailReport $true
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Email.db'
        $r.EmailPath | Should -Be 'C:\out\LArtley Messages_Email.db'
        $r.TeamsPath | Should -BeNullOrEmpty
    }

    It 'dual-mode result fields prefer typed write targets over DisplayPath OutputPath' {
        $fields = @{
            OutputPath = 'C:\out\LArtley Messages.html'
            LogPath = 'C:\out\LArtley Messages.log'
            TeamsOutputPath = 'C:\out\LArtley Messages_Teams.html'
            EmailOutputPath = 'C:\out\LArtley Messages_Email.db'
            TeamsLogPath = 'C:\out\LArtley Messages_Teams.log'
            EmailLogPath = 'C:\out\LArtley Messages_Email.log'
        }
        $w = Get-WritePathsFromResultFields -Fields $fields
        $w.ReportPaths | Should -Be @('C:\out\LArtley Messages_Teams.html', 'C:\out\LArtley Messages_Email.db')
        $w.LogPaths | Should -Be @('C:\out\LArtley Messages_Teams.log', 'C:\out\LArtley Messages_Email.log')
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
