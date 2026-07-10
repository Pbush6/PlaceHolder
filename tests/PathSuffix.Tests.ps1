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
        Get-ReportPathBaseName 'C:\out\LArtley Messages.html' | Should -Be 'LArtley Messages'
    }

    It 'both selected -> display base, write two sibling files' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.html' -TeamsReport $true -EmailReport $true
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages.html'
        $r.TeamsPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.EmailPath | Should -Be 'C:\out\LArtley Messages_Email.html'
    }

    It 'Teams only -> display and write _Teams' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages.html' -TeamsReport $true -EmailReport $false
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.TeamsPath | Should -Be 'C:\out\LArtley Messages_Teams.html'
        $r.EmailPath | Should -BeNullOrEmpty
    }

    It 'Email only -> display and write _Email' {
        $r = Get-ReportOutputPaths -DisplayPath 'C:\out\LArtley Messages_Teams.html' -TeamsReport $false -EmailReport $true
        $r.DisplayPath | Should -Be 'C:\out\LArtley Messages_Email.html'
        $r.EmailPath | Should -Be 'C:\out\LArtley Messages_Email.html'
        $r.TeamsPath | Should -BeNullOrEmpty
    }
}
