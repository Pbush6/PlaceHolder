Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Item report bucket' {
    BeforeAll {
        . (Join-Path $PSScriptRoot '..\src\ReportClassification.ps1')
    }

    It 'Teams folder wins even for IPM.Note' {
        Get-ItemReportBucket -FolderPath 'X\TeamsMessagesData' -MessageClass 'IPM.Note' | Should -Be 'Teams'
    }

    It 'TeamsMeetings is Teams' {
        Get-ItemReportBucket -FolderPath 'X\SkypeSpacesData\TeamsMeetings' -MessageClass 'IPM.AppointmentSnapshot.SkypeTeams.Meeting' | Should -Be 'Teams'
    }

    It 'Inbox IPM.Note is Email' {
        Get-ItemReportBucket -FolderPath 'X\Inbox' -MessageClass 'IPM.Note' | Should -Be 'Email'
    }

    It 'IPM.Note.SMIME is Email' {
        Get-ItemReportBucket -FolderPath 'X\Inbox' -MessageClass 'IPM.Note.SMIME' | Should -Be 'Email'
    }

    It 'meeting request is Skip' {
        Get-ItemReportBucket -FolderPath 'X\Inbox' -MessageClass 'IPM.Schedule.Meeting.Request' | Should -Be 'Skip'
    }

    It 'contact is Skip' {
        Get-ItemReportBucket -FolderPath 'X\Contacts' -MessageClass 'IPM.Contact' | Should -Be 'Skip'
    }
}
