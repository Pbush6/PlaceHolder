Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Item report bucket' {
    BeforeAll {
        . (Join-Path $PSScriptRoot '..\src\ReportClassification.ps1')
    }

    It 'Teams folder override wins even for IPM.Note' {
        Get-ItemReportBucket -FolderPath 'X\TeamsMessagesData' -MessageClass 'IPM.Note' | Should -Be 'Teams'
    }

    It 'TeamsMeetings folder keeps Teams precedence over calendar classes' {
        Get-ItemReportBucket -FolderPath 'X\SkypeSpacesData\TeamsMeetings' -MessageClass 'IPM.Schedule.Meeting.Request' | Should -Be 'Teams'
    }

    It 'Inbox IPM.Note is Email' {
        Get-ItemReportBucket -FolderPath 'X\Inbox' -MessageClass 'IPM.Note' | Should -Be 'Email'
    }

    It 'IPM.Note.SMIME is Email' {
        Get-ItemReportBucket -FolderPath 'X\Inbox' -MessageClass 'IPM.Note.SMIME' | Should -Be 'Email'
    }

    It 'meeting request is Calendar' {
        Get-ItemReportBucket -FolderPath 'X\Inbox' -MessageClass 'IPM.Schedule.Meeting.Request' | Should -Be 'Calendar'
    }

    It 'appointment is Calendar' {
        Get-ItemReportBucket -FolderPath 'X\Calendar' -MessageClass 'IPM.Appointment' | Should -Be 'Calendar'
    }

    It 'contact is Contacts' {
        Get-ItemReportBucket -FolderPath 'X\Contacts' -MessageClass 'IPM.Contact' | Should -Be 'Contacts'
    }

    It 'distribution list is Contacts' {
        Get-ItemReportBucket -FolderPath 'X\Contacts' -MessageClass 'IPM.DistList' | Should -Be 'Contacts'
    }

    It 'unknown class is Skip' {
        Get-ItemReportBucket -FolderPath 'X\Misc' -MessageClass 'IPM.StickyNote' | Should -Be 'Skip'
    }
}
