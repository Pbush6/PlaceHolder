Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Calendar and contact record extraction' {
    BeforeAll {
        $script:corePath = Join-Path $PSScriptRoot '..\src\Convert-PurviewTeamsPstToHtml.ps1'
        $tokens = $null
        $errors = $null
        $script:coreAst = [Management.Automation.Language.Parser]::ParseFile(
            $script:corePath,
            [ref]$tokens,
            [ref]$errors)
        $errors | Should -BeNullOrEmpty

        function Import-CoreFunction([string]$Name) {
            $functionAst = $script:coreAst.Find({
                param($node)
                $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq $Name
            }, $true)
            $functionAst | Should -Not -BeNullOrEmpty -Because "$Name should exist in the core script"
            $definition = $functionAst.Extent.Text -replace
                ('^function\s+' + [regex]::Escape($Name)),
                ('function script:' + $Name)
            Invoke-Expression $definition
        }

        foreach ($name in @(
            'Close-ComObjectSafe',
            'Get-PropSafe',
            'ConvertTo-NormalizedPersonName',
            'Split-RecipientName',
            'Get-ParticipantName',
            'Get-ParticipantKey',
            'Test-MissingDate',
            'Get-AttachmentMetadata',
            'Get-AttachmentSummaryHtml',
            'Test-IsTeamsMessagesFolder',
            'Test-IsEmailMessageClass',
            'Test-IsCalendarMessageClass',
            'Test-IsContactsMessageClass',
            'Get-ItemReportBucket',
            'Get-MessageRecord',
            'Get-RecurrenceSummary',
            'Get-CalendarRecord',
            'Get-DistributionListMembers',
            'Get-ContactRecord',
            'Get-SampleRecord',
            'Read-OutlookFolder'
        )) {
            Import-CoreFunction $name
        }

        function script:Write-ReportLog {
            param([string]$Message, [string]$Level = 'INFO')
        }

        function script:Write-ConversionProgress {
            param([string]$FolderPath)
        }

        function script:Invoke-OutlookComOperation {
            param(
                [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
                [Parameter(Mandatory = $true)][string]$Operation
            )
            & $ScriptBlock
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

        function script:New-OutlookCollection {
            param([object[]]$Values)
            $collection = [pscustomobject]@{
                Values = @($Values)
                Count = @($Values).Count
            }
            Add-Member -InputObject $collection -MemberType ScriptMethod -Name Item -Value {
                param($Index)
                $this.Values[$Index - 1]
            }
            return $collection
        }

        function script:New-OutlookFolder {
            param(
                [object[]]$Items = @(),
                [object[]]$Folders = @()
            )
            [pscustomobject]@{
                Items = New-OutlookCollection -Values $Items
                Folders = New-OutlookCollection -Values $Folders
            }
        }

        function script:Reset-TestStats {
            $script:Stats = [ordered]@{
                StartedAt               = Get-Date
                FoldersScanned          = 0
                ItemsAttempted          = 0
                ItemsExported           = 0
                ItemsSkipped            = 0
                TeamsItemsExported      = 0
                EmailItemsExported      = 0
                CalendarItemsExported   = 0
                ContactsItemsExported   = 0
                ItemReadFailures        = 0
                AttachmentReadFailures  = 0
                SubfolderScanFailures   = 0
            }
            $script:TeamsReport = $true
            $script:EmailReport = $true
            $script:CalendarReport = $true
            $script:ContactsReport = $true
            $script:LogEvery = 1000
        }
    }

    BeforeEach {
        Reset-TestStats
    }

    It 'builds a stable recurring calendar record shape' {
        $recurrence = [pscustomobject]@{
            RecurrenceType = 1
            Interval = 2
            PatternStartDate = [datetime]'2024-03-01T00:00:00'
            PatternEndDate = [datetime]'2024-06-01T00:00:00'
            NoEndDate = $false
            Occurrences = 8
        }
        $item = [pscustomobject]@{
            Start = [datetime]'2024-03-01T09:00:00'
            End = [datetime]'2024-03-01T10:00:00'
            Subject = 'Weekly staff meeting'
            MessageClass = 'IPM.Appointment'
            Body = 'Agenda here'
            Categories = 'Blue Category'
            Location = 'Conf Room A'
            Organizer = 'Torey Page'
            RequiredAttendees = 'Linda Artley; Sam Reed'
            OptionalAttendees = 'Ava Stone'
            AllDayEvent = $false
            IsRecurring = $true
            Sensitivity = 2
            CreationTime = [datetime]'2024-02-20T08:30:00'
            LastModificationTime = [datetime]'2024-02-22T11:45:00'
            EntryID = 'calendar-1'
            RecurrencePattern = $recurrence
        }
        Add-Member -InputObject $item -MemberType ScriptMethod -Name GetRecurrencePattern -Value {
            $this.RecurrencePattern
        }

        $record = Get-CalendarRecord -Item $item -FolderPath 'Mailbox\Calendar'

        $record.SortTime | Should -Be ([datetime]'2024-03-01T09:00:00')
        $record.StartTime | Should -Be ([datetime]'2024-03-01T09:00:00')
        $record.EndTime | Should -Be ([datetime]'2024-03-01T10:00:00')
        $record.Subject | Should -Be 'Weekly staff meeting'
        $record.ItemType | Should -Be 'Appointment'
        $record.AllDayEvent | Should -BeFalse
        $record.Location | Should -Be 'Conf Room A'
        $record.Organizer | Should -Be 'Torey Page'
        $record.RequiredAttendees | Should -Be 'Linda Artley; Sam Reed'
        $record.OptionalAttendees | Should -Be 'Ava Stone'
        $record.IsRecurring | Should -BeTrue
        $record.RecurrenceSummary | Should -Match 'Weekly'
        $record.RecurrenceSummary | Should -Match 'every 2'
        $record.RecurrenceSummary | Should -Match '2024-03-01'
        $record.RecurrenceSummary | Should -Match '2024-06-01'
        $record.RecurrenceSummary | Should -Match '8 occurrence'
        $record.Categories | Should -Be 'Blue Category'
        $record.Sensitivity | Should -Be 2
        $record.FolderPath | Should -Be 'Mailbox\Calendar'
        $record.BodyText | Should -Be 'Agenda here'
        $record.Notes | Should -Be 'Agenda here'
        $record.MessageClass | Should -Be 'IPM.Appointment'
        $record.EntryId | Should -Be 'calendar-1'
        $record.CreationTime | Should -Be ([datetime]'2024-02-20T08:30:00')
        $record.LastModificationTime | Should -Be ([datetime]'2024-02-22T11:45:00')
        @($record.Attachments).Count | Should -Be 0
    }

    It 'tolerates missing calendar optionals and throwing recurrence access' {
        $item = [pscustomobject]@{
            Start = [datetime]'2024-04-10T00:00:00'
            MessageClass = 'IPM.Schedule.Meeting.Request'
            IsRecurring = $true
        }
        Add-Member -InputObject $item -MemberType ScriptMethod -Name GetRecurrencePattern -Value {
            throw 'recurrence unavailable'
        }

        $record = Get-CalendarRecord -Item $item -FolderPath 'Mailbox\Inbox'

        $record.Subject | Should -Be ''
        $record.EndTime | Should -BeNullOrEmpty
        $record.Location | Should -Be ''
        $record.RequiredAttendees | Should -Be ''
        $record.OptionalAttendees | Should -Be ''
        $record.IsRecurring | Should -BeTrue
        $record.RecurrenceSummary | Should -Be ''
        $record.Sensitivity | Should -BeNullOrEmpty
        $record.BodyText | Should -Be ''
        $record.Notes | Should -Be ''
        @($record.Attachments).Count | Should -Be 0
    }

    It 'builds a stable contact record shape including distribution list members' {
        $members = @(
            [pscustomobject]@{ Name = 'Linda Artley' },
            [pscustomobject]@{ Name = 'Sam Reed' }
        )
        $item = [pscustomobject]@{
            MessageClass = 'IPM.DistList'
            DLName = 'Curriculum Team'
            FirstName = ''
            MiddleName = ''
            LastName = ''
            CompanyName = 'Perfection Learning'
            JobTitle = 'Department List'
            Department = 'Curriculum'
            Email1Address = 'curriculum@example.com'
            Email2Address = 'curriculum-alt@example.com'
            Email3Address = ''
            BusinessTelephoneNumber = '555-1000'
            HomeTelephoneNumber = '555-2000'
            MobileTelephoneNumber = '555-3000'
            OtherTelephoneNumber = '555-4000'
            BusinessAddress = '100 Main St'
            HomeAddress = '200 Oak Ave'
            OtherAddress = '300 Pine Rd'
            WebPage = 'https://example.com'
            Birthday = [datetime]'1990-05-01T00:00:00'
            Anniversary = [datetime]'2015-06-15T00:00:00'
            Categories = 'Vendors'
            Body = 'Shared list'
            EntryID = 'contact-1'
            CreationTime = [datetime]'2024-01-15T12:00:00'
            LastModificationTime = [datetime]'2024-01-20T15:30:00'
            MemberCount = 2
            MemberObjects = $members
        }
        Add-Member -InputObject $item -MemberType ScriptMethod -Name GetMember -Value {
            param($Index)
            $this.MemberObjects[$Index - 1]
        }

        $record = Get-ContactRecord -Item $item -FolderPath 'Mailbox\Contacts'

        $record.DisplayName | Should -Be 'Curriculum Team'
        $record.FullName | Should -Be 'Curriculum Team'
        $record.FirstName | Should -Be ''
        $record.MiddleName | Should -Be ''
        $record.LastName | Should -Be ''
        $record.CompanyName | Should -Be 'Perfection Learning'
        $record.JobTitle | Should -Be 'Department List'
        $record.Department | Should -Be 'Curriculum'
        $record.Email1 | Should -Be 'curriculum@example.com'
        $record.Email2 | Should -Be 'curriculum-alt@example.com'
        $record.Email3 | Should -Be ''
        $record.BusinessPhone | Should -Be '555-1000'
        $record.HomePhone | Should -Be '555-2000'
        $record.MobilePhone | Should -Be '555-3000'
        $record.OtherPhone | Should -Be '555-4000'
        $record.BusinessAddress | Should -Be '100 Main St'
        $record.HomeAddress | Should -Be '200 Oak Ave'
        $record.OtherAddress | Should -Be '300 Pine Rd'
        $record.WebPage | Should -Be 'https://example.com'
        $record.Birthday | Should -Be ([datetime]'1990-05-01T00:00:00')
        $record.Anniversary | Should -Be ([datetime]'2015-06-15T00:00:00')
        $record.Categories | Should -Be 'Vendors'
        $record.DistributionListMembers | Should -Be @('Linda Artley', 'Sam Reed')
        $record.BodyText | Should -Be 'Shared list'
        $record.Notes | Should -Be 'Shared list'
        $record.MessageClass | Should -Be 'IPM.DistList'
        $record.EntryId | Should -Be 'contact-1'
        $record.CreationTime | Should -Be ([datetime]'2024-01-15T12:00:00')
        $record.LastModificationTime | Should -Be ([datetime]'2024-01-20T15:30:00')
        @($record.Attachments).Count | Should -Be 0
    }

    It 'tolerates inaccessible distribution members and missing contact optionals' {
        $item = [pscustomobject]@{
            MessageClass = 'IPM.Contact'
            FirstName = 'Linda'
            LastName = 'Artley'
            MemberCount = 1
        }
        Add-Member -InputObject $item -MemberType ScriptMethod -Name GetMember -Value {
            throw 'member lookup failed'
        }

        $record = Get-ContactRecord -Item $item -FolderPath 'Mailbox\Contacts'

        $record.DisplayName | Should -Be 'Linda Artley'
        $record.FirstName | Should -Be 'Linda'
        $record.LastName | Should -Be 'Artley'
        $record.Email1 | Should -Be ''
        $record.BusinessAddress | Should -Be ''
        $record.DistributionListMembers | Should -Be @()
        $record.BodyText | Should -Be ''
        $record.Notes | Should -Be ''
        @($record.Attachments).Count | Should -Be 0
    }

    It 'routes calendar and contacts items without subject or body through one-pass bucketing' {
        $calendarItem = [pscustomobject]@{
            MessageClass = 'IPM.Appointment'
            Start = [datetime]'2024-04-01T13:00:00'
            End = [datetime]'2024-04-01T14:00:00'
            Subject = ''
            Body = ''
            EntryID = 'calendar-no-body'
        }
        $contactItem = [pscustomobject]@{
            MessageClass = 'IPM.Contact'
            FullName = 'No Subject Contact'
            Subject = ''
            Body = ''
            EntryID = 'contact-no-body'
        }
        $folder = New-OutlookFolder -Items @($calendarItem, $contactItem)
        $teamsRecords = New-Object 'System.Collections.Generic.List[object]'
        $emailRecords = New-Object 'System.Collections.Generic.List[object]'
        $calendarRecords = New-Object 'System.Collections.Generic.List[object]'
        $contactRecords = New-Object 'System.Collections.Generic.List[object]'

        Read-OutlookFolder -Folder $folder -FolderPath 'Mailbox\Root' -TeamsRecords $teamsRecords -EmailRecords $emailRecords -CalendarRecords $calendarRecords -ContactsRecords $contactRecords

        $calendarRecords.Count | Should -Be 1
        $contactRecords.Count | Should -Be 1
        $calendarRecords[0].MessageClass | Should -Be 'IPM.Appointment'
        $contactRecords[0].DisplayName | Should -Be 'No Subject Contact'
        $script:Stats.CalendarItemsExported | Should -Be 1
        $script:Stats.ContactsItemsExported | Should -Be 1
        $script:Stats.ItemsExported | Should -Be 2
    }

    It 'extends sample records with representative calendar and contacts shapes while keeping teams and email counts unchanged' {
        $sample = Get-SampleRecord

        @($sample.Teams).Count | Should -Be 2
        @($sample.Email).Count | Should -Be 2
        @($sample.Calendar).Count | Should -Be 1
        @($sample.Contacts).Count | Should -Be 1
        $sample.Calendar[0].IsRecurring | Should -BeTrue
        $sample.Calendar[0].RecurrenceSummary | Should -Match 'Weekly'
        $sample.Contacts[0].DisplayName | Should -Not -BeNullOrEmpty
        $sample.Contacts[0].DistributionListMembers.Count | Should -BeGreaterThan 0
    }

    It 'passes calendar-only and contacts-only flags through full conversion path resolution' {
        $calendarReportPath = Join-Path $TestDrive 'calendar-only.html'
        $calendarLogPath = Join-Path $TestDrive 'calendar-only.log'
        $calendarResult = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @(
            '-UseSampleData',
            '-NoPrompt',
            '-TeamsReport:$false',
            '-EmailReport:$false',
            '-CalendarReport:$true',
            '-ContactsReport:$false',
            '-OutputPath', $calendarReportPath,
            '-LogPath', $calendarLogPath
        )

        $calendarResult.ExitCode | Should -Be 0
        $calendarResult.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $calendarResult.StdOut | Should -Match 'OutputPath=.*calendar-only_Calendar\.html'
        $calendarResult.StdOut | Should -Match 'LogPath=.*calendar-only_Calendar\.log'
        $calendarResult.StdOut | Should -Match 'CalendarOutputPath=.*calendar-only_Calendar\.html'
        $calendarResult.StdOut | Should -Match 'CalendarLogPath=.*calendar-only_Calendar\.log'
        $calendarResult.StdOut | Should -Match 'CalendarItemsExported=1'
        $calendarResult.StdOut | Should -Match 'ContactsItemsExported=0'

        $contactsReportPath = Join-Path $TestDrive 'contacts-only.html'
        $contactsLogPath = Join-Path $TestDrive 'contacts-only.log'
        $contactsResult = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @(
            '-UseSampleData',
            '-NoPrompt',
            '-TeamsReport:$false',
            '-EmailReport:$false',
            '-CalendarReport:$false',
            '-ContactsReport:$true',
            '-OutputPath', $contactsReportPath,
            '-LogPath', $contactsLogPath
        )

        $contactsResult.ExitCode | Should -Be 0
        $contactsResult.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $contactsResult.StdOut | Should -Match 'OutputPath=.*contacts-only_Contacts\.html'
        $contactsResult.StdOut | Should -Match 'LogPath=.*contacts-only_Contacts\.log'
        $contactsResult.StdOut | Should -Match 'ContactsOutputPath=.*contacts-only_Contacts\.html'
        $contactsResult.StdOut | Should -Match 'ContactsLogPath=.*contacts-only_Contacts\.log'
        $contactsResult.StdOut | Should -Match 'CalendarItemsExported=0'
        $contactsResult.StdOut | Should -Match 'ContactsItemsExported=1'
    }
}
