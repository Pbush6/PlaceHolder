Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Calendar and contacts HTML reports' {
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

        function script:Reset-WriterTestState {
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
            $script:LogPath = Join-Path $TestDrive 'writer-test.log'
            $script:LogWriter = $null
            $script:LogWriters = @()
            $script:RunId = ''
            $script:DefaultConversationParticipants = @()
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
        Reset-WriterTestState
    }

    It 'defines dedicated calendar and contacts HTML report writers' {
        $calendarWriter = $script:coreAst.Find({
            param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Write-CalendarHtmlReport'
        }, $true)
        $contactsWriter = $script:coreAst.Find({
            param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Write-ContactsHtmlReport'
        }, $true)

        $calendarWriter | Should -Not -BeNullOrEmpty -Because 'Task 3 requires a dedicated Calendar HTML writer'
        $contactsWriter | Should -Not -BeNullOrEmpty -Because 'Task 3 requires a dedicated Contacts HTML writer'
    }

    It 'writes sample calendar and contacts reports with dataset filters and typed outputs' {
        $baseOutputPath = Join-Path $TestDrive 'calendar-contacts.html'
        $baseLogPath = Join-Path $TestDrive 'calendar-contacts.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @(
            '-UseSampleData',
            '-NoPrompt',
            '-TeamsReport:$false',
            '-EmailReport:$false',
            '-CalendarReport:$true',
            '-ContactsReport:$true',
            '-OutputPath', $baseOutputPath,
            '-LogPath', $baseLogPath
        )

        $calendarOutputPath = Join-Path $TestDrive 'calendar-contacts_Calendar.html'
        $contactsOutputPath = Join-Path $TestDrive 'calendar-contacts_Contacts.html'

        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Match 'CONVERSION_RESULT\|'
        $result.StdOut | Should -Match 'CalendarOutputPath=.*calendar-contacts_Calendar\.html'
        $result.StdOut | Should -Match 'ContactsOutputPath=.*calendar-contacts_Contacts\.html'
        $result.StdOut | Should -Match 'CalendarItemsExported=1'
        $result.StdOut | Should -Match 'ContactsItemsExported=1'
        (Test-Path -LiteralPath $calendarOutputPath -PathType Leaf) | Should -BeTrue
        (Test-Path -LiteralPath $contactsOutputPath -PathType Leaf) | Should -BeTrue

        $calendarHtml = Get-Content -LiteralPath $calendarOutputPath -Raw
        $contactsHtml = Get-Content -LiteralPath $contactsOutputPath -Raw

        $calendarHtml | Should -Match "id='calendarSearch'"
        $calendarHtml | Should -Match "id='calendarFromDateFilter'"
        $calendarHtml | Should -Match "id='calendarToDateFilter'"
        $calendarHtml | Should -Match "id='calendarFolderFilter'"
        $calendarHtml | Should -Match "id='calendarTypeFilter'"
        $calendarHtml | Should -Match "id='calendarAllDayFilter'"
        $calendarHtml | Should -Match "id='calendarRecurringFilter'"
        $calendarHtml | Should -Match "id='calendarClearFiltersBtn'"
        $calendarHtml | Should -Match "id='calendarVisibleCount'"
        $calendarHtml | Should -Match "data-search='"
        $calendarHtml | Should -Match "data-start-date='"
        $calendarHtml | Should -Match "data-end-date='"
        $calendarHtml | Should -Match "data-folder='"
        $calendarHtml | Should -Match "data-item-type='"
        $calendarHtml | Should -Match "data-all-day='"
        $calendarHtml | Should -Match "data-recurring='"
        $calendarHtml | Should -Match '\.dataset\.search'
        $calendarHtml | Should -Match '\.dataset\.startDate'
        $calendarHtml | Should -Match '\.dataset\.endDate'
        $calendarHtml | Should -Match '\.dataset\.folder'
        $calendarHtml | Should -Match '\.dataset\.itemType'
        $calendarHtml | Should -Match '\.dataset\.allDay'
        $calendarHtml | Should -Match '\.dataset\.recurring'
        $calendarHtml | Should -Match 'requestAnimationFrame|setTimeout'
        $calendarHtml | Should -Match 'content-visibility:\s*auto'
        $calendarHtml | Should -Match 'contain-intrinsic-size'
        $calendarHtml | Should -Match 'Subject'
        $calendarHtml | Should -Match 'Item type'
        $calendarHtml | Should -Match 'Required attendees'
        $calendarHtml | Should -Match 'Optional attendees'
        $calendarHtml | Should -Match 'Recurrence summary'
        $calendarHtml | Should -Match 'Attachment metadata'
        $calendarHtml | Should -Match 'Message class'
        $calendarHtml | Should -Match 'Entry ID'
        $calendarHtml | Should -Not -Match 'textContent\s*\|\|'
        $calendarHtml | Should -Not -Match '_searchText'
        $calendarHtml | Should -Not -Match "id='resizeHandle'"
        $calendarLayoutOverride = [regex]::Match($calendarHtml, '(?s)@media\s*\(min-width:\s*901px\)\s*\{\s*\.review-layout\s*\{\s*grid-template-columns:\s*minmax\([^;]+\)\s+minmax\(0,\s*1fr\)[^}]*\}\s*\}')
        $calendarLayoutOverride.Success | Should -BeTrue
        $calendarLayoutOverride.Value | Should -Not -Match '10px'

        $contactsHtml | Should -Match "id='contactsSearch'"
        $contactsHtml | Should -Match "id='contactsFolderFilter'"
        $contactsHtml | Should -Match "id='contactsCategoryFilter'"
        $contactsHtml | Should -Match "id='contactsClearFiltersBtn'"
        $contactsHtml | Should -Match "id='contactsVisibleCount'"
        $contactsHtml | Should -Match "data-search='"
        $contactsHtml | Should -Match "data-folder='"
        $contactsHtml | Should -Match "data-categories='"
        $contactsHtml | Should -Match '\.dataset\.search'
        $contactsHtml | Should -Match '\.dataset\.folder'
        $contactsHtml | Should -Match '\.dataset\.categories'
        $contactsHtml | Should -Match 'requestAnimationFrame|setTimeout'
        $contactsHtml | Should -Match 'content-visibility:\s*auto'
        $contactsHtml | Should -Match 'contain-intrinsic-size'
        $contactsHtml | Should -Match 'Organization'
        $contactsHtml | Should -Match 'Job title'
        $contactsHtml | Should -Match 'Department'
        $contactsHtml | Should -Match 'Website'
        $contactsHtml | Should -Match 'Birthday'
        $contactsHtml | Should -Match 'Anniversary'
        $contactsHtml | Should -Match 'Distribution list members'
        $contactsHtml | Should -Match 'Attachment metadata'
        $contactsHtml | Should -Match 'Message class'
        $contactsHtml | Should -Match 'Entry ID'
        $contactsHtml | Should -Not -Match 'textContent\s*\|\|'
        $contactsHtml | Should -Not -Match '_searchText'
        $contactsHtml | Should -Not -Match "id='resizeHandle'"
        $contactsLayoutOverride = [regex]::Match($contactsHtml, '(?s)@media\s*\(min-width:\s*901px\)\s*\{\s*\.review-layout\s*\{\s*grid-template-columns:\s*minmax\([^;]+\)\s+minmax\(0,\s*1fr\)[^}]*\}\s*\}')
        $contactsLayoutOverride.Success | Should -BeTrue
        $contactsLayoutOverride.Value | Should -Not -Match '10px'

        $teamsHeader = $script:coreAst.Find({
            param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Write-ReportHeader'
        }, $true)
        $teamsHeader.Extent.Text | Should -Match "id='resizeHandle'" -Because 'Teams column resizing remains supported'
    }

    It 'HTML-encodes hostile calendar and contacts values in markup and data attributes' {
        $calendarPath = Join-Path $TestDrive 'hostile-calendar.html'
        $contactsPath = Join-Path $TestDrive 'hostile-contacts.html'
        $pstItem = [pscustomobject]@{
            Name = 'Hostile <PST>.pst'
        }

        $calendarRecord = [pscustomobject]@{
            SortTime = [datetime]'2024-04-10T09:30:00'
            StartTime = [datetime]'2024-04-10T09:30:00'
            EndTime = [datetime]'2024-04-10T11:00:00'
            Subject = 'Quarterly <script>alert(1)</script> review'
            ItemType = 'Meeting "Series"'
            AllDayEvent = $false
            Location = 'Lab "A" & <Room>'
            Organizer = 'Pat <script>'
            RequiredAttendees = 'Alpha <beta>; O''Brien'
            OptionalAttendees = 'Guest "One"'
            IsRecurring = $true
            RecurrenceSummary = 'Every <2> weeks & "Fridays"'
            Categories = 'Blue & <Red>'
            Sensitivity = 3
            FolderPath = 'Mailbox\Calendar <Main>'
            BodyText = 'Body <b>unsafe</b> & more'
            Notes = 'Body <b>unsafe</b> & more'
            MessageClass = 'IPM.Appointment'
            EntryId = 'entry-<calendar>'
            CreationTime = [datetime]'2024-04-01T08:00:00'
            LastModificationTime = [datetime]'2024-04-09T17:15:00'
            Attachments = @(
                [pscustomobject]@{
                    FileName = 'agenda<script>.ics'
                    DisplayName = 'Agenda "Q2"'
                    Size = 2048
                }
            )
        }

        $contactRecord = [pscustomobject]@{
            DisplayName = 'Alex <img src=x onerror=alert(2)>'
            FullName = 'Alex <img src=x onerror=alert(2)>'
            FirstName = 'Alex'
            MiddleName = '"Middle"'
            LastName = 'O''Brien <script>'
            CompanyName = 'Perfection <Learning>'
            JobTitle = 'Director & "Lead"'
            Department = 'R&D <Blue>'
            Email1 = 'alex@example.com'
            Email2 = 'alt+<x>@example.com'
            Email3 = ''
            BusinessPhone = '555-1000'
            HomePhone = '555-2000'
            MobilePhone = '555-3000'
            OtherPhone = '555-4000'
            BusinessAddress = '1 Main <Street>'
            HomeAddress = '2 Oak & Pine'
            OtherAddress = '"Suite" <3>'
            WebPage = 'https://example.com/?q=<script>'
            Birthday = [datetime]'1990-05-01T00:00:00'
            Anniversary = [datetime]'2015-06-15T00:00:00'
            Categories = 'Blue & <Gold>'
            DistributionListMembers = @('One <Member>', 'Two & "Friend"')
            FolderPath = 'Mailbox\Contacts <VIP>'
            BodyText = 'Notes <script>bad()</script>'
            Notes = 'Notes <script>bad()</script>'
            MessageClass = 'IPM.Contact'
            EntryId = 'entry-<contact>'
            CreationTime = [datetime]'2024-03-01T08:00:00'
            LastModificationTime = [datetime]'2024-03-02T08:15:00'
            Attachments = @(
                [pscustomobject]@{
                    FileName = 'card<img>.vcf'
                    DisplayName = 'Contact "Card"'
                    Size = 512
                }
            )
        }

        Write-CalendarHtmlReport -Records @($calendarRecord) -PstItem $pstItem -ReportPath $calendarPath
        Write-ContactsHtmlReport -Records @($contactRecord) -PstItem $pstItem -ReportPath $contactsPath

        $calendarHtml = Get-Content -LiteralPath $calendarPath -Raw
        $contactsHtml = Get-Content -LiteralPath $contactsPath -Raw

        $calendarHtml | Should -Match '&lt;script&gt;alert\(1\)&lt;/script&gt;'
        $calendarHtml | Should -Match "data-search='[^']*&lt;script&gt;alert\(1\)&lt;/script&gt;"
        $calendarHtml | Should -Match 'Agenda &quot;Q2&quot;'
        $calendarHtml | Should -Match '&lt;Room&gt;'
        $calendarHtml | Should -Not -Match '<script>alert\(1\)</script>'
        $calendarHtml | Should -Not -Match '<b>unsafe</b>'

        $contactsHtml | Should -Match '&lt;img src=x onerror=alert\(2\)&gt;'
        $contactsHtml | Should -Match "data-search='[^']*&lt;img src=x onerror=alert\(2\)&gt;"
        $contactsHtml | Should -Match 'Two &amp; &quot;Friend&quot;'
        $contactsHtml | Should -Match '&lt;VIP&gt;'
        $contactsHtml | Should -Not -Match '<img src=x onerror=alert\(2\)>'
        $contactsHtml | Should -Not -Match '<script>bad\(\)</script>'
    }

    It 'serializes a single contact category as a JSON array in data-categories' {
        $contactsPath = Join-Path $TestDrive 'single-category-contacts.html'
        $pstItem = [pscustomobject]@{
            Name = 'Single Category PST.pst'
        }
        $contactRecord = [pscustomobject]@{
            DisplayName = 'One Category Contact'
            FullName = 'One Category Contact'
            FirstName = ''
            MiddleName = ''
            LastName = ''
            CompanyName = ''
            JobTitle = ''
            Department = ''
            Email1 = ''
            Email2 = ''
            Email3 = ''
            BusinessPhone = ''
            HomePhone = ''
            MobilePhone = ''
            OtherPhone = ''
            BusinessAddress = ''
            HomeAddress = ''
            OtherAddress = ''
            WebPage = ''
            Birthday = $null
            Anniversary = $null
            Categories = 'Only One'
            FolderPath = 'Mailbox\Contacts'
            Notes = ''
            BodyText = ''
            MessageClass = 'IPM.Contact'
            EntryId = 'single-category-contact'
            CreationTime = $null
            LastModificationTime = $null
            DistributionListMembers = @()
            Attachments = @()
        }

        Write-ContactsHtmlReport -Records @($contactRecord) -PstItem $pstItem -ReportPath $contactsPath

        $contactsHtml = Get-Content -LiteralPath $contactsPath -Raw
        $dataCategoriesMatch = [regex]::Match($contactsHtml, "data-categories='([^']+)'")
        $dataCategoriesMatch.Success | Should -BeTrue
        $jsonValue = [System.Net.WebUtility]::HtmlDecode($dataCategoriesMatch.Groups[1].Value)
        $parsed = $jsonValue | ConvertFrom-Json -NoEnumerate

        ($parsed -is [System.Array]) | Should -BeTrue
        @($parsed).Count | Should -Be 1
        @($parsed)[0] | Should -Be 'only one'
    }

    It 'renders missing calendar sensitivity with the normal empty-value treatment' {
        $calendarPath = Join-Path $TestDrive 'missing-sensitivity-calendar.html'
        $pstItem = [pscustomobject]@{ Name = 'Missing Sensitivity PST.pst' }
        $calendarRecord = [pscustomobject]@{
            SortTime = [datetime]'2024-07-01T09:00:00'
            StartTime = [datetime]'2024-07-01T09:00:00'
            EndTime = [datetime]'2024-07-01T10:00:00'
            Subject = 'No sensitivity value'
            ItemType = 'Appointment'
            AllDayEvent = $false
            Location = ''
            Organizer = ''
            RequiredAttendees = ''
            OptionalAttendees = ''
            IsRecurring = $false
            RecurrenceSummary = ''
            Categories = ''
            Sensitivity = $null
            FolderPath = 'Mailbox\Calendar'
            Notes = ''
            BodyText = ''
            MessageClass = 'IPM.Appointment'
            EntryId = 'missing-sensitivity'
            CreationTime = $null
            LastModificationTime = $null
            Attachments = @()
        }

        Write-CalendarHtmlReport -Records @($calendarRecord) -PstItem $pstItem -ReportPath $calendarPath

        $calendarHtml = Get-Content -LiteralPath $calendarPath -Raw
        $calendarHtml | Should -Match "(?s)<div class='detail-label'>Sensitivity</div><div class='detail-value'><span class='empty'>\(none\)</span>"
        $calendarHtml | Should -Not -Match 'Normal \(0\)'
    }

    It 'normalizes calendar filter end dates for records ending exactly at midnight' {
        $calendarPath = Join-Path $TestDrive 'midnight-end-calendar.html'
        $pstItem = [pscustomobject]@{
            Name = 'Midnight PST.pst'
        }
        $records = @(
            [pscustomobject]@{
                SortTime = [datetime]'2024-07-01T00:00:00'
                StartTime = [datetime]'2024-07-01T00:00:00'
                EndTime = [datetime]'2024-07-02T00:00:00'
                Subject = 'Independence Day'
                ItemType = 'Appointment'
                AllDayEvent = $true
                Location = ''
                Organizer = ''
                RequiredAttendees = ''
                OptionalAttendees = ''
                IsRecurring = $false
                RecurrenceSummary = ''
                Categories = ''
                Sensitivity = 0
                FolderPath = 'Mailbox\Calendar'
                Notes = ''
                BodyText = ''
                MessageClass = 'IPM.Appointment'
                EntryId = 'all-day-midnight'
                CreationTime = $null
                LastModificationTime = $null
                Attachments = @()
            },
            [pscustomobject]@{
                SortTime = [datetime]'2024-07-01T18:00:00'
                StartTime = [datetime]'2024-07-01T18:00:00'
                EndTime = [datetime]'2024-07-02T00:00:00'
                Subject = 'Ends At Midnight'
                ItemType = 'Meeting'
                AllDayEvent = $false
                Location = ''
                Organizer = ''
                RequiredAttendees = ''
                OptionalAttendees = ''
                IsRecurring = $false
                RecurrenceSummary = ''
                Categories = ''
                Sensitivity = 0
                FolderPath = 'Mailbox\Calendar'
                Notes = ''
                BodyText = ''
                MessageClass = 'IPM.Schedule.Meeting.Request'
                EntryId = 'meeting-midnight'
                CreationTime = $null
                LastModificationTime = $null
                Attachments = @()
            }
        )

        Write-CalendarHtmlReport -Records $records -PstItem $pstItem -ReportPath $calendarPath

        $calendarHtml = Get-Content -LiteralPath $calendarPath -Raw
        $allDayMatch = [regex]::Match($calendarHtml, "<article class='conversation record-card'[^>]*data-end-date='([^']+)'[^>]*>[\s\S]*?<h2>Independence Day</h2>")
        $midnightMeetingMatch = [regex]::Match($calendarHtml, "<article class='conversation record-card'[^>]*data-end-date='([^']+)'[^>]*>[\s\S]*?<h2>Ends At Midnight</h2>")

        $allDayMatch.Success | Should -BeTrue
        $midnightMeetingMatch.Success | Should -BeTrue
        $allDayMatch.Groups[1].Value | Should -Be '2024-07-01'
        $midnightMeetingMatch.Groups[1].Value | Should -Be '2024-07-01'
    }

    It 'writes each report footer with its own typed log path in multi-report runs' {
        $baseOutputPath = Join-Path $TestDrive 'footer-logs.html'
        $baseLogPath = Join-Path $TestDrive 'footer-logs.log'
        $result = & $script:invokePwshScriptCapture -FilePath $script:corePath -Arguments @(
            '-UseSampleData',
            '-NoPrompt',
            '-TeamsReport:$false',
            '-EmailReport:$false',
            '-CalendarReport:$true',
            '-ContactsReport:$true',
            '-OutputPath', $baseOutputPath,
            '-LogPath', $baseLogPath
        )

        $calendarOutputPath = Join-Path $TestDrive 'footer-logs_Calendar.html'
        $contactsOutputPath = Join-Path $TestDrive 'footer-logs_Contacts.html'
        $calendarHtml = Get-Content -LiteralPath $calendarOutputPath -Raw
        $contactsHtml = Get-Content -LiteralPath $contactsOutputPath -Raw

        $result.ExitCode | Should -Be 0
        $calendarHtml | Should -Match 'Log file: .*footer-logs_Calendar\.log'
        $contactsHtml | Should -Match 'Log file: .*footer-logs_Contacts\.log'
        $contactsHtml | Should -Not -Match 'Log file: .*footer-logs_Calendar\.log'
    }
}
