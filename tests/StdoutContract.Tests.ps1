Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Stdout JSON schema contract' {
    BeforeAll {
        $schemaPath = Join-Path $PSScriptRoot '..\docs\schemas\stdout-contract.schema.json'
        $script:schema = Get-Content -LiteralPath $schemaPath -Raw
        $script:validRunId = '0123456789abcdef0123456789abcdef'
        $script:cases = @(
            [ordered]@{
                prefix = 'CONVERSION_PROGRESS'
                RunId = $script:validRunId
                ItemsAttempted = 1
                ItemsExported = 1
                FoldersScanned = 1
                ItemReadFailures = 0
                ElapsedSeconds = 1
                RatePerMinute = 60
                FolderPath = 'Inbox'
            },
            [ordered]@{
                prefix = 'CONVERSION_STAGE'
                RunId = $script:validRunId
                Stage = 'FinishedReading'
            },
            [ordered]@{
                prefix = 'CONVERSION_RESULT'
                RunId = $script:validRunId
                OutputPath = 'C:\Reports\case.html'
                LogPath = 'C:\Reports\case.log'
                TeamsOutputPath = 'C:\Reports\case_Teams.html'
                EmailOutputPath = 'C:\Reports\case_Email.db'
                CalendarOutputPath = 'C:\Reports\case_Calendar.html'
                ContactsOutputPath = 'C:\Reports\case_Contacts.html'
                TeamsLogPath = 'C:\Reports\case_Teams.log'
                EmailLogPath = 'C:\Reports\case_Email.log'
                CalendarLogPath = 'C:\Reports\case_Calendar.log'
                ContactsLogPath = 'C:\Reports\case_Contacts.log'
                ItemsExported = 6
                TeamsItemsExported = 2
                EmailItemsExported = 2
                CalendarItemsExported = 1
                ContactsItemsExported = 1
                ItemReadFailures = 0
                AttachmentReadFailures = 0
                SubfolderScanFailures = 0
            },
            [ordered]@{
                prefix = 'CONVERSION_ERROR'
                RunId = $script:validRunId
                ExitCode = '1'
                Message = 'Sample failure'
            }
        )
    }

    It 'accepts a valid RunId on every machine line type' {
        foreach ($case in $script:cases) {
            (($case | ConvertTo-Json -Compress) | Test-Json -Schema $script:schema -ErrorAction Stop) |
                Should -BeTrue -Because "valid $($case.prefix) must satisfy the contract"
        }
    }

    It 'rejects an invalid RunId on every machine line type' {
        foreach ($case in $script:cases) {
            $invalid = [ordered]@{}
            foreach ($key in $case.Keys) { $invalid[$key] = $case[$key] }
            $invalid.RunId = 'NOT-A-VALID-RUN-ID'

            $accepted = (($invalid | ConvertTo-Json -Compress) | Test-Json -Schema $script:schema -ErrorAction SilentlyContinue) 2>$null
            $accepted | Should -BeFalse -Because "invalid RunId must be rejected for $($case.prefix)"
        }
    }

    It 'rejects unknown properties on every machine line type' {
        foreach ($case in $script:cases) {
            $unknown = [ordered]@{}
            foreach ($key in $case.Keys) { $unknown[$key] = $case[$key] }
            $unknown.UnexpectedField = 'not contracted'

            $accepted = (($unknown | ConvertTo-Json -Compress) | Test-Json -Schema $script:schema -ErrorAction SilentlyContinue) 2>$null
            $accepted | Should -BeFalse -Because "unknown properties must be rejected for $($case.prefix)"
        }
    }
}
