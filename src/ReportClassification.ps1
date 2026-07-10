# Source of truth; build.ps1 will inline into core for packaging.
Set-StrictMode -Version Latest

function Test-IsTeamsMessagesFolder {
    param([Parameter(Mandatory = $true)][string]$FolderPath)
    return $FolderPath -match '(?i)(^|\\)(TeamsMessagesData|TeamsMeetings|Migrated-Teams-Chat|SubstrateHolds)(\\|$)'
}

function Test-IsEmailMessageClass {
    param([AllowNull()][string]$MessageClass)
    if ([string]::IsNullOrWhiteSpace($MessageClass)) { return $false }
    if ($MessageClass -match '(?i)^IPM\.Schedule\.Meeting') { return $false }
    return $MessageClass -match '(?i)^IPM\.Note'
}

function Get-ItemReportBucket {
    param(
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [AllowNull()][string]$MessageClass
    )
    if (Test-IsTeamsMessagesFolder -FolderPath $FolderPath) { return 'Teams' }
    if (Test-IsEmailMessageClass -MessageClass $MessageClass) { return 'Email' }
    return 'Skip'
}
