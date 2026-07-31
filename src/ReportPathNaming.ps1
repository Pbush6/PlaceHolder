# Source of truth; build.ps1 will inline into core/launcher for packaging.
Set-StrictMode -Version Latest

function Get-ReportPathBaseName {
    param([Parameter(Mandatory = $true)][string]$FilePath)
    $name = [IO.Path]::GetFileNameWithoutExtension($FilePath)
    if ($name -match '(?i)_Teams$') { $name = $name.Substring(0, $name.Length - 6) }
    elseif ($name -match '(?i)_Email$') { $name = $name.Substring(0, $name.Length - 6) }
    elseif ($name -match '(?i)_Calendar$') { $name = $name.Substring(0, $name.Length - 9) }
    elseif ($name -match '(?i)_Contacts$') { $name = $name.Substring(0, $name.Length - 9) }
    elseif ($name -match '(?i)_Dashboard$') { $name = $name.Substring(0, $name.Length - 10) }
    return $name
}

function Get-ReportOutputPaths {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayPath,
        [Parameter(Mandatory = $true)][bool]$TeamsReport,
        [Parameter(Mandatory = $true)][bool]$EmailReport,
        [bool]$CalendarReport = $false,
        [bool]$ContactsReport = $false
    )
    if (-not $TeamsReport -and -not $EmailReport -and -not $CalendarReport -and -not $ContactsReport) {
        throw 'At least one report type must be selected.'
    }
    $dir = [IO.Path]::GetDirectoryName($DisplayPath)
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = (Get-Location).Path }
    $inputExt = [IO.Path]::GetExtension($DisplayPath)
    $isLogPath = $inputExt -in @('.log', '.txt')
    $teamsExt = if ($isLogPath) { $inputExt } else { '.html' }
    $emailExt = if ($isLogPath) { $inputExt } else { '.db' }
    $calendarExt = if ($isLogPath) { $inputExt } else { '.html' }
    $contactsExt = if ($isLogPath) { $inputExt } else { '.html' }
    $displayExt = if ($isLogPath) { $inputExt } else { '.html' }
    $base = Get-ReportPathBaseName -FilePath $DisplayPath
    $teamsPath = if ($TeamsReport) { Join-Path $dir ($base + '_Teams' + $teamsExt) } else { $null }
    $emailPath = if ($EmailReport) { Join-Path $dir ($base + '_Email' + $emailExt) } else { $null }
    $calendarPath = if ($CalendarReport) { Join-Path $dir ($base + '_Calendar' + $calendarExt) } else { $null }
    $contactsPath = if ($ContactsReport) { Join-Path $dir ($base + '_Contacts' + $contactsExt) } else { $null }
    $selectedPaths = @(@($teamsPath, $emailPath, $calendarPath, $contactsPath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $display = if (@($selectedPaths).Count -eq 1) { $selectedPaths[0] } else { Join-Path $dir ($base + $displayExt) }
    [pscustomobject]@{
        DisplayPath = $display
        TeamsPath = $teamsPath
        EmailPath = $emailPath
        CalendarPath = $calendarPath
        ContactsPath = $contactsPath
    }
}

function Get-DashboardOutputPath {
    param([Parameter(Mandatory = $true)][string]$DisplayPath)
    $dir = [IO.Path]::GetDirectoryName($DisplayPath)
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = (Get-Location).Path }
    return (Join-Path $dir ((Get-ReportPathBaseName -FilePath $DisplayPath) + '_Dashboard.html'))
}

function Get-WritePathsFromResultFields {
    # Prefer typed report write targets. Never treat a multi-report DisplayPath (OutputPath /
    # LogPath when multiple reports are selected and typed paths exist) as a required file.
    param([AllowNull()][hashtable]$Fields)
    if ($null -eq $Fields) { $Fields = @{} }
    $reportPaths = [System.Collections.Generic.List[string]]::new()
    $hasTypedReport = $false
    foreach ($key in @('TeamsOutputPath', 'EmailOutputPath', 'CalendarOutputPath', 'ContactsOutputPath')) {
        if ($Fields.ContainsKey($key)) {
            $path = [string]$Fields[$key]
            if (-not [string]::IsNullOrWhiteSpace($path)) {
                $hasTypedReport = $true
                if (-not $reportPaths.Contains($path)) { [void]$reportPaths.Add($path) }
            }
        }
    }
    if (-not $hasTypedReport -and $Fields.ContainsKey('OutputPath')) {
        $path = [string]$Fields['OutputPath']
        if (-not [string]::IsNullOrWhiteSpace($path)) { [void]$reportPaths.Add($path) }
    }
    $logPaths = [System.Collections.Generic.List[string]]::new()
    $hasTypedLog = $false
    foreach ($key in @('TeamsLogPath', 'EmailLogPath', 'CalendarLogPath', 'ContactsLogPath')) {
        if ($Fields.ContainsKey($key)) {
            $path = [string]$Fields[$key]
            if (-not [string]::IsNullOrWhiteSpace($path)) {
                $hasTypedLog = $true
                if (-not $logPaths.Contains($path)) { [void]$logPaths.Add($path) }
            }
        }
    }
    if (-not $hasTypedLog -and $Fields.ContainsKey('LogPath')) {
        $path = [string]$Fields['LogPath']
        if (-not [string]::IsNullOrWhiteSpace($path)) { [void]$logPaths.Add($path) }
    }
    [pscustomobject]@{
        ReportPaths = $reportPaths.ToArray()
        LogPaths = $logPaths.ToArray()
    }
}
