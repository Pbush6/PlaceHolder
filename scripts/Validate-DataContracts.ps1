<#
.SYNOPSIS
Validates Purview data contracts: stdout samples, inflight.json, memory index.json.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Join-Path $PSScriptRoot '..')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-SchemaShape {
    param(
        [Parameter(Mandatory = $true)][object]$Data,
        [Parameter(Mandatory = $true)][string[]]$RequiredKeys
    )
    foreach ($key in $RequiredKeys) {
        $hasKey = $false
        if ($Data -is [System.Collections.IDictionary]) {
            $hasKey = $Data.Contains($key)
        }
        else {
            $hasKey = ($null -ne $Data.PSObject.Properties[$key])
        }
        if (-not $hasKey) { throw "Missing required key: $key" }
    }
}

function Test-StdoutLineShape {
    param(
        [Parameter(Mandatory = $true)][string]$Line,
        [Parameter(Mandatory = $true)][string]$ExpectedPrefix,
        [Parameter(Mandatory = $true)][string[]]$RequiredFields
    )
    if (-not $Line.StartsWith("$ExpectedPrefix|")) { throw "Line does not start with $ExpectedPrefix" }
    $fields = @{}
    foreach ($part in ($Line -split '\|')) {
        $idx = $part.IndexOf('=')
        if ($idx -gt 0) { $fields[$part.Substring(0, $idx)] = $part.Substring($idx + 1) }
    }
    foreach ($req in $RequiredFields) {
        if (-not $fields.ContainsKey($req)) { throw "$ExpectedPrefix missing field: $req" }
    }
}

$failures = [System.Collections.Generic.List[string]]::new()

try {
    $core = Join-Path $RepoRoot 'src\Convert-PurviewTeamsPstToHtml.ps1'
    $report = Join-Path $env:TEMP "purview-contract-$([guid]::NewGuid().ToString('N')).html"
    $log = Join-Path $env:TEMP "purview-contract-$([guid]::NewGuid().ToString('N')).log"
    $reportBase = [IO.Path]::Combine([IO.Path]::GetDirectoryName($report), [IO.Path]::GetFileNameWithoutExtension($report))
    $logBase = [IO.Path]::Combine([IO.Path]::GetDirectoryName($log), [IO.Path]::GetFileNameWithoutExtension($log))
    $generatedPaths = @(
        "${reportBase}_Teams.html", "${reportBase}_Email.db", "${reportBase}_Calendar.html", "${reportBase}_Contacts.html",
        "${reportBase}_Dashboard.html",
        (Join-Path ([IO.Path]::GetDirectoryName($report)) 'Open-EmailReport.cmd'),
        "${logBase}_Teams.log", "${logBase}_Email.log", "${logBase}_Calendar.log", "${logBase}_Contacts.log"
    )
    $runId = [guid]::NewGuid().ToString('N')

    $stdout = & pwsh -NoProfile -File $core -UseSampleData -NoPrompt -OutputPath $report -LogPath $log -RunId $runId 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "Core sample run failed with exit $LASTEXITCODE" }

    $lines = $stdout -split "`r?`n" | Where-Object { $_ -match '^CONVERSION_' }
    if (-not ($lines | Where-Object { $_ -like 'CONVERSION_RESULT|*' })) {
        throw 'Sample run did not emit CONVERSION_RESULT'
    }

    $result = ($lines | Where-Object { $_ -like 'CONVERSION_RESULT|*' } | Select-Object -Last 1)
    Test-StdoutLineShape -Line $result -ExpectedPrefix 'CONVERSION_RESULT' -RequiredFields @(
        'OutputPath', 'LogPath', 'ItemsExported', 'ItemReadFailures', 'AttachmentReadFailures', 'SubfolderScanFailures',
        'TeamsOutputPath', 'EmailOutputPath', 'CalendarOutputPath', 'ContactsOutputPath', 'DashboardOutputPath',
        'TeamsLogPath', 'EmailLogPath', 'CalendarLogPath', 'ContactsLogPath',
        'TeamsItemsExported', 'EmailItemsExported', 'CalendarItemsExported', 'ContactsItemsExported'
    )
    if ($result -notmatch 'ItemsExported=6') { throw 'CONVERSION_RESULT missing expected ItemsExported=6' }
    if ($result -notmatch 'TeamsItemsExported=2') { throw 'CONVERSION_RESULT missing expected TeamsItemsExported=2' }
    if ($result -notmatch 'EmailItemsExported=2') { throw 'CONVERSION_RESULT missing expected EmailItemsExported=2' }
    if ($result -notmatch 'CalendarItemsExported=1') { throw 'CONVERSION_RESULT missing expected CalendarItemsExported=1' }
    if ($result -notmatch 'ContactsItemsExported=1') { throw 'CONVERSION_RESULT missing expected ContactsItemsExported=1' }
    if ($result -notmatch 'TeamsOutputPath=.*_Teams\.html') { throw 'CONVERSION_RESULT missing expected TeamsOutputPath' }
    if ($result -notmatch 'EmailOutputPath=.*_Email\.db') { throw 'CONVERSION_RESULT missing expected EmailOutputPath' }
    if ($result -notmatch 'CalendarOutputPath=.*_Calendar\.html') { throw 'CONVERSION_RESULT missing expected CalendarOutputPath' }
    if ($result -notmatch 'ContactsOutputPath=.*_Contacts\.html') { throw 'CONVERSION_RESULT missing expected ContactsOutputPath' }
    if ($result -notmatch 'DashboardOutputPath=.*_Dashboard\.html') { throw 'CONVERSION_RESULT missing expected DashboardOutputPath' }
    if ($result -notmatch 'TeamsLogPath=.*_Teams\.log') { throw 'CONVERSION_RESULT missing expected TeamsLogPath' }
    if ($result -notmatch 'EmailLogPath=.*_Email\.log') { throw 'CONVERSION_RESULT missing expected EmailLogPath' }
    if ($result -notmatch 'CalendarLogPath=.*_Calendar\.log') { throw 'CONVERSION_RESULT missing expected CalendarLogPath' }
    if ($result -notmatch 'ContactsLogPath=.*_Contacts\.log') { throw 'CONVERSION_RESULT missing expected ContactsLogPath' }
    foreach ($generatedPath in $generatedPaths) {
        if (-not (Test-Path -LiteralPath $generatedPath -PathType Leaf)) {
            throw "Sample run did not create selected typed output: $generatedPath"
        }
    }
    $contractObject = [ordered]@{ prefix = 'CONVERSION_RESULT' }
    $integerFields = @(
        'ItemsExported', 'TeamsItemsExported', 'EmailItemsExported', 'CalendarItemsExported', 'ContactsItemsExported',
        'ItemReadFailures', 'AttachmentReadFailures', 'SubfolderScanFailures'
    )
    foreach ($part in ($result -split '\|' | Select-Object -Skip 1)) {
        $idx = $part.IndexOf('=')
        if ($idx -le 0) { continue }
        $key = $part.Substring(0, $idx)
        $value = $part.Substring($idx + 1)
        $contractObject[$key] = if ($key -in $integerFields) { [int64]$value } else { $value }
    }
    $stdoutSchema = Get-Content -LiteralPath (Join-Path $RepoRoot 'docs\schemas\stdout-contract.schema.json') -Raw
    if (-not (($contractObject | ConvertTo-Json -Compress) | Test-Json -Schema $stdoutSchema -ErrorAction Stop)) {
        throw 'CONVERSION_RESULT did not validate against stdout-contract.schema.json'
    }
    $invalidRunIdObject = [ordered]@{}
    foreach ($key in $contractObject.Keys) { $invalidRunIdObject[$key] = $contractObject[$key] }
    $invalidRunIdObject.RunId = 'INVALID-RUN-ID'
    $invalidRunIdAccepted = (($invalidRunIdObject | ConvertTo-Json -Compress) |
        Test-Json -Schema $stdoutSchema -ErrorAction SilentlyContinue) 2>$null
    if ($invalidRunIdAccepted) {
        throw 'stdout-contract.schema.json accepted an invalid RunId'
    }
    $unknownPropertyObject = [ordered]@{}
    foreach ($key in $contractObject.Keys) { $unknownPropertyObject[$key] = $contractObject[$key] }
    $unknownPropertyObject.UnexpectedField = 'not contracted'
    $unknownPropertyAccepted = (($unknownPropertyObject | ConvertTo-Json -Compress) |
        Test-Json -Schema $stdoutSchema -ErrorAction SilentlyContinue) 2>$null
    if ($unknownPropertyAccepted) {
        throw 'stdout-contract.schema.json accepted an unknown property'
    }

    $progress = ($lines | Where-Object { $_ -like 'CONVERSION_PROGRESS|*' } | Select-Object -First 1)
    if ($progress) {
        Test-StdoutLineShape -Line $progress -ExpectedPrefix 'CONVERSION_PROGRESS' -RequiredFields @(
            'ItemsAttempted', 'ItemsExported', 'FoldersScanned', 'ItemReadFailures', 'ElapsedSeconds', 'RatePerMinute', 'FolderPath'
        )
        if ($progress -notmatch "RunId=$runId") { throw 'CONVERSION_PROGRESS missing expected RunId' }
    }
}
catch {
    $failures.Add("Stdout contract: $($_.Exception.Message)")
}
finally {
    foreach ($p in @($report, $log) + @($generatedPaths)) {
        if ($p -and (Test-Path -LiteralPath $p)) { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
    }
}

try {
    $inflightPath = Join-Path $RepoRoot '..\.cursor\state\inflight.json' | Resolve-Path -ErrorAction SilentlyContinue
    if ($inflightPath) {
        $inflight = Get-Content -LiteralPath $inflightPath -Raw | ConvertFrom-Json
        Test-SchemaShape -Data $inflight -RequiredKeys @('version', 'lastUpdated', 'tasks')
        foreach ($task in $inflight.tasks) {
            Test-SchemaShape -Data $task -RequiredKeys @('id', 'title', 'status', 'priority')
        }
    }
}
catch {
    $failures.Add("inflight.json: $($_.Exception.Message)")
}

try {
    $memoryIndex = 'C:\Users\pbush\OneDrive - Perfection Learning\Documents\AI\Cursor Output\Curt-Memory\index.json'
    if (Test-Path -LiteralPath $memoryIndex) {
        $index = Get-Content -LiteralPath $memoryIndex -Raw | ConvertFrom-Json
        Test-SchemaShape -Data $index -RequiredKeys @('version', 'lastUpdated', 'entries')
        foreach ($entry in $index.entries) {
            Test-SchemaShape -Data $entry -RequiredKeys @('id', 'type', 'path', 'tags', 'summary', 'created', 'updated')
        }
    }
}
catch {
    $failures.Add("memory index.json: $($_.Exception.Message)")
}

if ($failures.Count -gt 0) {
    Write-Output 'CONTRACT_VALIDATION_FAILED'
    $failures | ForEach-Object { Write-Output $_ }
    exit 1
}

Write-Output 'CONTRACT_VALIDATION_PASSED'
exit 0
