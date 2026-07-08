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
    $runId = [guid]::NewGuid().ToString('N')

    $stdout = & pwsh -NoProfile -File $core -UseSampleData -NoPrompt -OutputPath $report -LogPath $log -RunId $runId 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "Core sample run failed with exit $LASTEXITCODE" }

    $lines = $stdout -split "`r?`n" | Where-Object { $_ -match '^CONVERSION_' }
    if (-not ($lines | Where-Object { $_ -like 'CONVERSION_RESULT|*' })) {
        throw 'Sample run did not emit CONVERSION_RESULT'
    }

    $result = ($lines | Where-Object { $_ -like 'CONVERSION_RESULT|*' } | Select-Object -Last 1)
    Test-StdoutLineShape -Line $result -ExpectedPrefix 'CONVERSION_RESULT' -RequiredFields @(
        'OutputPath', 'LogPath', 'ItemsExported', 'ItemReadFailures', 'AttachmentReadFailures', 'SubfolderScanFailures'
    )

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
    foreach ($p in @($report, $log)) {
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
