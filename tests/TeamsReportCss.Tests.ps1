Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Teams viewer CSS stays aligned with the HTML report' {
    It 'embeds the Get-ReportCss colors, fonts, and shapes' {
        $repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
        $corePath = Join-Path $repoRoot 'src\Convert-PurviewTeamsPstToHtml.ps1'
        $cssPath = Join-Path $repoRoot 'EmailReviewViewer\EmailReviewViewer.App\Assets\TeamsReport.css'
        $core = [IO.File]::ReadAllText($corePath)
        $match = [regex]::Match($core, "(?s)function Get-ReportCss \{.*?return @'(.*?)'@")
        $match.Success | Should -BeTrue
        $reportCss = $match.Groups[1].Value.Trim() -replace "`r`n", "`n"
        $viewerCss = ([IO.File]::ReadAllText($cssPath) -replace "`r`n", "`n").Trim()

        $viewerCss.Contains($reportCss) | Should -BeTrue
        $viewerCss | Should -Match 'font-family: "Segoe UI", Arial, sans-serif'
        $viewerCss | Should -Match 'linear-gradient\(135deg, #213b78, #5c7dde\)'
        $viewerCss | Should -Match '\.sender-blue \{ border-left-color: #2f6fec; \}'
    }
}
