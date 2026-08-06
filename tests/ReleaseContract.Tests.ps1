Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Release 1.3.0.0 contract' -Tag 'Build', 'Release' {
    BeforeAll {
        $script:repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
        $script:readmePath = Join-Path $script:repoRoot 'README.md'
        $script:buildPath = Join-Path $script:repoRoot 'build.ps1'
        $script:packagePath = Join-Path $script:repoRoot 'scripts\Build-DeploymentPackage.ps1'
        $script:viewerProjectPath = Join-Path $script:repoRoot 'EmailReviewViewer\EmailReviewViewer.App\EmailReviewViewer.App.csproj'
        $script:launcherPath = Join-Path $script:repoRoot 'src\Start-PurviewTeamsPstToHtmlApp.ps1'
        $script:corePath = Join-Path $script:repoRoot 'src\Convert-PurviewTeamsPstToHtml.ps1'
        $script:contractsPath = Join-Path $script:repoRoot 'docs\DATA_CONTRACTS.md'
    }

    It 'keeps every executable and package version at 1.3.0.0' {
        (Get-Content -LiteralPath $script:readmePath -Raw) |
            Should -Match '(?m)^\*\*Current version:\*\*\s*1\.3\.0\.0\s*$'
        (Get-Content -LiteralPath $script:buildPath -Raw) |
            Should -Match "\[string\]\`$Version\s*=\s*'1\.3\.0\.0'"

        [xml]$viewerProject = Get-Content -LiteralPath $script:viewerProjectPath -Raw
        [string]$viewerProject.Project.PropertyGroup.Version | Should -Be '1.3.0.0'
        [string]$viewerProject.Project.PropertyGroup.FileVersion | Should -Be '1.3.0.0'
        [string]$viewerProject.Project.PropertyGroup.AssemblyVersion | Should -Be '1.3.0.0'

        $packageText = Get-Content -LiteralPath $script:packagePath -Raw
        $packageText | Should -Match 'PurviewTeamsPstToHtmlApp-\$version-win-x64'
        $packageText | Should -Match 'Version mismatch'
    }

    It 'keeps the embedded launcher core byte-for-byte current' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw
        $match = [regex]::Match($launcherText, '(?m)^\$script:EmbeddedCoreBase64 = ''([^'']*)''')
        $match.Success | Should -BeTrue
        $embeddedHash = [Convert]::ToHexString(
            [Security.Cryptography.SHA256]::HashData([Convert]::FromBase64String($match.Groups[1].Value))
        )
        $sourceHash = [Convert]::ToHexString(
            [Security.Cryptography.SHA256]::HashData([IO.File]::ReadAllBytes($script:corePath))
        )
        $embeddedHash | Should -Be $sourceHash
    }

    It 'packages the release converter, debug tool, and final Email Reviewer' {
        $packageText = Get-Content -LiteralPath $script:packagePath -Raw
        foreach ($artifact in @(
            'PurviewTeamsPstToHtmlConverter.exe',
            'PurviewTeamsPstToHtmlConverter_Debug.exe',
            'EmailReviewViewer.App.exe',
            'README.txt',
            'RELEASE_NOTES.md',
            'Verify-Prerequisites.ps1',
            'checksums.sha256'
        )) {
            $packageText | Should -Match ([regex]::Escape($artifact))
        }
        $packageText | Should -Match 'By Patrick Bush'
    }

    It 'documents the four GUI reports, dashboard, typed outputs, filters, launches, and limitations' {
        $readme = Get-Content -LiteralPath $script:readmePath -Raw
        foreach ($requiredText in @(
            'Teams',
            'Email',
            'Calendar',
            'Contacts',
            '_Teams.html',
            '_Email.db',
            '_Calendar.html',
            '_Contacts.html',
            '_Dashboard.html',
            'GUI always generates',
            'search',
            'filter',
            'default browser',
            'EmailReviewViewer.App.exe',
            'Classic Microsoft Outlook',
            'PowerShell 7',
            'unsigned'
        )) {
            $readme | Should -Match ('(?i)' + [regex]::Escape($requiredText))
        }
    }

    It 'documents Contacts filters and release NoOutput automation accurately' {
        $readme = Get-Content -LiteralPath $script:readmePath -Raw
        $contracts = Get-Content -LiteralPath $script:contractsPath -Raw

        $readme | Should -Match 'Contacts HTML supports text search plus folder and category filters'
        $contracts | Should -Match 'dataset-backed text search, folder, and category filters'
        $readme | Should -Not -Match 'Contacts HTML supports .*type filters'
        $contracts | Should -Not -Match 'Contacts report:.*type filters'
        foreach ($text in @($readme, $contracts)) {
            $text | Should -Match '-NoOutput'
            $text | Should -Match 'Debug converter'
            $text | Should -Match 'typed artifacts'
        }
    }
}
