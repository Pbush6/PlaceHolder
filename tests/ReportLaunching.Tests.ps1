Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Report launching' {
    BeforeAll {
        $script:launcherPath = Join-Path $PSScriptRoot '..\src\Start-PurviewTeamsPstToHtmlApp.ps1'
        $tokens = $null
        $errors = $null
        $script:launcherAst = [Management.Automation.Language.Parser]::ParseFile(
            $script:launcherPath,
            [ref]$tokens,
            [ref]$errors)
        $errors | Should -BeNullOrEmpty

        function Import-LauncherFunction([string]$Name) {
            $functionAst = $script:launcherAst.Find({
                param($node)
                $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq $Name
            }, $true)
            $functionAst | Should -Not -BeNullOrEmpty
            $definition = $functionAst.Extent.Text -replace
                ('^function\s+' + [regex]::Escape($Name)),
                ('function script:' + $Name)
            Invoke-Expression $definition
        }

        Import-LauncherFunction 'ConvertTo-NativeArgumentString'
        Import-LauncherFunction 'Get-CurrentExecutablePath'
        Import-LauncherFunction 'Resolve-EmailViewerPath'
    }

    It 'resolves the viewer beside the packaged converter process' {
        $converterDirectory = Join-Path $TestDrive 'Delivered Folder'
        $viewerDirectory = Join-Path $converterDirectory 'EmailReviewViewer'
        New-Item -ItemType Directory -Path $viewerDirectory -Force | Out-Null
        $converterPath = Join-Path $converterDirectory 'PurviewTeamsPstToHtmlConverter.exe'
        $viewerPath = Join-Path $viewerDirectory 'EmailReviewViewer.App.exe'
        Set-Content -LiteralPath $converterPath -Value ''
        Set-Content -LiteralPath $viewerPath -Value ''

        Resolve-EmailViewerPath -ProcessPath $converterPath -ScriptDirectory '' -AppBaseDirectory '' |
            Should -Be ([IO.Path]::GetFullPath($viewerPath))
    }

    It 'resolves the viewer from a source launcher tree' {
        $sourceDirectory = Join-Path $TestDrive 'repo\src'
        $viewerDirectory = Join-Path $TestDrive 'repo\EmailReviewViewer\EmailReviewViewer.App\bin\Release\net8.0-windows'
        New-Item -ItemType Directory -Path $sourceDirectory, $viewerDirectory -Force | Out-Null
        $viewerPath = Join-Path $viewerDirectory 'EmailReviewViewer.App.exe'
        Set-Content -LiteralPath $viewerPath -Value ''

        Resolve-EmailViewerPath -ProcessPath 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' `
            -ScriptDirectory $sourceDirectory -AppBaseDirectory 'C:\Windows\System32\WindowsPowerShell\v1.0\' |
            Should -Be ([IO.Path]::GetFullPath($viewerPath))
    }

    It 'quotes database paths for Windows native process arguments' {
        $databasePath = 'C:\Review files\O''Brien (Final)\mail "Q4".db'

        ConvertTo-NativeArgumentString -Argument @($databasePath) |
            Should -Be '"C:\Review files\O''Brien (Final)\mail \"Q4\".db"'
    }

    It 'uses the Windows PowerShell compatible Arguments property' {
        $launcherText = Get-Content -LiteralPath $script:launcherPath -Raw

        $launcherText | Should -Not -Match '\.ArgumentList\.Add'
        $launcherText | Should -Match '\$psi\.Arguments\s*=\s*ConvertTo-NativeArgumentString'
    }
}
