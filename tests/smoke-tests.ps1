#Requires -Version 7.0

BeforeAll {
    Set-StrictMode -Version Latest
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:Source = Join-Path -Path $script:RepoRoot -ChildPath 'src'
    $script:OnboardScript = Join-Path -Path $script:Source -ChildPath 'New-M365UserFromCsv.ps1'
    $script:DisableScript = Join-Path -Path $script:Source -ChildPath 'Disable-M365User.ps1'
    $script:LicenseScript = Join-Path -Path $script:Source -ChildPath 'Remove-M365UserLicenses.ps1'
    $script:GroupScript = Join-Path -Path $script:Source -ChildPath 'Set-M365UserGroups.ps1'
    $script:ExportScript = Join-Path -Path $script:Source -ChildPath 'Export-M365UserAudit.ps1'
    $script:ConnectScript = Join-Path -Path $script:Source -ChildPath 'Connect-M365Graph.ps1'
    $script:OnboardSample = Join-Path -Path $script:RepoRoot -ChildPath 'samples/users-onboard.sample.csv'
    $script:OffboardSample = Join-Path -Path $script:RepoRoot -ChildPath 'samples/users-offboard.sample.csv'
    Import-Module -Name (Join-Path -Path $script:Source -ChildPath 'M365UserLifecycle.Helpers.psm1') -Force
}

Describe 'Script files parse and expose WhatIf' {
    It 'parses every PowerShell file under src' {
        $files = @(Get-ChildItem -LiteralPath $script:Source -Include '*.ps1', '*.psm1' -Recurse)
        $files.Count | Should -BeGreaterThan 5
        foreach ($file in $files) {
            $tokens = $null
            $parseErrors = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
            @($parseErrors).Count | Should -Be 0
        }
    }

    It 'gives every mutating script WhatIf and Confirm support' {
        foreach ($path in @($script:OnboardScript, $script:DisableScript, $script:LicenseScript, $script:GroupScript)) {
            $command = Get-Command -Name $path
            $command.Parameters.ContainsKey('WhatIf') | Should -BeTrue
            $command.Parameters.ContainsKey('Confirm') | Should -BeTrue
            $text = Get-Content -LiteralPath $path -Raw
            $text | Should -Match 'SupportsShouldProcess'
        }
    }

    It 'uses High confirm impact for destructive scripts and Medium for onboarding' {
        (Get-Content -LiteralPath $script:OnboardScript -Raw) | Should -Match "ConfirmImpact = 'Medium'"
        foreach ($path in @($script:DisableScript, $script:LicenseScript, $script:GroupScript)) {
            (Get-Content -LiteralPath $path -Raw) | Should -Match "ConfirmImpact = 'High'"
        }
    }

    It 'calls the Graph SDK cmdlets from the graph file rather than legacy modules' {
            $graph = Get-Content -LiteralPath (Join-Path $script:Source 'M365UserLifecycle.Graph.ps1') -Raw
            foreach ($command in @(
                'New-MgUser'
                'Update-MgUser'
                'Set-MgUserLicense'
                'Revoke-MgUserSignInSession'
                'New-MgGroupMember'
                'Remove-MgGroupMemberDirectoryObjectByRef'
                'Set-MgUserManagerByRef'
                'Get-MgSubscribedSku'
            )) {
                $graph.Contains($command) | Should -BeTrue
            }

            $combined = Get-ChildItem -LiteralPath $script:Source -Filter '*.ps1' -Recurse | Get-Content -Raw
            $moduleText = Get-Content -LiteralPath (Join-Path $script:Source 'M365UserLifecycle.Helpers.psm1') -Raw
            $sourceText = ($combined + $moduleText) -join "`n"
            $sourceText.Contains('Connect-MsolService') | Should -BeFalse
            $sourceText.Contains('Connect-AzureAD') | Should -BeFalse
            $sourceText.Contains('Set-MsolUser') | Should -BeFalse
    }
}

Describe 'WhatIf previews do not contact Graph' {
    It 'previews the onboarding sample' {
        $log = Join-Path $TestDrive 'onboard-logs'
        $validation = Test-M365OnboardCsv -Path $script:OnboardSample
        $expected = 0
        foreach ($row in @($validation.Rows)) {
            $expected += @(Get-M365OnboardPlan -Row $row -UsageLocation 'AU').Count
        }

        $results = @( & $script:OnboardScript -CsvPath $script:OnboardSample -WhatIf -LogDirectory $log )
        $results.Count | Should -Be $expected
        @($results | Where-Object { $_.Result -ne 'WhatIf' }).Count | Should -Be 0
        $results.Operation | Should -Contain 'Create Entra ID user'
        foreach ($item in $results) {
            $item.PSObject.Properties.Name | Should -Not -Contain 'TemporaryPassword'
        }

        Test-Path -LiteralPath $log | Should -BeFalse
    }

    It 'previews the offboarding sample' {
        $log = Join-Path $TestDrive 'offboard-logs'
        $validation = Test-M365OffboardCsv -Path $script:OffboardSample
        $expected = 0
        foreach ($row in @($validation.Rows)) {
            $expected += @(Get-M365OffboardPlan -User $row).Count
        }

        $results = @( & $script:DisableScript -CsvPath $script:OffboardSample -WhatIf -LogDirectory $log )
        $results.Count | Should -Be $expected
        $results.Operation | Should -Contain 'Disable sign-in'
        $results.Operation | Should -Contain 'Revoke sign-in sessions'
        $results.Operation | Should -Contain 'Record shared-mailbox conversion note'
        Test-Path -LiteralPath $log | Should -BeFalse
    }

    It 'previews license removal for the sample users' {
        $log = Join-Path $TestDrive 'license-logs'
        $results = @( & $script:LicenseScript -CsvPath $script:OffboardSample -WhatIf -LogDirectory $log )
        $results.Count | Should -Be 3
        @($results | Where-Object { $_.Result -eq 'WhatIf' }).Count | Should -Be 3
        Test-Path -LiteralPath $log | Should -BeFalse
    }

    It 'previews a group membership change for a fictional group id' {
        $csv = Join-Path $TestDrive 'groups.csv'
        @(
            'UserPrincipalName,GroupIds'
            'alex.example@contoso.com,11111111-2222-3333-4444-555555555555'
        ) | Set-Content -LiteralPath $csv -Encoding utf8
        $log = Join-Path $TestDrive 'group-logs'
        $results = @( & $script:GroupScript -CsvPath $csv -Action Add -WhatIf -LogDirectory $log )
        $results.Count | Should -Be 1
        $results[0].Result | Should -Be 'WhatIf'
        $results[0].Target | Should -Match '11111111-2222-3333-4444-555555555555'
        Test-Path -LiteralPath $log | Should -BeFalse
    }

    It 'rejects an invalid onboarding CSV before any preview' {
        $csv = Join-Path $TestDrive 'bad.csv'
        @(
            'DisplayName,UserPrincipalName,Department,JobTitle,LicenseSku,GroupIds'
            'Alex Example,not-an-email,ICT,Support Officer,SPE_E3,'
        ) | Set-Content -LiteralPath $csv -Encoding utf8
        $log = Join-Path $TestDrive 'bad-logs'
        { & $script:OnboardScript -CsvPath $csv -WhatIf -LogDirectory $log } | Should -Throw '*CSV validation failed*'
        Test-Path -LiteralPath $log | Should -BeFalse
    }

    It 'rejects a single offboard UPN that is not valid' {
        { & $script:DisableScript -UserPrincipalName 'not-an-email' -WhatIf } | Should -Throw '*not a valid UPN*'
    }
}

Describe 'Export safety and help' {
    It 'refuses to write an audit file into samples' {
        $blocked = Join-Path $script:RepoRoot 'samples/blocked-audit.csv'
        { & $script:ExportScript -CsvPath $script:OffboardSample -OutputPath $blocked } | Should -Throw '*Refusing to write*'
    }

    It 'refuses to replace an existing audit file without Force' {
        $existing = Join-Path $TestDrive 'audit.csv'
        'already' | Set-Content -LiteralPath $existing -Encoding utf8
        { & $script:ExportScript -CsvPath $script:OffboardSample -OutputPath $existing } | Should -Throw '*already exists*'
    }

    It 'has comment help on each entry script' {
        foreach ($path in @(
            $script:OnboardScript
            $script:DisableScript
            $script:LicenseScript
            $script:GroupScript
            $script:ExportScript
            $script:ConnectScript
        )) {
            $help = Get-Help -Name $path
            $help.Synopsis | Should -Not -BeNullOrEmpty
            ($help.Examples | Out-String) | Should -Not -BeNullOrEmpty
        }
    }
}

Describe 'Repository content stays fictional' {
    It 'does not contain employer names, private keys, or non-contoso sample addresses' {
        $roots = @(
            (Join-Path $script:RepoRoot 'src')
            (Join-Path $script:RepoRoot 'samples')
            (Join-Path $script:RepoRoot 'config')
            (Join-Path $script:RepoRoot 'docs')
        )
        $files = [System.Collections.Generic.List[string]]::new()
        foreach ($root in $roots) {
            Get-ChildItem -LiteralPath $root -File -Recurse | ForEach-Object { [void]$files.Add($_.FullName) }
        }

        foreach ($extra in @('README.md', 'requirements.md', 'LICENSE', '.env.example', '.gitignore')) {
            [void]$files.Add((Join-Path $script:RepoRoot $extra))
        }

            $banned = @('Lutheran', 'SecureStack', 'Tech Domain', 'BEGIN PRIVATE KEY', 'BEGIN CERTIFICATE')
            foreach ($file in $files) {
                $text = Get-Content -LiteralPath $file -Raw
                foreach ($term in $banned) {
                    $text.Contains($term) | Should -BeFalse
                }

                $addresses = [regex]::Matches($text, '(?i)[a-z0-9._+\-]+@[a-z0-9.-]+\.[a-z]{2,}')
                foreach ($match in $addresses) {
                $match.Value | Should -BeLike '*@contoso.com'
            }
        }
    }

    It 'states in the README that the scripts have not been used in production' {
        $readme = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'README.md') -Raw
        $readme | Should -Match 'has not been used in production'
    }
}
