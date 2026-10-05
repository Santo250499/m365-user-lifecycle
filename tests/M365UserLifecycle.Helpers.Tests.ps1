#Requires -Version 7.0

BeforeAll {
    Set-StrictMode -Version Latest
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    Import-Module -Name (Join-Path -Path $script:RepoRoot -ChildPath 'src/M365UserLifecycle.Helpers.psm1') -Force
    $script:OnboardSample = Join-Path -Path $script:RepoRoot -ChildPath 'samples/users-onboard.sample.csv'
    $script:OffboardSample = Join-Path -Path $script:RepoRoot -ChildPath 'samples/users-offboard.sample.csv'
    $script:SettingsExample = Join-Path -Path $script:RepoRoot -ChildPath 'config/settings.example.psd1'
}

Describe 'UPN and mail nickname' {
    It 'accepts a contoso user principal name' {
        Test-M365UserPrincipalName -UserPrincipalName 'alex.example@contoso.com' | Should -BeTrue
    }

    It 'rejects an empty user principal name' {
        Test-M365UserPrincipalName -UserPrincipalName '' | Should -BeFalse
    }

    It 'rejects a value with no domain' {
        Test-M365UserPrincipalName -UserPrincipalName 'alex.example' | Should -BeFalse
    }

    It 'rejects spaces' {
        Test-M365UserPrincipalName -UserPrincipalName 'alex example@contoso.com' | Should -BeFalse
    }

    It 'derives the mail nickname from the local part' {
        Get-M365MailNickname -UserPrincipalName 'alex.example@contoso.com' | Should -Be 'alex.example'
    }
}

Describe 'Password generation' {
    It 'rejects a password that contains the UPN local part' {
        $result = Test-M365PasswordComplexity -Password 'alex.exampleAa1!' -UserPrincipalName 'alex.example@contoso.com' -DisplayName 'Alex Example'
        $result.IsValid | Should -BeFalse
        $result.Reasons -join ' ' | Should -Match 'local part'
    }

    It 'rejects a short password' {
        $result = Test-M365PasswordComplexity -Password 'Aa1!' -UserPrincipalName 'alex.example@contoso.com'
        $result.IsValid | Should -BeFalse
    }

    It 'rejects a password with too few character classes' {
        $result = Test-M365PasswordComplexity -Password 'abcdefghijk' -UserPrincipalName 'alex.example@contoso.com'
        $result.IsValid | Should -BeFalse
    }

    It 'accepts a lab fixture password that meets the rules' {
        $result = Test-M365PasswordComplexity -Password 'Contoso-Lab-Temp-1!' -UserPrincipalName 'alex.example@contoso.com' -DisplayName 'Alex Example'
        $result.IsValid | Should -BeTrue
    }

    It 'generates passwords that satisfy the complexity rules' {
        1..10 | ForEach-Object {
            $password = Get-M365TemporaryPassword -UserPrincipalName 'alex.example@contoso.com' -DisplayName 'Alex Example' -Length 20
            $password.Length | Should -Be 20
            $check = Test-M365PasswordComplexity -Password $password -UserPrincipalName 'alex.example@contoso.com' -DisplayName 'Alex Example'
            $check.IsValid | Should -BeTrue
        }
    }
}

Describe 'Name parts and lists' {
    It 'splits a display name into given name and surname' {
        $parts = Get-M365PersonName -DisplayName 'Riley Placeholder'
        $parts.GivenName | Should -Be 'Riley'
        $parts.Surname | Should -Be 'Placeholder'
    }

    It 'keeps an explicit given name' {
        $parts = Get-M365PersonName -DisplayName 'Riley Placeholder' -GivenName 'R.'
        $parts.GivenName | Should -Be 'R.'
        $parts.Surname | Should -Be 'Placeholder'
    }

    It 'splits group cells on commas, semicolons, and pipes' {
        $values = @(Split-M365ListValue -Value ' alpha; beta, gamma | delta ')
        $values -join ',' | Should -Be 'alpha,beta,gamma,delta'
    }

    It 'parses boolean cells and defaults blank values' {
        (ConvertTo-M365Boolean -Value '' -Default $true).Value | Should -BeTrue
        (ConvertTo-M365Boolean -Value 'no' -Default $true).Value | Should -BeFalse
        (ConvertTo-M365Boolean -Value 'maybe' -Default $true).Ok | Should -BeFalse
    }
}

Describe 'Onboarding CSV validation' {
    It 'accepts the fictional sample file' {
        $result = Test-M365OnboardCsv -Path $script:OnboardSample -UsageLocation 'AU'
        $result.IsValid | Should -BeTrue
        @($result.Rows).Count | Should -Be 3
        @($result.Rows)[2].ManagerUserPrincipalName | Should -Be 'alex.example@contoso.com'
    }

    It 'matches the documented sample header' {
        $header = (Get-Content -LiteralPath $script:OnboardSample -TotalCount 1)
        $header | Should -Be ((Get-M365OnboardColumn) -join ',')
    }

    It 'accepts the minimum column set and derives names' {
        $path = Join-Path $TestDrive 'minimal.csv'
        @(
            'DisplayName,UserPrincipalName,Department,JobTitle,LicenseSku,GroupIds'
            'Alex Example,alex.example@contoso.com,,,SPE_E3,'
        ) | Set-Content -LiteralPath $path -Encoding utf8

        $result = Test-M365OnboardCsv -Path $path -UsageLocation 'AU'
        $result.IsValid | Should -BeTrue
        @($result.Rows)[0].GivenName | Should -Be 'Alex'
        @($result.Rows)[0].Surname | Should -Be 'Example'
    }

    It 'rejects a duplicate UPN regardless of case' {
        $path = Join-Path $TestDrive 'duplicate.csv'
        @(
            'DisplayName,UserPrincipalName,Department,JobTitle,LicenseSku,GroupIds'
            'Alex Example,alex.example@contoso.com,ICT,Support Officer,SPE_E3,'
            'Alex Example,Alex.Example@contoso.com,ICT,Support Officer,SPE_E3,'
        ) | Set-Content -LiteralPath $path -Encoding utf8

        $result = Test-M365OnboardCsv -Path $path
        $result.IsValid | Should -BeFalse
        $result.Errors -join ' ' | Should -Match 'duplicated'
    }

    It 'rejects a license sku with a space' {
        $path = Join-Path $TestDrive 'sku.csv'
        @(
            'DisplayName,UserPrincipalName,Department,JobTitle,LicenseSku,GroupIds'
            'Alex Example,alex.example@contoso.com,ICT,Support Officer,SPE E3,'
        ) | Set-Content -LiteralPath $path -Encoding utf8

        $result = Test-M365OnboardCsv -Path $path
        $result.IsValid | Should -BeFalse
        $result.Errors -join ' ' | Should -Match 'LicenseSku'
    }

    It 'rejects the all-zero group id placeholder' {
        $path = Join-Path $TestDrive 'group.csv'
        @(
            'DisplayName,UserPrincipalName,Department,JobTitle,LicenseSku,GroupIds'
            'Alex Example,alex.example@contoso.com,ICT,Support Officer,SPE_E3,00000000-0000-0000-0000-000000000000'
        ) | Set-Content -LiteralPath $path -Encoding utf8

        $result = Test-M365OnboardCsv -Path $path
        $result.IsValid | Should -BeFalse
        $result.Errors -join ' ' | Should -Match 'placeholder'
    }

    It 'rejects a manager that is the same user' {
        $path = Join-Path $TestDrive 'manager.csv'
        @(
            'DisplayName,UserPrincipalName,Department,JobTitle,LicenseSku,GroupIds,ManagerUserPrincipalName'
            'Alex Example,alex.example@contoso.com,ICT,Support Officer,SPE_E3,,alex.example@contoso.com'
        ) | Set-Content -LiteralPath $path -Encoding utf8

        $result = Test-M365OnboardCsv -Path $path
        $result.IsValid | Should -BeFalse
        $result.Errors -join ' ' | Should -Match 'ManagerUserPrincipalName'
    }

    It 'rejects a header-only file' {
        $path = Join-Path $TestDrive 'empty.csv'
        'DisplayName,UserPrincipalName,Department,JobTitle,LicenseSku,GroupIds' | Set-Content -LiteralPath $path -Encoding utf8
        $result = Test-M365OnboardCsv -Path $path
        $result.IsValid | Should -BeFalse
        $result.Errors -join ' ' | Should -Match 'no user rows'
    }
}

Describe 'Onboarding plan and user body' {
    It 'plans create, license, and manager actions for the sample' {
        $validation = Test-M365OnboardCsv -Path $script:OnboardSample
        $riley = @($validation.Rows) | Where-Object { $_.UserPrincipalName -eq 'riley.placeholder@contoso.com' }
        $plan = @(Get-M365OnboardPlan -Row $riley -UsageLocation 'AU')
        $plan.Name -join ',' | Should -Be 'CreateUser,AssignLicense,SetManager'
        $plan[1].SkuPartNumber | Should -Be 'SPE_E3'
    }

    It 'builds a New-MgUser body without the license or the group list' {
        $validation = Test-M365OnboardCsv -Path $script:OnboardSample
        $row = @($validation.Rows)[0]
        $body = ConvertTo-M365OnboardUserBody -Row $row -Password 'Contoso-Lab-Temp-1!' -UsageLocation 'AU'
        $body.accountEnabled | Should -BeTrue
        $body.mailNickname | Should -Be 'alex.example'
        $body.usageLocation | Should -Be 'AU'
        $body.passwordProfile.forceChangePasswordNextSignIn | Should -BeTrue
        $body.passwordProfile.password | Should -Be 'Contoso-Lab-Temp-1!'
        $body.Keys | Should -Not -Contain 'licenseSku'
        $body.Keys | Should -Not -Contain 'groupIds'
    }
}

Describe 'Offboarding plan' {
    It 'accepts the fictional offboard sample' {
        $result = Test-M365OffboardCsv -Path $script:OffboardSample
        $result.IsValid | Should -BeTrue
        @($result.Rows).Count | Should -Be 3
    }

    It 'keeps licenses when the mailbox note asks for the mail to be kept' {
        $result = Test-M365OffboardCsv -Path $script:OffboardSample
        $alex = @($result.Rows) | Where-Object { $_.UserPrincipalName -eq 'alex.example@contoso.com' }
        $plan = @(Get-M365OffboardPlan -User $alex)
        $plan.Name -join ',' | Should -Be 'DisableSignIn,RevokeSessions,SharedMailboxNote,RemoveGroups'
        $note = Get-M365SharedMailboxNote -UserPrincipalName $alex.UserPrincipalName
        $note | Should -Match 'does not convert'
        $note | Should -Match 'Set-Mailbox'
    }

    It 'warns when a row both records the mailbox note and removes licenses' {
        $user = Get-M365NormalizedOffboardUser -UserPrincipalName 'jamie.sample@contoso.com' -AddSharedMailboxNote:$true
        $plan = @(Get-M365OffboardPlan -User $user)
        $license = $plan | Where-Object { $_.Name -eq 'RemoveLicenses' }
        $license.Detail | Should -Match 'Warning'
    }

    It 'rejects a row with every action turned off' {
        $path = Join-Path $TestDrive 'none.csv'
        @(
            'UserPrincipalName,DisableSignIn,RevokeSessions,RemoveLicenses,RemoveGroups,AddSharedMailboxNote'
            'alex.example@contoso.com,false,false,false,false,false'
        ) | Set-Content -LiteralPath $path -Encoding utf8

        $result = Test-M365OffboardCsv -Path $path
        $result.IsValid | Should -BeFalse
        $result.Errors -join ' ' | Should -Match 'No offboarding actions'
    }
}

Describe 'Settings, logging, and safety helpers' {
    It 'loads the example settings and treats the tenant id as a placeholder' {
        $settings = Import-M365SettingsFile -Path $script:SettingsExample
        Test-M365PlaceholderIdentifier -Value $settings.TenantId | Should -BeTrue
        $settings.Scopes | Should -Contain 'User.ReadWrite.All'
        (Get-Content -LiteralPath $script:SettingsExample -Raw) | Should -Not -Match 'ClientSecret'
    }

    It 'rejects certificate mode while the sample ids are still in place' {
        $settings = Import-M365SettingsFile -Path $script:SettingsExample
        $resolved = Resolve-M365ConnectSetting -Settings $settings -AuthMode 'Certificate'
        $check = Test-M365ConnectReadiness -Resolved $resolved
        $check.IsValid | Should -BeFalse
        $check.Problems -join ' ' | Should -Match 'placeholder'
    }

    It 'does not copy a client secret into the resolved settings object' {
        $settings = Import-M365SettingsFile -Path $script:SettingsExample
        $resolved = Resolve-M365ConnectSetting -Settings $settings -AuthMode 'ClientSecret' -Environment @{
            M365_TENANT_ID     = '11111111-2222-3333-4444-555555555555'
            M365_CLIENT_ID     = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
            M365_CLIENT_SECRET = 'not-a-real-secret'
        }
        $resolved.PSObject.Properties.Name | Should -Not -Contain 'ClientSecret'
        $resolved.HasClientSecret | Should -BeTrue
        $check = Test-M365ConnectReadiness -Resolved $resolved
        $check.IsValid | Should -BeTrue
    }

    It 'resolves the default log directory under the repository root' {
        $resolved = Resolve-M365LogDirectory -RepoRoot $script:RepoRoot
        $resolved | Should -Be (Join-Path -Path $script:RepoRoot -ChildPath 'logs')
    }

    It 'writes a log line without the temporary password' {
        $logDirectory = Join-Path $TestDrive 'logs'
        $result = Get-M365ActionResult -UserPrincipalName 'alex.example@contoso.com' -Operation 'Create Entra ID user' -Target 'alex.example@contoso.com' -Result Applied -Detail 'Created.' -TemporaryPassword 'TEMP-PASSWORD-SHOULD-NOT-BE-LOGGED'
        $path = Write-M365ActionLog -ActionResult $result -LogDirectory $logDirectory
        $content = Get-Content -LiteralPath $path -Raw
        $content | Should -Match 'Create Entra ID user'
        $content | Should -Not -Match 'TEMP-PASSWORD-SHOULD-NOT-BE-LOGGED'
    }

    It 'recognises a Graph 404 message' {
        $record = [System.Management.Automation.ErrorRecord]::new(
            [System.Exception]::new('Status: 404 (NotFound) ErrorCode: Request_ResourceNotFound'),
            'NotFound',
            [System.Management.Automation.ErrorCategory]::ObjectNotFound,
            $null
        )
        Test-M365GraphNotFoundError -ErrorRecord $record | Should -BeTrue
        Test-M365GraphNotFoundError -ErrorRecord $null | Should -BeFalse
    }
}
