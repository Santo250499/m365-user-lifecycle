#Requires -Version 7.0

Set-StrictMode -Version Latest

<#
.SYNOPSIS
    Pure helpers for the M365 user lifecycle scripts.

.DESCRIPTION
    CSV checks, UPN and password generation, action plans, settings checks, and
    log formatting. This module does not connect to Microsoft Graph and does not
    read secrets. Entry scripts in this folder import it, including during -WhatIf.
#>

$script:OnboardRequiredColumns = @(
    'DisplayName'
    'UserPrincipalName'
    'Department'
    'JobTitle'
    'LicenseSku'
    'GroupIds'
)

$script:OffboardColumns = @(
    'UserPrincipalName'
    'DisableSignIn'
    'RevokeSessions'
    'RemoveLicenses'
    'RemoveGroups'
    'AddSharedMailboxNote'
)

function Get-M365OnboardColumn {
    [CmdletBinding()]
    param()

    return [string[]]@(
        'DisplayName'
        'UserPrincipalName'
        'GivenName'
        'Surname'
        'Department'
        'JobTitle'
        'LicenseSku'
        'GroupIds'
        'ManagerUserPrincipalName'
        'UsageLocation'
    )
}

function Get-M365OffboardColumn {
    [CmdletBinding()]
    param()

    return [string[]]$script:OffboardColumns
}

function Get-M365DefaultGraphScope {
    <#
    .SYNOPSIS
        Returns the Graph scopes the lifecycle scripts are written for.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Lifecycle', 'Audit')]
        [string]$ScopeProfile = 'Lifecycle'
    )

    if ($ScopeProfile -eq 'Audit') {
        return [string[]]@(
            'User.Read.All'
            'GroupMember.Read.All'
            'Organization.Read.All'
        )
    }

    return [string[]]@(
        'User.ReadWrite.All'
        'GroupMember.ReadWrite.All'
        'Organization.Read.All'
        'LicenseAssignment.ReadWrite.All'
        'User.RevokeSessions.All'
    )
}

function Test-M365UserPrincipalName {
    <#
    .SYNOPSIS
        Returns true when a string is a usable Entra user principal name.
    #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$UserPrincipalName = ''
    )

    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        return $false
    }

    $value = $UserPrincipalName.Trim()
    if ($value.Length -gt 113) {
        return $false
    }

    return $value -match '^[A-Za-z0-9._+\-]+@[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?)+$'
}

function Get-M365MailNickname {
    <#
    .SYNOPSIS
        Derives a mail nickname from the local part of a UPN.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    if (-not (Test-M365UserPrincipalName -UserPrincipalName $UserPrincipalName)) {
        throw "UserPrincipalName '$UserPrincipalName' is not valid."
    }

    $local = ($UserPrincipalName.Trim() -split '@', 2)[0]
    if ($local.Length -gt 64) {
        throw "Mail nickname derived from '$UserPrincipalName' exceeds 64 characters."
    }

    return $local
}

function Test-M365Guid {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value = ''
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $parsed = [guid]::Empty
    return [guid]::TryParse($Value.Trim(), [ref]$parsed)
}

function Test-M365PlaceholderIdentifier {
    <#
    .SYNOPSIS
        Returns true for blank values and the all-zero sample GUID.
    #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value = ''
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $true
    }

    return $Value.Trim() -eq '00000000-0000-0000-0000-000000000000'
}

function Split-M365ListValue {
    <#
    .SYNOPSIS
        Splits a CSV cell on commas, semicolons, or pipes.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return [string[]]@()
    }

    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($part in ($Value -split '[;,|]')) {
        $trimmed = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
            [void]$parts.Add($trimmed)
        }
    }

    return [string[]]$parts.ToArray()
}

function ConvertTo-M365Boolean {
    <#
    .SYNOPSIS
        Parses true/false, yes/no, and 1/0. Blank values use the default.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value,

        [Parameter(Mandatory)]
        [bool]$Default
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return [pscustomobject]@{
            Ok    = $true
            Value = $Default
            Error = ''
        }
    }

    $token = $Value.Trim().ToLowerInvariant()
    if ($token -in @('true', 'yes', '1')) {
        return [pscustomobject]@{
            Ok    = $true
            Value = $true
            Error = ''
        }
    }

    if ($token -in @('false', 'no', '0')) {
        return [pscustomobject]@{
            Ok    = $true
            Value = $false
            Error = ''
        }
    }

    return [pscustomobject]@{
        Ok    = $false
        Value = $Default
        Error = "Cannot interpret '$Value' as a boolean. Use true or false."
    }
}

function Test-M365PasswordComplexity {
    <#
    .SYNOPSIS
        Checks a password against the Entra complexity rules this repo enforces.
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingPlainTextForPassword',
        '',
        Justification = 'Entra password complexity checks have to inspect the characters. This value is not a stored credential.')]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingUsernameAndPasswordParams',
        '',
        Justification = 'UserPrincipalName is the account the password must not contain. It is not a sign-in credential pair.')]
    param(
        [AllowEmptyString()]
        [string]$Password = '',

        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [AllowEmptyString()]
        [string]$DisplayName = ''
    )

    $reasons = [System.Collections.Generic.List[string]]::new()
    if ($Password.Length -lt 8 -or $Password.Length -gt 256) {
        [void]$reasons.Add('Length must be between 8 and 256 characters.')
    }

    $classes = 0
    if ($Password -cmatch '[a-z]') { $classes++ }
    if ($Password -cmatch '[A-Z]') { $classes++ }
    if ($Password -match '[0-9]') { $classes++ }
    if ($Password -match '[^A-Za-z0-9]') { $classes++ }
    if ($classes -lt 3) {
        [void]$reasons.Add('Password must include at least three of: lowercase, uppercase, digit, symbol.')
    }

    $local = ($UserPrincipalName -split '@', 2)[0]
    if (-not [string]::IsNullOrWhiteSpace($local) -and $local.Length -ge 3) {
        if ($Password.ToLowerInvariant().Contains($local.ToLowerInvariant())) {
            [void]$reasons.Add('Password must not contain the user principal name local part.')
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($DisplayName)) {
        foreach ($part in ($DisplayName -split '\s+')) {
            if ($part.Length -ge 3 -and $Password.ToLowerInvariant().Contains($part.ToLowerInvariant())) {
                [void]$reasons.Add("Password must not contain the name part '$part'.")
            }
        }
    }

    return [pscustomobject]@{
        IsValid = ($reasons.Count -eq 0)
        Reasons = [string[]]$reasons.ToArray()
    }
}

function Get-M365RandomIndex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Random,

        [Parameter(Mandatory)]
        [int]$UpperExclusive
    )

    if ($UpperExclusive -le 0) {
        throw 'UpperExclusive must be greater than zero.'
    }

    $buffer = [byte[]]::new(4)
    $Random.GetBytes($buffer)
    $value = [System.BitConverter]::ToUInt32($buffer, 0)
    return [int]($value % [uint32]$UpperExclusive)
}

function Get-M365RandomCharacter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Alphabet,

        [Parameter(Mandatory)]
        $Random
    )

    $index = Get-M365RandomIndex -Random $Random -UpperExclusive $Alphabet.Length
    return $Alphabet[$index]
}

function Get-M365ShuffledString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [char[]]$Characters,

        [Parameter(Mandatory)]
        $Random
    )

    $items = [char[]]::new($Characters.Length)
    [Array]::Copy($Characters, $items, $Characters.Length)
    for ($index = $items.Length - 1; $index -gt 0; $index--) {
        $swap = Get-M365RandomIndex -Random $Random -UpperExclusive ($index + 1)
        $temporary = $items[$index]
        $items[$index] = $items[$swap]
        $items[$swap] = $temporary
    }

    return -join $items
}

function Get-M365TemporaryPassword {
    <#
    .SYNOPSIS
        Builds a random temporary password that satisfies Test-M365PasswordComplexity.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [AllowEmptyString()]
        [string]$DisplayName = '',

        [ValidateRange(12, 128)]
        [int]$Length = 20
    )

    if (-not (Test-M365UserPrincipalName -UserPrincipalName $UserPrincipalName)) {
        throw "UserPrincipalName '$UserPrincipalName' is not valid."
    }

    $sets = @(
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        'abcdefghijklmnopqrstuvwxyz'
        '0123456789'
        '!@#$%^&*-_=+'
    )
    $alphabet = -join $sets
    $random = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        for ($attempt = 0; $attempt -lt 25; $attempt++) {
            $chars = [System.Collections.Generic.List[char]]::new()
            foreach ($set in $sets) {
                [void]$chars.Add((Get-M365RandomCharacter -Alphabet $set -Random $random))
            }

            while ($chars.Count -lt $Length) {
                [void]$chars.Add((Get-M365RandomCharacter -Alphabet $alphabet -Random $random))
            }

            $password = Get-M365ShuffledString -Characters $chars.ToArray() -Random $random
            $check = Test-M365PasswordComplexity -Password $password -UserPrincipalName $UserPrincipalName -DisplayName $DisplayName
            if ($check.IsValid) {
                return $password
            }
        }
    }
    finally {
        $random.Dispose()
    }

    throw "Could not generate a password that satisfies complexity rules for $UserPrincipalName."
}

function Get-M365PersonName {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$DisplayName = '',

        [AllowEmptyString()]
        [string]$GivenName = '',

        [AllowEmptyString()]
        [string]$Surname = ''
    )

    $given = $GivenName.Trim()
    $surname = $Surname.Trim()
    if (-not [string]::IsNullOrWhiteSpace($DisplayName)) {
        $tokens = [System.Collections.Generic.List[string]]::new()
        foreach ($token in ($DisplayName.Trim() -split '\s+')) {
            if (-not [string]::IsNullOrWhiteSpace($token)) {
                [void]$tokens.Add($token)
            }
        }

        if ([string]::IsNullOrWhiteSpace($given) -and $tokens.Count -ge 1) {
            $given = $tokens[0]
        }

        if ([string]::IsNullOrWhiteSpace($surname) -and $tokens.Count -ge 2) {
            $surname = ($tokens[1..($tokens.Count - 1)] -join ' ')
        }
    }

    return [pscustomobject]@{
        GivenName = $given
        Surname   = $surname
    }
}

function Resolve-M365UsageLocation {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$RowValue = '',

        [Parameter(Mandatory)]
        [string]$Default
    )

    $selected = $Default
    if (-not [string]::IsNullOrWhiteSpace($RowValue)) {
        $selected = $RowValue.Trim()
    }

    $selected = $selected.ToUpperInvariant()
    if ($selected -notmatch '^[A-Z]{2}$') {
        throw "UsageLocation '$selected' must be a two-letter ISO country code."
    }

    return $selected
}

function Get-M365CsvField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Row,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $property = $Row.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return ''
    }

    return ([string]$property.Value).Trim()
}

function Import-M365CsvDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "CSV file not found: $Path"
    }

    $headerLine = Get-Content -LiteralPath $Path -TotalCount 1
    if ([string]::IsNullOrWhiteSpace($headerLine)) {
        throw "CSV file is empty: $Path"
    }

    $imported = Import-Csv -LiteralPath $Path
    $rows = @()
    if ($null -ne $imported) {
        $rows = @($imported)
    }

    $headers = @()
    if ($rows.Count -gt 0) {
        $headers = @($rows[0].PSObject.Properties.Name)
    }
    else {
        foreach ($header in ($headerLine.Split(','))) {
            $trimmed = $header.Trim().Trim('"')
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                $headers += $trimmed
            }
        }
    }

    return [pscustomobject]@{
        Headers = [string[]]$headers
        Rows    = @($rows)
    }
}

function Get-M365MissingColumn {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()]
        [string[]]$Header = @(),

        [Parameter(Mandatory)]
        [string[]]$Required
    )

    $missing = [System.Collections.Generic.List[string]]::new()
    foreach ($column in $Required) {
        $found = $false
        foreach ($name in $Header) {
            if ($name -eq $column) {
                $found = $true
                break
            }
        }

        if (-not $found) {
            [void]$missing.Add($column)
        }
    }

    return [string[]]$missing.ToArray()
}

function Test-M365BoundedText {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value = '',

        [Parameter(Mandatory)]
        [string]$Label,

        [Parameter(Mandatory)]
        [int]$MaximumLength,

        [switch]$Required
    )

    $problems = [System.Collections.Generic.List[string]]::new()
    if ($Required -and [string]::IsNullOrWhiteSpace($Value)) {
        [void]$problems.Add("$Label is required.")
        return [string[]]$problems.ToArray()
    }

    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        if ($Value -match '\p{C}') {
            [void]$problems.Add("$Label contains a control character.")
        }

        if ($Value.Length -gt $MaximumLength) {
            [void]$problems.Add("$Label must be $MaximumLength characters or fewer.")
        }
    }

    return [string[]]$problems.ToArray()
}

function Get-M365ValidationResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.IList]$Errors,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.IList]$Rows
    )

    return [pscustomobject]@{
        IsValid = ($Errors.Count -eq 0)
        Errors  = [string[]]$Errors.ToArray()
        Rows    = @($Rows.ToArray())
    }
}

function Test-M365OnboardCsv {
    <#
    .SYNOPSIS
        Validates an onboarding CSV and returns normalized rows when it is valid.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [string]$UsageLocation = 'AU'
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    $users = [System.Collections.Generic.List[object]]::new()
    $document = Import-M365CsvDocument -Path $Path
    foreach ($missing in @(Get-M365MissingColumn -Header $document.Headers -Required $script:OnboardRequiredColumns)) {
        [void]$errors.Add("CSV is missing required column '$missing'.")
    }

    if ($errors.Count -gt 0) {
        return Get-M365ValidationResult -Errors $errors -Rows $users
    }

    if ($document.Rows.Count -eq 0) {
        [void]$errors.Add('CSV contains no user rows.')
        return Get-M365ValidationResult -Errors $errors -Rows $users
    }

    try {
        $null = Resolve-M365UsageLocation -Default $UsageLocation
    }
    catch {
        [void]$errors.Add($_.Exception.Message)
        return Get-M365ValidationResult -Errors $errors -Rows $users
    }

    $seen = @{}
    $index = 0
    foreach ($raw in $document.Rows) {
        $index++
        $rowNumber = $index + 1
        $rowErrors = [System.Collections.Generic.List[string]]::new()
        $displayName = Get-M365CsvField -Row $raw -Name 'DisplayName'
        $upn = Get-M365CsvField -Row $raw -Name 'UserPrincipalName'
        $department = Get-M365CsvField -Row $raw -Name 'Department'
        $jobTitle = Get-M365CsvField -Row $raw -Name 'JobTitle'
        $licenseSku = Get-M365CsvField -Row $raw -Name 'LicenseSku'
        $groupCell = Get-M365CsvField -Row $raw -Name 'GroupIds'
        $givenName = Get-M365CsvField -Row $raw -Name 'GivenName'
        $surname = Get-M365CsvField -Row $raw -Name 'Surname'
        $manager = Get-M365CsvField -Row $raw -Name 'ManagerUserPrincipalName'
        $rowLocation = Get-M365CsvField -Row $raw -Name 'UsageLocation'

        foreach ($problem in @(Test-M365BoundedText -Value $displayName -Label 'DisplayName' -MaximumLength 256 -Required)) {
            [void]$rowErrors.Add($problem)
        }

        if (-not (Test-M365UserPrincipalName -UserPrincipalName $upn)) {
            [void]$rowErrors.Add("UserPrincipalName '$upn' is not a valid UPN.")
        }
        elseif ($seen.ContainsKey($upn.ToLowerInvariant())) {
            [void]$rowErrors.Add("UserPrincipalName '$upn' is duplicated (first seen on row $($seen[$upn.ToLowerInvariant()])).")
        }
        else {
            $seen[$upn.ToLowerInvariant()] = $rowNumber
        }

        foreach ($problem in @(Test-M365BoundedText -Value $department -Label 'Department' -MaximumLength 64)) {
            [void]$rowErrors.Add($problem)
        }

        foreach ($problem in @(Test-M365BoundedText -Value $jobTitle -Label 'JobTitle' -MaximumLength 128)) {
            [void]$rowErrors.Add($problem)
        }

        foreach ($problem in @(Test-M365BoundedText -Value $givenName -Label 'GivenName' -MaximumLength 64)) {
            [void]$rowErrors.Add($problem)
        }

        foreach ($problem in @(Test-M365BoundedText -Value $surname -Label 'Surname' -MaximumLength 64)) {
            [void]$rowErrors.Add($problem)
        }

        if (-not [string]::IsNullOrWhiteSpace($licenseSku) -and $licenseSku -notmatch '^[A-Za-z0-9_\-]{1,64}$') {
            [void]$rowErrors.Add("LicenseSku '$licenseSku' must contain only letters, numbers, underscores, or hyphens.")
        }

        $groupIds = [System.Collections.Generic.List[string]]::new()
        foreach ($groupId in @(Split-M365ListValue -Value $groupCell)) {
            if (-not (Test-M365Guid -Value $groupId)) {
                [void]$rowErrors.Add("Group id '$groupId' is not a GUID.")
                continue
            }

            if (Test-M365PlaceholderIdentifier -Value $groupId) {
                [void]$rowErrors.Add("Group id '$groupId' is the sample placeholder. Replace it with a lab group object id or leave the cell blank.")
                continue
            }

            $already = $false
            foreach ($existing in $groupIds) {
                if ($existing -eq $groupId) {
                    $already = $true
                    break
                }
            }

            if (-not $already) {
                [void]$groupIds.Add($groupId)
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($manager)) {
            if (-not (Test-M365UserPrincipalName -UserPrincipalName $manager)) {
                [void]$rowErrors.Add("ManagerUserPrincipalName '$manager' is not a valid UPN.")
            }
            elseif ($manager.Equals($upn, [System.StringComparison]::OrdinalIgnoreCase)) {
                [void]$rowErrors.Add('ManagerUserPrincipalName must not be the same as UserPrincipalName.')
            }
        }

        $resolvedLocation = ''
        try {
            $resolvedLocation = Resolve-M365UsageLocation -RowValue $rowLocation -Default $UsageLocation
        }
        catch {
            [void]$rowErrors.Add($_.Exception.Message)
        }

        foreach ($problem in $rowErrors) {
            [void]$errors.Add("Row ${rowNumber}: $problem")
        }

        if ($rowErrors.Count -gt 0) {
            continue
        }

        $names = Get-M365PersonName -DisplayName $displayName -GivenName $givenName -Surname $surname
        [void]$users.Add([pscustomobject]@{
            DisplayName               = $displayName
            UserPrincipalName         = $upn
            GivenName                 = $names.GivenName
            Surname                   = $names.Surname
            Department                = $department
            JobTitle                  = $jobTitle
            LicenseSku                = $licenseSku
            GroupIds                  = [string[]]$groupIds.ToArray()
            ManagerUserPrincipalName  = $manager
            UsageLocation             = $resolvedLocation
            SourceRow                 = $rowNumber
        })
    }

    return Get-M365ValidationResult -Errors $errors -Rows $users
}

function Get-M365OffboardFlag {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Row,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [bool]$Default,

        [Parameter(Mandatory)]
        $Errors
    )

    $parsed = ConvertTo-M365Boolean -Value (Get-M365CsvField -Row $Row -Name $Name) -Default $Default
    if (-not $parsed.Ok) {
        [void]$Errors.Add($parsed.Error)
    }

    return [bool]$parsed.Value
}

function Get-M365NormalizedOffboardUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [bool]$DisableSignIn = $true,
        [bool]$RevokeSessions = $true,
        [bool]$RemoveLicenses = $true,
        [bool]$RemoveGroups = $true,
        [bool]$AddSharedMailboxNote = $false,
        [int]$SourceRow = 0
    )

    return [pscustomobject]@{
        UserPrincipalName    = $UserPrincipalName.Trim()
        DisableSignIn        = $DisableSignIn
        RevokeSessions       = $RevokeSessions
        RemoveLicenses       = $RemoveLicenses
        RemoveGroups         = $RemoveGroups
        AddSharedMailboxNote = $AddSharedMailboxNote
        SourceRow            = $SourceRow
    }
}

function Test-M365OffboardCsv {
    <#
    .SYNOPSIS
        Validates an offboarding CSV. Blank flag cells use the supplied defaults.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [bool]$DisableSignIn = $true,
        [bool]$RevokeSessions = $true,
        [bool]$RemoveLicenses = $true,
        [bool]$RemoveGroups = $true,
        [bool]$AddSharedMailboxNote = $false
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    $users = [System.Collections.Generic.List[object]]::new()
    $document = Import-M365CsvDocument -Path $Path
    foreach ($missing in @(Get-M365MissingColumn -Header $document.Headers -Required @('UserPrincipalName'))) {
        [void]$errors.Add("CSV is missing required column '$missing'.")
    }

    if ($errors.Count -gt 0) {
        return Get-M365ValidationResult -Errors $errors -Rows $users
    }

    if ($document.Rows.Count -eq 0) {
        [void]$errors.Add('CSV contains no user rows.')
        return Get-M365ValidationResult -Errors $errors -Rows $users
    }

    $seen = @{}
    $index = 0
    foreach ($raw in $document.Rows) {
        $index++
        $rowNumber = $index + 1
        $rowErrors = [System.Collections.Generic.List[string]]::new()
        $upn = Get-M365CsvField -Row $raw -Name 'UserPrincipalName'
        if (-not (Test-M365UserPrincipalName -UserPrincipalName $upn)) {
            [void]$rowErrors.Add("UserPrincipalName '$upn' is not a valid UPN.")
        }
        elseif ($seen.ContainsKey($upn.ToLowerInvariant())) {
            [void]$rowErrors.Add("UserPrincipalName '$upn' is duplicated (first seen on row $($seen[$upn.ToLowerInvariant()])).")
        }
        else {
            $seen[$upn.ToLowerInvariant()] = $rowNumber
        }

        $disable = Get-M365OffboardFlag -Row $raw -Name 'DisableSignIn' -Default $DisableSignIn -Errors $rowErrors
        $revoke = Get-M365OffboardFlag -Row $raw -Name 'RevokeSessions' -Default $RevokeSessions -Errors $rowErrors
        $licenses = Get-M365OffboardFlag -Row $raw -Name 'RemoveLicenses' -Default $RemoveLicenses -Errors $rowErrors
        $groups = Get-M365OffboardFlag -Row $raw -Name 'RemoveGroups' -Default $RemoveGroups -Errors $rowErrors
        $mailbox = Get-M365OffboardFlag -Row $raw -Name 'AddSharedMailboxNote' -Default $AddSharedMailboxNote -Errors $rowErrors

        if ($rowErrors.Count -eq 0) {
            $normalized = Get-M365NormalizedOffboardUser -UserPrincipalName $upn -DisableSignIn:$disable -RevokeSessions:$revoke -RemoveLicenses:$licenses -RemoveGroups:$groups -AddSharedMailboxNote:$mailbox -SourceRow $rowNumber
            $plan = @(Get-M365OffboardPlan -User $normalized)
            if ($plan.Count -eq 0) {
                [void]$rowErrors.Add('No offboarding actions were selected.')
            }
            else {
                [void]$users.Add($normalized)
            }
        }

        foreach ($problem in $rowErrors) {
            [void]$errors.Add("Row ${rowNumber}: $problem")
        }
    }

    return Get-M365ValidationResult -Errors $errors -Rows $users
}

function Test-M365UserListCsv {
    <#
    .SYNOPSIS
        Validates a CSV that only needs a UserPrincipalName column.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    $users = [System.Collections.Generic.List[object]]::new()
    $document = Import-M365CsvDocument -Path $Path
    foreach ($missing in @(Get-M365MissingColumn -Header $document.Headers -Required @('UserPrincipalName'))) {
        [void]$errors.Add("CSV is missing required column '$missing'.")
    }

    if ($errors.Count -gt 0) {
        return Get-M365ValidationResult -Errors $errors -Rows $users
    }

    if ($document.Rows.Count -eq 0) {
        [void]$errors.Add('CSV contains no user rows.')
        return Get-M365ValidationResult -Errors $errors -Rows $users
    }

    $seen = @{}
    $index = 0
    foreach ($raw in $document.Rows) {
        $index++
        $rowNumber = $index + 1
        $upn = Get-M365CsvField -Row $raw -Name 'UserPrincipalName'
        if (-not (Test-M365UserPrincipalName -UserPrincipalName $upn)) {
            [void]$errors.Add("Row ${rowNumber}: UserPrincipalName '$upn' is not a valid UPN.")
            continue
        }

        if ($seen.ContainsKey($upn.ToLowerInvariant())) {
            [void]$errors.Add("Row ${rowNumber}: UserPrincipalName '$upn' is duplicated (first seen on row $($seen[$upn.ToLowerInvariant()])).")
            continue
        }

        $seen[$upn.ToLowerInvariant()] = $rowNumber
        [void]$users.Add([pscustomobject]@{
            UserPrincipalName = $upn
            SourceRow         = $rowNumber
        })
    }

    return Get-M365ValidationResult -Errors $errors -Rows $users
}

function Test-M365GroupMembershipCsv {
    <#
    .SYNOPSIS
        Validates a CSV with UserPrincipalName and GroupIds.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    $users = [System.Collections.Generic.List[object]]::new()
    $document = Import-M365CsvDocument -Path $Path
    foreach ($missing in @(Get-M365MissingColumn -Header $document.Headers -Required @('UserPrincipalName', 'GroupIds'))) {
        [void]$errors.Add("CSV is missing required column '$missing'.")
    }

    if ($errors.Count -gt 0) {
        return Get-M365ValidationResult -Errors $errors -Rows $users
    }

    if ($document.Rows.Count -eq 0) {
        [void]$errors.Add('CSV contains no user rows.')
        return Get-M365ValidationResult -Errors $errors -Rows $users
    }

    $seen = @{}
    $index = 0
    foreach ($raw in $document.Rows) {
        $index++
        $rowNumber = $index + 1
        $rowErrors = [System.Collections.Generic.List[string]]::new()
        $upn = Get-M365CsvField -Row $raw -Name 'UserPrincipalName'
        if (-not (Test-M365UserPrincipalName -UserPrincipalName $upn)) {
            [void]$rowErrors.Add("UserPrincipalName '$upn' is not a valid UPN.")
        }
        elseif ($seen.ContainsKey($upn.ToLowerInvariant())) {
            [void]$rowErrors.Add("UserPrincipalName '$upn' is duplicated (first seen on row $($seen[$upn.ToLowerInvariant()])).")
        }
        else {
            $seen[$upn.ToLowerInvariant()] = $rowNumber
        }

        $groupIds = [System.Collections.Generic.List[string]]::new()
        foreach ($groupId in @(Split-M365ListValue -Value (Get-M365CsvField -Row $raw -Name 'GroupIds'))) {
            if (-not (Test-M365Guid -Value $groupId)) {
                [void]$rowErrors.Add("Group id '$groupId' is not a GUID.")
                continue
            }

            if (Test-M365PlaceholderIdentifier -Value $groupId) {
                [void]$rowErrors.Add("Group id '$groupId' is the sample placeholder. Replace it with a lab group object id.")
                continue
            }

            [void]$groupIds.Add($groupId)
        }

        if ($groupIds.Count -eq 0) {
            [void]$rowErrors.Add('At least one group id is required.')
        }

        foreach ($problem in $rowErrors) {
            [void]$errors.Add("Row ${rowNumber}: $problem")
        }

        if ($rowErrors.Count -eq 0) {
            [void]$users.Add([pscustomobject]@{
                UserPrincipalName = $upn
                GroupIds          = [string[]]$groupIds.ToArray()
                SourceRow         = $rowNumber
            })
        }
    }

    return Get-M365ValidationResult -Errors $errors -Rows $users
}

function ConvertTo-M365OnboardUserBody {
    <#
    .SYNOPSIS
        Builds the hashtable passed to New-MgUser -BodyParameter.
    #>
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingPlainTextForPassword',
        '',
        Justification = 'New-MgUser passwordProfile.password requires a string. The password is generated at runtime and is not written to the log.')]
    param(
        [Parameter(Mandatory)]
        $Row,

        [Parameter(Mandatory)]
        [string]$Password,

        [Parameter(Mandatory)]
        [string]$UsageLocation
    )

    $complexity = Test-M365PasswordComplexity -Password $Password -UserPrincipalName $Row.UserPrincipalName -DisplayName $Row.DisplayName
    if (-not $complexity.IsValid) {
        throw ("Password does not meet complexity rules: " + ($complexity.Reasons -join ' '))
    }

    $location = Resolve-M365UsageLocation -RowValue $Row.UsageLocation -Default $UsageLocation
    $body = @{
        accountEnabled    = $true
        displayName       = $Row.DisplayName
        mailNickname      = (Get-M365MailNickname -UserPrincipalName $Row.UserPrincipalName)
        userPrincipalName = $Row.UserPrincipalName
        usageLocation     = $location
        passwordProfile   = @{
            forceChangePasswordNextSignIn = $true
            password                      = $Password
        }
    }

    $optional = @{
        givenName  = $Row.GivenName
        surname    = $Row.Surname
        jobTitle   = $Row.JobTitle
        department = $Row.Department
    }

    foreach ($key in @($optional.Keys)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$optional[$key])) {
            $body[$key] = [string]$optional[$key]
        }
    }

    return $body
}

function Get-M365OnboardPlan {
    <#
    .SYNOPSIS
        Lists the directory changes an onboarding row would make.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Row,

        [Parameter(Mandatory)]
        [string]$UsageLocation
    )

    $upn = $Row.UserPrincipalName
    $location = Resolve-M365UsageLocation -RowValue $Row.UsageLocation -Default $UsageLocation
    $actions = [System.Collections.Generic.List[object]]::new()
    [void]$actions.Add([pscustomobject]@{
        Name                     = 'CreateUser'
        Operation                = 'Create Entra ID user'
        Target                   = $upn
        Detail                   = "Create an enabled user with usage location $location and a generated temporary password. The password is returned only after a real run and is not written to the log."
        SkuPartNumber            = ''
        GroupId                  = ''
        ManagerUserPrincipalName = ''
    })

    if (-not [string]::IsNullOrWhiteSpace([string]$Row.LicenseSku)) {
        [void]$actions.Add([pscustomobject]@{
            Name                     = 'AssignLicense'
            Operation                = "Assign license $($Row.LicenseSku)"
            Target                   = $upn
            Detail                   = "Assign SkuPartNumber $($Row.LicenseSku). If the account has no usage location, set $location first."
            SkuPartNumber            = [string]$Row.LicenseSku
            GroupId                  = ''
            ManagerUserPrincipalName = ''
        })
    }

    foreach ($groupId in @($Row.GroupIds)) {
        [void]$actions.Add([pscustomobject]@{
            Name                     = 'AddGroup'
            Operation                = 'Add group member'
            Target                   = "$upn -> $groupId"
            Detail                   = "Add the user to group $groupId if they are not already a member."
            SkuPartNumber            = ''
            GroupId                  = [string]$groupId
            ManagerUserPrincipalName = ''
        })
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$Row.ManagerUserPrincipalName)) {
        [void]$actions.Add([pscustomobject]@{
            Name                     = 'SetManager'
            Operation                = 'Set manager'
            Target                   = $upn
            Detail                   = "Set the manager to $($Row.ManagerUserPrincipalName)."
            SkuPartNumber            = ''
            GroupId                  = ''
            ManagerUserPrincipalName = [string]$Row.ManagerUserPrincipalName
        })
    }

    return @($actions.ToArray())
}

function Get-M365SharedMailboxNote {
    <#
    .SYNOPSIS
        Text recorded when a leaver mailbox might need an Exchange conversion.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    return "Manual Exchange Online step for ${UserPrincipalName}. This script does not convert mailboxes. Microsoft Graph cannot set a mailbox type to shared. If the mailbox must be kept, an Exchange administrator should convert it before the license is removed: Set-Mailbox -Identity '${UserPrincipalName}' -Type Shared. Removing the license first can schedule the mailbox for deletion. Do this only in a lab tenant, and only with change control."
}

function Get-M365OffboardPlan {
    <#
    .SYNOPSIS
        Lists offboarding actions in the order the disable script performs them.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $User
    )

    $upn = $User.UserPrincipalName
    $actions = [System.Collections.Generic.List[object]]::new()
    if ($User.DisableSignIn) {
        [void]$actions.Add([pscustomobject]@{
            Name      = 'DisableSignIn'
            Operation = 'Disable sign-in'
            Target    = $upn
            Detail    = 'Set accountEnabled to false.'
        })
    }

    if ($User.RevokeSessions) {
        [void]$actions.Add([pscustomobject]@{
            Name      = 'RevokeSessions'
            Operation = 'Revoke sign-in sessions'
            Target    = $upn
            Detail    = 'Revoke refresh tokens and session cookies for the user.'
        })
    }

    if ($User.AddSharedMailboxNote) {
        [void]$actions.Add([pscustomobject]@{
            Name      = 'SharedMailboxNote'
            Operation = 'Record shared-mailbox conversion note'
            Target    = $upn
            Detail    = (Get-M365SharedMailboxNote -UserPrincipalName $upn)
        })
    }

    if ($User.RemoveLicenses) {
        $detail = 'Remove every directly assigned license. Group-based licenses remain until the user is removed from the licensing group.'
        if ($User.AddSharedMailboxNote) {
            $detail += ' Warning: this run still removes licenses. Convert the mailbox first if the mail must be kept.'
        }

        [void]$actions.Add([pscustomobject]@{
            Name      = 'RemoveLicenses'
            Operation = 'Remove direct licenses'
            Target    = $upn
            Detail    = $detail
        })
    }

    if ($User.RemoveGroups) {
        [void]$actions.Add([pscustomobject]@{
            Name      = 'RemoveGroups'
            Operation = 'Remove group memberships'
            Target    = $upn
            Detail    = 'Remove the user from security and Microsoft 365 groups. Directory roles are left unchanged. Dynamic groups may refuse the removal and are reported as failures.'
        })
    }

    return @($actions.ToArray())
}

function Get-M365ActionResult {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingPlainTextForPassword',
        '',
        Justification = 'The temporary password is returned to the operator once. The log writer deliberately ignores this property.')]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingUsernameAndPasswordParams',
        '',
        Justification = 'UserPrincipalName identifies the row. TemporaryPassword is optional output for the operator, not a credential used to sign in.')]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$Operation,

        [Parameter(Mandatory)]
        [string]$Target,

        [Parameter(Mandatory)]
        [ValidateSet('WhatIf', 'Applied', 'Skipped', 'Failed', 'Noted')]
        [string]$Result,

        [AllowEmptyString()]
        [string]$Detail = '',

        [string]$TemporaryPassword
    )

    $properties = [ordered]@{
        UserPrincipalName = $UserPrincipalName
        Operation         = $Operation
        Target            = $Target
        Result            = $Result
        Detail            = $Detail
    }

    if ($PSBoundParameters.ContainsKey('TemporaryPassword')) {
        $properties['TemporaryPassword'] = $TemporaryPassword
    }

    return [pscustomobject]$properties
}

function Format-M365LifecycleLogLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Level,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,

        [datetime]$Timestamp = ([datetime]::UtcNow)
    )

    $clean = ($Message -replace '[\r\n]+', ' ').Trim()
    return ('{0:yyyy-MM-ddTHH:mm:ssZ} [{1}] {2}' -f $Timestamp.ToUniversalTime(), $Level.ToUpperInvariant(), $clean)
}

function Write-M365LifecycleLog {
    <#
    .SYNOPSIS
        Appends one single-line record under the log directory.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'WHATIF')]
        [string]$Level = 'INFO',

        [Parameter(Mandatory)]
        [string]$LogDirectory
    )

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        $null = New-Item -ItemType Directory -Path $LogDirectory -Force
    }

    $path = Join-Path -Path $LogDirectory -ChildPath ('m365-lifecycle-{0:yyyyMMdd}.log' -f ([datetime]::UtcNow))
    $line = Format-M365LifecycleLogLine -Level $Level -Message $Message
    Add-Content -LiteralPath $path -Value $line -Encoding utf8
    Write-Information -MessageData $line -InformationAction Continue
    return $path
}

function Write-M365Result {
    <#
    .SYNOPSIS
        Emits an action result. File logging is optional and never includes a password property.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ActionResult,

        [string]$LogDirectory,

        [switch]$Log
    )

    if ($Log) {
        if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
            throw 'LogDirectory is required when logging an action result.'
        }

        $null = Write-M365ActionLog -ActionResult $ActionResult -LogDirectory $LogDirectory
    }
    else {
        $line = '{0}: {1} [{2}]' -f $ActionResult.Result, $ActionResult.Operation, $ActionResult.UserPrincipalName
        Write-Information -MessageData $line -InformationAction Continue
    }

    Write-Output $ActionResult
}

function Write-M365ActionLog {
    <#
    .SYNOPSIS
        Writes an action result to the log without the temporary password.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ActionResult,

        [Parameter(Mandatory)]
        [string]$LogDirectory
    )

    $level = 'INFO'
    if ($ActionResult.Result -eq 'Failed') {
        $level = 'ERROR'
    }
    elseif ($ActionResult.Result -eq 'WhatIf') {
        $level = 'WHATIF'
    }

    $message = '{0} {1} target={2} user={3} detail={4}' -f $ActionResult.Result, $ActionResult.Operation, $ActionResult.Target, $ActionResult.UserPrincipalName, $ActionResult.Detail
    return Write-M365LifecycleLog -Level $level -Message $message -LogDirectory $LogDirectory
}

function Resolve-M365LogDirectory {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$LogDirectory = '',

        [Parameter(Mandatory)]
        [string]$RepoRoot
    )

    if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
        return (Join-Path -Path $RepoRoot -ChildPath 'logs')
    }

    if ([System.IO.Path]::IsPathRooted($LogDirectory)) {
        return $LogDirectory
    }

    return (Join-Path -Path $RepoRoot -ChildPath $LogDirectory)
}

function Import-M365SettingsFile {
    <#
    .SYNOPSIS
        Loads a settings data file and checks the keys the connect script expects.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Settings file not found: $Path"
    }

    $data = Import-PowerShellDataFile -LiteralPath $Path
    foreach ($key in @('TenantId', 'ClientId', 'Scopes', 'UsageLocation', 'LogDirectory', 'AuthMode')) {
        if (-not $data.ContainsKey($key)) {
            throw "Settings file is missing required key '$key'."
        }
    }

    $mode = [string]$data.AuthMode
    if ($mode -notin @('Interactive', 'Certificate', 'ClientSecret')) {
        throw "AuthMode '$mode' is not supported. Use Interactive, Certificate, or ClientSecret."
    }

    $location = [string]$data.UsageLocation
    if ($location -notmatch '^[A-Z]{2}$') {
        throw "UsageLocation '$location' must be a two-letter uppercase ISO country code."
    }

    $scopes = @($data.Scopes | ForEach-Object { [string]$_ })
    if ($scopes.Count -eq 0) {
        throw 'Scopes must contain at least one Graph scope.'
    }

    $thumbprint = ''
    if ($data.ContainsKey('CertificateThumbprint') -and $null -ne $data.CertificateThumbprint) {
        $thumbprint = [string]$data.CertificateThumbprint
    }

    return [pscustomobject]@{
        TenantId              = [string]$data.TenantId
        ClientId              = [string]$data.ClientId
        CertificateThumbprint = $thumbprint
        AuthMode              = $mode
        UsageLocation         = $location
        LogDirectory          = [string]$data.LogDirectory
        Scopes                = [string[]]$scopes
    }
}

function Resolve-M365ConnectSetting {
    <#
    .SYNOPSIS
        Applies environment overrides. The client secret itself is not copied.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Settings,

        [AllowEmptyString()]
        [string]$AuthMode = '',

        [hashtable]$Environment = @{}
    )

    $mode = $AuthMode
    if ([string]::IsNullOrWhiteSpace($mode)) {
        $mode = [string]$Settings.AuthMode
    }

    if ([string]::IsNullOrWhiteSpace($mode)) {
        $mode = 'Interactive'
    }

    $tenant = [string]$Settings.TenantId
    $client = [string]$Settings.ClientId
    $thumb = [string]$Settings.CertificateThumbprint
    if ($Environment.ContainsKey('M365_TENANT_ID') -and -not [string]::IsNullOrWhiteSpace([string]$Environment['M365_TENANT_ID'])) {
        $tenant = [string]$Environment['M365_TENANT_ID']
    }

    if ($Environment.ContainsKey('M365_CLIENT_ID') -and -not [string]::IsNullOrWhiteSpace([string]$Environment['M365_CLIENT_ID'])) {
        $client = [string]$Environment['M365_CLIENT_ID']
    }

    if ($Environment.ContainsKey('M365_CERT_THUMBPRINT') -and -not [string]::IsNullOrWhiteSpace([string]$Environment['M365_CERT_THUMBPRINT'])) {
        $thumb = [string]$Environment['M365_CERT_THUMBPRINT']
    }

    $hasSecret = $Environment.ContainsKey('M365_CLIENT_SECRET') -and -not [string]::IsNullOrWhiteSpace([string]$Environment['M365_CLIENT_SECRET'])
    return [pscustomobject]@{
        AuthMode              = $mode
        TenantId              = $tenant
        ClientId              = $client
        CertificateThumbprint = $thumb
        HasClientSecret       = [bool]$hasSecret
        Scopes                = [string[]]@($Settings.Scopes)
        UsageLocation         = [string]$Settings.UsageLocation
        LogDirectory          = [string]$Settings.LogDirectory
    }
}

function Test-M365ConnectReadiness {
    <#
    .SYNOPSIS
        Checks app-only settings before Connect-MgGraph is called.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Resolved
    )

    $problems = [System.Collections.Generic.List[string]]::new()
    switch ($Resolved.AuthMode) {
        'Interactive' {
            if (@($Resolved.Scopes).Count -eq 0) {
                [void]$problems.Add('At least one delegated scope is required for interactive sign-in.')
            }
        }
        'Certificate' {
            if (Test-M365PlaceholderIdentifier -Value $Resolved.TenantId) {
                [void]$problems.Add('TenantId is empty or still the sample placeholder.')
            }
            elseif (-not (Test-M365Guid -Value $Resolved.TenantId)) {
                [void]$problems.Add('TenantId is not a GUID.')
            }

            if (Test-M365PlaceholderIdentifier -Value $Resolved.ClientId) {
                [void]$problems.Add('ClientId is empty or still the sample placeholder.')
            }
            elseif (-not (Test-M365Guid -Value $Resolved.ClientId)) {
                [void]$problems.Add('ClientId is not a GUID.')
            }

            if ([string]::IsNullOrWhiteSpace([string]$Resolved.CertificateThumbprint)) {
                [void]$problems.Add('CertificateThumbprint is empty.')
            }
            elseif ($Resolved.CertificateThumbprint -notmatch '^[A-Fa-f0-9]{40}$') {
                [void]$problems.Add('CertificateThumbprint must be 40 hexadecimal characters.')
            }
        }
        'ClientSecret' {
            if (Test-M365PlaceholderIdentifier -Value $Resolved.TenantId) {
                [void]$problems.Add('TenantId is empty or still the sample placeholder.')
            }
            elseif (-not (Test-M365Guid -Value $Resolved.TenantId)) {
                [void]$problems.Add('TenantId is not a GUID.')
            }

            if (Test-M365PlaceholderIdentifier -Value $Resolved.ClientId) {
                [void]$problems.Add('ClientId is empty or still the sample placeholder.')
            }
            elseif (-not (Test-M365Guid -Value $Resolved.ClientId)) {
                [void]$problems.Add('ClientId is not a GUID.')
            }

            if (-not $Resolved.HasClientSecret) {
                [void]$problems.Add('M365_CLIENT_SECRET is not set.')
            }
        }
        default {
            [void]$problems.Add("AuthMode '$($Resolved.AuthMode)' is not supported.")
        }
    }

    return [pscustomobject]@{
        IsValid  = ($problems.Count -eq 0)
        Problems = [string[]]$problems.ToArray()
    }
}

function Test-M365GraphNotFoundError {
    <#
    .SYNOPSIS
        Returns true when a Graph error record looks like HTTP 404.
    #>
    [CmdletBinding()]
    param(
        $ErrorRecord
    )

    if ($null -eq $ErrorRecord) {
        return $false
    }

    $message = [string]$ErrorRecord.Exception.Message
    if ($message -match 'Request_ResourceNotFound' -or $message -match '\(404\)' -or $message -match 'Status:\s*404') {
        return $true
    }

    $responseProperty = $ErrorRecord.Exception.PSObject.Properties['Response']
    if ($null -ne $responseProperty -and $null -ne $responseProperty.Value) {
        $code = [string]$responseProperty.Value.StatusCode
        if ($code -eq '404' -or $code -eq 'NotFound') {
            return $true
        }
    }

    return $false
}

function Get-M365SettingsPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot,

        [AllowEmptyString()]
        [string]$RequestedPath = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return $RequestedPath
    }

    $local = Join-Path -Path $RepoRoot -ChildPath 'config/settings.psd1'
    if (Test-Path -LiteralPath $local -PathType Leaf) {
        return $local
    }

    return (Join-Path -Path $RepoRoot -ChildPath 'config/settings.example.psd1')
}

Export-ModuleMember -Function @(
    'Get-M365OnboardColumn'
    'Get-M365OffboardColumn'
    'Get-M365DefaultGraphScope'
    'Test-M365UserPrincipalName'
    'Get-M365MailNickname'
    'Test-M365Guid'
    'Test-M365PlaceholderIdentifier'
    'Split-M365ListValue'
    'ConvertTo-M365Boolean'
    'Test-M365PasswordComplexity'
    'Get-M365TemporaryPassword'
    'Get-M365PersonName'
    'Resolve-M365UsageLocation'
    'Test-M365OnboardCsv'
    'Test-M365OffboardCsv'
    'Test-M365UserListCsv'
    'Test-M365GroupMembershipCsv'
    'Get-M365NormalizedOffboardUser'
    'ConvertTo-M365OnboardUserBody'
    'Get-M365OnboardPlan'
    'Get-M365OffboardPlan'
    'Get-M365SharedMailboxNote'
    'Get-M365ActionResult'
    'Format-M365LifecycleLogLine'
    'Write-M365LifecycleLog'
    'Write-M365ActionLog'
    'Write-M365Result'
    'Resolve-M365LogDirectory'
    'Import-M365SettingsFile'
    'Resolve-M365ConnectSetting'
    'Test-M365ConnectReadiness'
    'Test-M365GraphNotFoundError'
    'Get-M365SettingsPath'
)
