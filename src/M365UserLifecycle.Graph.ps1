#Requires -Version 7.0

<#
.SYNOPSIS
    Microsoft Graph calls used by the lifecycle scripts.

.DESCRIPTION
    Dot-source this file only for a real run. -WhatIf returns before this file is loaded.
    Each function performs one directory change or read. The entry script is responsible
    for ShouldProcess, so these functions do not prompt again.

    Cmdlets used:
      Get-MgContext, Get-MgUser, New-MgUser, Update-MgUser,
      Get-MgSubscribedSku, Set-MgUserLicense, Get-MgUserLicenseDetail,
      Revoke-MgUserSignInSession, Get-MgUserMemberOf, New-MgGroupMember,
      Remove-MgGroupMemberDirectoryObjectByRef, Set-MgUserManagerByRef
#>

Set-StrictMode -Version Latest

function Get-M365GraphOutcome {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Applied', 'Skipped', 'Failed', 'Noted')]
        [string]$Status,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Detail
    )

    return [pscustomobject]@{
        Status = $Status
        Detail = $Detail
    }
}

function Get-M365GraphCollection {
    [CmdletBinding()]
    param(
        $Value
    )

    if ($null -eq $Value) {
        return @()
    }

    return @($Value)
}

function Get-M365OdataType {
    [CmdletBinding()]
    param(
        $DirectoryObject
    )

    if ($null -eq $DirectoryObject) {
        return ''
    }

    $direct = $DirectoryObject.PSObject.Properties['@odata.type']
    if ($null -ne $direct -and -not [string]::IsNullOrWhiteSpace([string]$direct.Value)) {
        return [string]$direct.Value
    }

    $additional = $DirectoryObject.PSObject.Properties['AdditionalProperties']
    if ($null -eq $additional -or $null -eq $additional.Value) {
        return ''
    }

    $bag = $additional.Value
    if ($bag -is [System.Collections.IDictionary] -and $bag.Contains('@odata.type')) {
        return [string]$bag['@odata.type']
    }

    return ''
}

function Import-M365GraphSdk {
    <#
    .SYNOPSIS
        Imports the Graph SDK modules these scripts call.
    #>
    [CmdletBinding()]
    param()

    $modules = @(
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Users'
        'Microsoft.Graph.Users.Actions'
        'Microsoft.Graph.Groups'
        'Microsoft.Graph.Identity.DirectoryManagement'
    )

    $missing = [System.Collections.Generic.List[string]]::new()
    foreach ($name in $modules) {
        $available = Get-Module -ListAvailable -Name $name | Select-Object -First 1
        if ($null -eq $available) {
            [void]$missing.Add($name)
        }
    }

    if ($missing.Count -gt 0) {
        throw "Missing Microsoft Graph modules: $($missing -join ', '). Install them with the commands in requirements.md."
    }

    foreach ($name in $modules) {
        Import-Module -Name $name -ErrorAction Stop
    }

    $commands = @(
        'Connect-MgGraph'
        'Get-MgContext'
        'Get-MgUser'
        'New-MgUser'
        'Update-MgUser'
        'Get-MgSubscribedSku'
        'Set-MgUserLicense'
        'Get-MgUserLicenseDetail'
        'Revoke-MgUserSignInSession'
        'Get-MgUserMemberOf'
        'New-MgGroupMember'
        'Remove-MgGroupMemberDirectoryObjectByRef'
        'Set-MgUserManagerByRef'
    )

    $missingCommands = [System.Collections.Generic.List[string]]::new()
    foreach ($command in $commands) {
        if ($null -eq (Get-Command -Name $command -ErrorAction SilentlyContinue)) {
            [void]$missingCommands.Add($command)
        }
    }

    if ($missingCommands.Count -gt 0) {
        throw "The Graph modules loaded, but these cmdlets are missing: $($missingCommands -join ', '). Install Microsoft.Graph SDK 2.x."
    }
}

function Assert-M365GraphConnection {
    [CmdletBinding()]
    param()

    $context = Get-MgContext
    if ($null -eq $context) {
        throw 'Not connected to Microsoft Graph. Run src/Connect-M365Graph.ps1 first.'
    }

    return $context
}

function Get-M365GraphUserByUpn {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    try {
        return Get-MgUser -UserId $UserPrincipalName -Property @(
            'Id'
            'DisplayName'
            'UserPrincipalName'
            'AccountEnabled'
            'UsageLocation'
            'AssignedLicenses'
            'Department'
            'JobTitle'
        ) -ErrorAction Stop
    }
    catch {
        if (Test-M365GraphNotFoundError -ErrorRecord $_) {
            return $null
        }

        throw
    }
}

function Get-M365SubscribedSkuId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SkuPartNumber
    )

    if (-not (Get-Variable -Name 'M365SkuCache' -Scope Script -ErrorAction SilentlyContinue)) {
        $script:M365SkuCache = @{}
    }

    if (-not $script:M365SkuCache.ContainsKey($SkuPartNumber)) {
        $skus = Get-M365GraphCollection -Value (Get-MgSubscribedSku -All -ErrorAction Stop)
        foreach ($sku in $skus) {
            if ($null -eq $sku) {
                continue
            }

            $part = [string]$sku.SkuPartNumber
            if (-not [string]::IsNullOrWhiteSpace($part)) {
                $script:M365SkuCache[$part] = $sku.SkuId
            }
        }
    }

    if (-not $script:M365SkuCache.ContainsKey($SkuPartNumber)) {
        $known = @($script:M365SkuCache.Keys) -join ', '
        if ([string]::IsNullOrWhiteSpace($known)) {
            $known = '(none)'
        }

        throw "No subscribed SKU matches SkuPartNumber '$SkuPartNumber'. Subscribed part numbers: $known."
    }

    return $script:M365SkuCache[$SkuPartNumber]
}

function Invoke-M365UserCreation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Body
    )

    # New-MgUser returns the created user. The password in $Body is not logged here.
    return New-MgUser -BodyParameter $Body -ErrorAction Stop
}

function Invoke-M365SignInBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $User
    )

    $enabled = $true
    $enabledProperty = $User.PSObject.Properties['AccountEnabled']
    if ($null -ne $enabledProperty -and $null -ne $enabledProperty.Value) {
        $enabled = [bool]$enabledProperty.Value
    }

    if (-not $enabled) {
        return Get-M365GraphOutcome -Status Skipped -Detail 'Sign-in is already disabled.'
    }

    Update-MgUser -UserId $User.Id -AccountEnabled:$false -ErrorAction Stop
    return Get-M365GraphOutcome -Status Applied -Detail 'Set accountEnabled to false.'
}

function Invoke-M365SessionRevocation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    $null = Revoke-MgUserSignInSession -UserId $UserId -ErrorAction Stop
    return Get-M365GraphOutcome -Status Applied -Detail 'Revoked sign-in sessions.'
}

function Invoke-M365LicenseAssignment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string]$SkuPartNumber,

        [Parameter(Mandatory)]
        [string]$UsageLocation
    )

    $skuId = Get-M365SubscribedSkuId -SkuPartNumber $SkuPartNumber
    $user = Get-MgUser -UserId $UserId -Property UsageLocation, AssignedLicenses -ErrorAction Stop
    $usageNote = ''
    $currentUsage = ''
    $usageProperty = $user.PSObject.Properties['UsageLocation']
    if ($null -ne $usageProperty -and $null -ne $usageProperty.Value) {
        $currentUsage = [string]$usageProperty.Value
    }

    if ([string]::IsNullOrWhiteSpace($currentUsage)) {
        Update-MgUser -UserId $UserId -UsageLocation $UsageLocation -ErrorAction Stop
        $usageNote = " Set usage location to $UsageLocation."
    }

    $assigned = @()
    $licenseProperty = $user.PSObject.Properties['AssignedLicenses']
    if ($null -ne $licenseProperty) {
        $assigned = Get-M365GraphCollection -Value $licenseProperty.Value
    }

    foreach ($license in $assigned) {
        if ($null -ne $license -and ([string]$license.SkuId) -eq ([string]$skuId)) {
            return Get-M365GraphOutcome -Status Skipped -Detail "License $SkuPartNumber is already assigned.$usageNote"
        }
    }

    Set-MgUserLicense -UserId $UserId -AddLicenses @(@{ SkuId = $skuId }) -RemoveLicenses @() -ErrorAction Stop
    return Get-M365GraphOutcome -Status Applied -Detail "Assigned $SkuPartNumber ($skuId).$usageNote"
}

function Invoke-M365LicenseRemoval {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    $user = Get-MgUser -UserId $UserId -Property AssignedLicenses -ErrorAction Stop
    $skuIds = [System.Collections.Generic.List[string]]::new()
    $licenseProperty = $user.PSObject.Properties['AssignedLicenses']
    $assigned = @()
    if ($null -ne $licenseProperty) {
        $assigned = Get-M365GraphCollection -Value $licenseProperty.Value
    }

    foreach ($license in $assigned) {
        if ($null -eq $license) {
            continue
        }

        $skuId = [string]$license.SkuId
        if (-not [string]::IsNullOrWhiteSpace($skuId)) {
            [void]$skuIds.Add($skuId)
        }
    }

    if ($skuIds.Count -eq 0) {
        return Get-M365GraphOutcome -Status Skipped -Detail 'No directly assigned licenses were found. Group-based licenses are removed by leaving the licensing group.'
    }

    Set-MgUserLicense -UserId $UserId -AddLicenses @() -RemoveLicenses $skuIds.ToArray() -ErrorAction Stop
    return Get-M365GraphOutcome -Status Applied -Detail ("Removed direct license SKU ids: " + ($skuIds -join ', ') + '.')
}

function Test-M365GraphMembership {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string]$GroupId
    )

    $memberOf = Get-M365GraphCollection -Value (Get-MgUserMemberOf -UserId $UserId -All -ErrorAction Stop)
    foreach ($entry in $memberOf) {
        if ($null -ne $entry -and ([string]$entry.Id) -eq $GroupId) {
            return $true
        }
    }

    return $false
}

function Invoke-M365GroupAdd {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string]$GroupId
    )

    if (Test-M365GraphMembership -UserId $UserId -GroupId $GroupId) {
        return Get-M365GraphOutcome -Status Skipped -Detail "User is already a member of $GroupId."
    }

    New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $UserId -ErrorAction Stop
    return Get-M365GraphOutcome -Status Applied -Detail "Added the user to group $GroupId."
}

function Invoke-M365GroupMemberRemoval {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string]$GroupId
    )

    if (-not (Test-M365GraphMembership -UserId $UserId -GroupId $GroupId)) {
        return Get-M365GraphOutcome -Status Skipped -Detail "User is not a member of $GroupId."
    }

    try {
        Remove-MgGroupMemberDirectoryObjectByRef -GroupId $GroupId -DirectoryObjectId $UserId -ErrorAction Stop
    }
    catch {
        if (Test-M365GraphNotFoundError -ErrorRecord $_) {
            return Get-M365GraphOutcome -Status Skipped -Detail "Group $GroupId no longer lists the user."
        }

        return Get-M365GraphOutcome -Status Failed -Detail "Could not remove group ${GroupId}: $($_.Exception.Message)"
    }

    return Get-M365GraphOutcome -Status Applied -Detail "Removed the user from group $GroupId."
}

function Invoke-M365GroupRemoval {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    $memberOf = Get-M365GraphCollection -Value (Get-MgUserMemberOf -UserId $UserId -All -ErrorAction Stop)
    $removed = [System.Collections.Generic.List[string]]::new()
    $failed = [System.Collections.Generic.List[string]]::new()
    $groupCount = 0

    foreach ($entry in $memberOf) {
        if ($null -eq $entry) {
            continue
        }

        if ((Get-M365OdataType -DirectoryObject $entry) -ne '#microsoft.graph.group') {
            continue
        }

        $groupCount++
        $groupId = [string]$entry.Id
        try {
            Remove-MgGroupMemberDirectoryObjectByRef -GroupId $groupId -DirectoryObjectId $UserId -ErrorAction Stop
            [void]$removed.Add($groupId)
        }
        catch {
            if (Test-M365GraphNotFoundError -ErrorRecord $_) {
                continue
            }

            [void]$failed.Add("$groupId`: $($_.Exception.Message)")
        }
    }

    if ($groupCount -eq 0) {
        return Get-M365GraphOutcome -Status Skipped -Detail 'No group memberships were found. Directory roles were not changed.'
    }

    if ($failed.Count -gt 0) {
        return Get-M365GraphOutcome -Status Failed -Detail ("Removed $($removed.Count) group(s). Failed: " + ($failed -join '; '))
    }

    return Get-M365GraphOutcome -Status Applied -Detail ("Removed group ids: " + ($removed -join ', ') + '.')
}

function Invoke-M365ManagerAssignment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string]$ManagerUserPrincipalName
    )

    $manager = Get-M365GraphUserByUpn -UserPrincipalName $ManagerUserPrincipalName
    if ($null -eq $manager) {
        return Get-M365GraphOutcome -Status Failed -Detail "Manager $ManagerUserPrincipalName was not found."
    }

    $body = @{
        '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($manager.Id)"
    }
    Set-MgUserManagerByRef -UserId $UserId -BodyParameter $body -ErrorAction Stop
    return Get-M365GraphOutcome -Status Applied -Detail "Set manager to $ManagerUserPrincipalName."
}

function Get-M365GraphAuditRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    $user = Get-M365GraphUserByUpn -UserPrincipalName $UserPrincipalName
    if ($null -eq $user) {
        return [pscustomobject]@{
            UserPrincipalName      = $UserPrincipalName
            DisplayName            = ''
            AccountEnabled         = ''
            Department             = ''
            JobTitle               = ''
            LicenseSkuPartNumbers  = ''
            GroupIds               = ''
            Status                 = 'NotFound'
        }
    }

    $licenseNames = [System.Collections.Generic.List[string]]::new()
    $details = Get-M365GraphCollection -Value (Get-MgUserLicenseDetail -UserId $user.Id -All -ErrorAction Stop)
    foreach ($detail in $details) {
        if ($null -eq $detail) {
            continue
        }

        $part = [string]$detail.SkuPartNumber
        if (-not [string]::IsNullOrWhiteSpace($part)) {
            [void]$licenseNames.Add($part)
        }
    }

    $groupIds = [System.Collections.Generic.List[string]]::new()
    $memberOf = Get-M365GraphCollection -Value (Get-MgUserMemberOf -UserId $user.Id -All -ErrorAction Stop)
    foreach ($entry in $memberOf) {
        if ($null -eq $entry) {
            continue
        }

        if ((Get-M365OdataType -DirectoryObject $entry) -eq '#microsoft.graph.group') {
            [void]$groupIds.Add([string]$entry.Id)
        }
    }

    $department = ''
    $departmentProperty = $user.PSObject.Properties['Department']
    if ($null -ne $departmentProperty -and $null -ne $departmentProperty.Value) {
        $department = [string]$departmentProperty.Value
    }

    $jobTitle = ''
    $jobProperty = $user.PSObject.Properties['JobTitle']
    if ($null -ne $jobProperty -and $null -ne $jobProperty.Value) {
        $jobTitle = [string]$jobProperty.Value
    }

    return [pscustomobject]@{
        UserPrincipalName     = [string]$user.UserPrincipalName
        DisplayName           = [string]$user.DisplayName
        AccountEnabled        = [string]$user.AccountEnabled
        Department            = $department
        JobTitle              = $jobTitle
        LicenseSkuPartNumbers = ($licenseNames -join ';')
        GroupIds              = ($groupIds -join ';')
        Status                = 'Found'
    }
}
