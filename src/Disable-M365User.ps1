#Requires -Version 7.0

<#
.SYNOPSIS
    Disables Entra ID users and records the related leaver steps.

.DESCRIPTION
    Offboarding script for a lab tenant. For each user it can disable sign-in,
    revoke sessions, remove direct licenses, remove group memberships, and record
    a manual Exchange shared-mailbox note. The mailbox is not converted: Microsoft
    Graph cannot set a mailbox type to shared, and this script does not connect
    to Exchange Online.

    -WhatIf lists the plan and does not contact Graph. ConfirmImpact is High, so a
    real run asks before each change unless you pass -Confirm:$false after reviewing
    the preview.

.PARAMETER CsvPath
    Offboarding CSV. See samples/users-offboard.sample.csv. A non-blank cell overrides
    the matching switch or boolean parameter for that row.

.PARAMETER UserPrincipalName
    One user to offboard instead of a CSV.

.PARAMETER DisableSignIn
    Set accountEnabled to false. Default is true. CSV column DisableSignIn overrides this.

.PARAMETER RevokeSessions
    Call Revoke-MgUserSignInSession. Default is true.

.PARAMETER RemoveLicenses
    Remove directly assigned licenses. Default is true. This does not remove licenses
    inherited from a group until the group removal step runs.

.PARAMETER RemoveGroups
    Remove security and Microsoft 365 group memberships. Directory roles are not changed.

.PARAMETER AddSharedMailboxNote
    Record the manual Set-Mailbox conversion step in the log. This does not change the mailbox.
    If you also remove licenses, convert the mailbox first or the mailbox may be scheduled for deletion.

.PARAMETER LogDirectory
    Directory for the applied-change log. -WhatIf does not write a log file.

.EXAMPLE
    ./src/Disable-M365User.ps1 -CsvPath ./samples/users-offboard.sample.csv -WhatIf

.EXAMPLE
    ./src/Disable-M365User.ps1 -UserPrincipalName alex.example@contoso.com -AddSharedMailboxNote -RemoveLicenses:$false -WhatIf

.EXAMPLE
    ./src/Connect-M365Graph.ps1
    ./src/Disable-M365User.ps1 -CsvPath ./samples/users-offboard.sample.csv

.NOTES
    Use a lab tenant. This script has not been run against a production directory.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'Csv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [ValidateScript({
        if (Test-Path -LiteralPath $_ -PathType Leaf) { return $true }
        throw "CSV file not found: $_"
    })]
    [string]$CsvPath,

    [Parameter(Mandatory, ParameterSetName = 'Single')]
    [string]$UserPrincipalName,

    [bool]$DisableSignIn = $true,
    [bool]$RevokeSessions = $true,
    [bool]$RemoveLicenses = $true,
    [bool]$RemoveGroups = $true,
    [switch]$AddSharedMailboxNote,
    [string]$LogDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Helpers.psm1') -Force

$repoRoot = Split-Path -Parent $PSScriptRoot
$LogDirectory = Resolve-M365LogDirectory -LogDirectory $LogDirectory -RepoRoot $repoRoot

$users = @()
if ($PSCmdlet.ParameterSetName -eq 'Single') {
    if (-not (Test-M365UserPrincipalName -UserPrincipalName $UserPrincipalName)) {
        throw "UserPrincipalName '$UserPrincipalName' is not a valid UPN."
    }

    $single = Get-M365NormalizedOffboardUser -UserPrincipalName $UserPrincipalName -DisableSignIn:$DisableSignIn -RevokeSessions:$RevokeSessions -RemoveLicenses:$RemoveLicenses -RemoveGroups:$RemoveGroups -AddSharedMailboxNote:([bool]$AddSharedMailboxNote)
    if (@(Get-M365OffboardPlan -User $single).Count -eq 0) {
        throw 'No offboarding actions were selected.'
    }

    $users = @($single)
}
else {
    $validation = Test-M365OffboardCsv -Path $CsvPath -DisableSignIn:$DisableSignIn -RevokeSessions:$RevokeSessions -RemoveLicenses:$RemoveLicenses -RemoveGroups:$RemoveGroups -AddSharedMailboxNote:([bool]$AddSharedMailboxNote)
    if (-not $validation.IsValid) {
        throw ("CSV validation failed:`n" + ($validation.Errors -join [Environment]::NewLine))
    }

    $users = @($validation.Rows)
}

$previewOnly = [bool]$WhatIfPreference
if (-not $previewOnly) {
    . (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Graph.ps1')
    Import-M365GraphSdk
    $null = Assert-M365GraphConnection
}

$results = [System.Collections.Generic.List[object]]::new()
$failureCount = 0

foreach ($user in $users) {
    $plan = @(Get-M365OffboardPlan -User $user)
    $directoryUser = $null
    if (-not $previewOnly) {
        $directoryUser = Get-M365GraphUserByUpn -UserPrincipalName $user.UserPrincipalName
        if ($null -eq $directoryUser) {
            $failureCount++
            $missing = Get-M365ActionResult -UserPrincipalName $user.UserPrincipalName -Operation 'Offboard user' -Target $user.UserPrincipalName -Result Failed -Detail 'User was not found. No changes were made for this row.'
            [void]$results.Add((Write-M365Result -ActionResult $missing -Log -LogDirectory $LogDirectory))
            continue
        }
    }

    foreach ($action in $plan) {
        if (-not $PSCmdlet.ShouldProcess($action.Target, $action.Operation)) {
            $preview = Get-M365ActionResult -UserPrincipalName $user.UserPrincipalName -Operation $action.Operation -Target $action.Target -Result WhatIf -Detail $action.Detail
            [void]$results.Add((Write-M365Result -ActionResult $preview))
            continue
        }

        try {
            $outcome = $null
            switch ($action.Name) {
                'DisableSignIn' {
                    $outcome = Invoke-M365SignInBlock -User $directoryUser
                }
                'RevokeSessions' {
                    $outcome = Invoke-M365SessionRevocation -UserId $directoryUser.Id
                }
                'SharedMailboxNote' {
                    $outcome = [pscustomobject]@{
                        Status = 'Noted'
                        Detail = $action.Detail
                    }
                }
                'RemoveLicenses' {
                    $outcome = Invoke-M365LicenseRemoval -UserId $directoryUser.Id
                }
                'RemoveGroups' {
                    $outcome = Invoke-M365GroupRemoval -UserId $directoryUser.Id
                }
                default {
                    throw "Unknown offboarding action '$($action.Name)'."
                }
            }

            $recorded = Get-M365ActionResult -UserPrincipalName $user.UserPrincipalName -Operation $action.Operation -Target $action.Target -Result $outcome.Status -Detail $outcome.Detail
            [void]$results.Add((Write-M365Result -ActionResult $recorded -Log -LogDirectory $LogDirectory))
            if ($outcome.Status -eq 'Failed') {
                $failureCount++
            }
        }
        catch {
            $failureCount++
            $failed = Get-M365ActionResult -UserPrincipalName $user.UserPrincipalName -Operation $action.Operation -Target $action.Target -Result Failed -Detail $_.Exception.Message
            [void]$results.Add((Write-M365Result -ActionResult $failed -Log -LogDirectory $LogDirectory))
        }
    }
}

foreach ($item in $results) {
    Write-Output $item
}

if ($failureCount -gt 0) {
    throw "Finished with $failureCount failed action(s). Review the log in $LogDirectory."
}
