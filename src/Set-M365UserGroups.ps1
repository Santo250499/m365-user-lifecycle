#Requires -Version 7.0

<#
.SYNOPSIS
    Adds or removes users from security and Microsoft 365 groups listed in a CSV.

.DESCRIPTION
    Each row needs UserPrincipalName and one or more GroupIds. Adding a user who is
    already a member is skipped. Removing a user who is not a member is skipped.
    -WhatIf does not contact Graph. ConfirmImpact is High because group removal and
    group-based licensing both change access.

.PARAMETER CsvPath
    CSV with UserPrincipalName and GroupIds. Separate multiple ids with semicolons.

.PARAMETER Action
    Add or Remove. The default is Add.

.PARAMETER LogDirectory
    Directory for the applied-change log. -WhatIf does not write a log file.

.EXAMPLE
    ./src/Set-M365UserGroups.ps1 -CsvPath ./out/groups.csv -Action Add -WhatIf

.EXAMPLE
    ./src/Set-M365UserGroups.ps1 -CsvPath ./out/groups.csv -Action Remove -WhatIf

.NOTES
    GroupIds must be object ids from the lab tenant, not the all-zero sample placeholder.
    Dynamic groups cannot have members removed and are reported as failures.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({
        if (Test-Path -LiteralPath $_ -PathType Leaf) { return $true }
        throw "CSV file not found: $_"
    })]
    [string]$CsvPath,

    [ValidateSet('Add', 'Remove')]
    [string]$Action = 'Add',

    [string]$LogDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Helpers.psm1') -Force

$repoRoot = Split-Path -Parent $PSScriptRoot
$LogDirectory = Resolve-M365LogDirectory -LogDirectory $LogDirectory -RepoRoot $repoRoot
$validation = Test-M365GroupMembershipCsv -Path $CsvPath
if (-not $validation.IsValid) {
    throw ("CSV validation failed:`n" + ($validation.Errors -join [Environment]::NewLine))
}

$previewOnly = [bool]$WhatIfPreference
if (-not $previewOnly) {
    . (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Graph.ps1')
    Import-M365GraphSdk
    $null = Assert-M365GraphConnection
}

$results = [System.Collections.Generic.List[object]]::new()
$failureCount = 0

foreach ($row in @($validation.Rows)) {
    $directoryUser = $null
    if (-not $previewOnly) {
        $directoryUser = Get-M365GraphUserByUpn -UserPrincipalName $row.UserPrincipalName
    }

    foreach ($groupId in @($row.GroupIds)) {
        $operation = if ($Action -eq 'Add') { 'Add group member' } else { 'Remove group member' }
        $target = "$($row.UserPrincipalName) -> $groupId"
        $detail = if ($Action -eq 'Add') {
            "Add the user to group $groupId when they are not already a member."
        }
        else {
            "Remove the user from group $groupId when they are a member. Dynamic groups may refuse this."
        }

        if (-not $PSCmdlet.ShouldProcess($target, $operation)) {
            $preview = Get-M365ActionResult -UserPrincipalName $row.UserPrincipalName -Operation $operation -Target $target -Result WhatIf -Detail $detail
            [void]$results.Add((Write-M365Result -ActionResult $preview))
            continue
        }

        if ($null -eq $directoryUser) {
            $failureCount++
            $missing = Get-M365ActionResult -UserPrincipalName $row.UserPrincipalName -Operation $operation -Target $target -Result Failed -Detail 'User was not found.'
            [void]$results.Add((Write-M365Result -ActionResult $missing -Log -LogDirectory $LogDirectory))
            continue
        }

        try {
            $outcome = $null
            if ($Action -eq 'Add') {
                $outcome = Invoke-M365GroupAdd -UserId $directoryUser.Id -GroupId $groupId
            }
            else {
                $outcome = Invoke-M365GroupMemberRemoval -UserId $directoryUser.Id -GroupId $groupId
            }

            $recorded = Get-M365ActionResult -UserPrincipalName $row.UserPrincipalName -Operation $operation -Target $target -Result $outcome.Status -Detail $outcome.Detail
            [void]$results.Add((Write-M365Result -ActionResult $recorded -Log -LogDirectory $LogDirectory))
            if ($outcome.Status -eq 'Failed') {
                $failureCount++
            }
        }
        catch {
            $failureCount++
            $failed = Get-M365ActionResult -UserPrincipalName $row.UserPrincipalName -Operation $operation -Target $target -Result Failed -Detail $_.Exception.Message
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
