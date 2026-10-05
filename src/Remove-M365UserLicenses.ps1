#Requires -Version 7.0

<#
.SYNOPSIS
    Removes directly assigned Microsoft 365 licenses from users in a CSV.

.DESCRIPTION
    Reads UserPrincipalName values and, after ShouldProcess, removes each direct
    license with Set-MgUserLicense. Group-based licenses are not removed here.
    -WhatIf does not contact Graph. ConfirmImpact is High.

    The offboarding sample CSV can be used because only UserPrincipalName is required.
    That file's RemoveLicenses column is ignored by this script: every listed user is
    in scope. Use Disable-M365User.ps1 when each row needs its own flags.

.PARAMETER CsvPath
    CSV with a UserPrincipalName column.

.PARAMETER UserPrincipalName
    One user instead of a CSV.

.PARAMETER LogDirectory
    Directory for the applied-change log. -WhatIf does not write a log file.

.EXAMPLE
    ./src/Remove-M365UserLicenses.ps1 -CsvPath ./samples/users-offboard.sample.csv -WhatIf

.EXAMPLE
    ./src/Remove-M365UserLicenses.ps1 -UserPrincipalName jamie.sample@contoso.com -WhatIf

.NOTES
    Removing the license from a user mailbox can schedule that mailbox for deletion.
    Convert the mailbox to shared first if the mail must be kept. See Disable-M365User.ps1.
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

    [string]$LogDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Helpers.psm1') -Force

$repoRoot = Split-Path -Parent $PSScriptRoot
$LogDirectory = Resolve-M365LogDirectory -LogDirectory $LogDirectory -RepoRoot $repoRoot

$upns = @()
if ($PSCmdlet.ParameterSetName -eq 'Single') {
    if (-not (Test-M365UserPrincipalName -UserPrincipalName $UserPrincipalName)) {
        throw "UserPrincipalName '$UserPrincipalName' is not a valid UPN."
    }

    $upns = @($UserPrincipalName.Trim())
}
else {
    $validation = Test-M365UserListCsv -Path $CsvPath
    if (-not $validation.IsValid) {
        throw ("CSV validation failed:`n" + ($validation.Errors -join [Environment]::NewLine))
    }

    $upns = @($validation.Rows | ForEach-Object { $_.UserPrincipalName })
}

$previewOnly = [bool]$WhatIfPreference
if (-not $previewOnly) {
    . (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Graph.ps1')
    Import-M365GraphSdk
    $null = Assert-M365GraphConnection
}

$results = [System.Collections.Generic.List[object]]::new()
$failureCount = 0

foreach ($upn in $upns) {
    $operation = 'Remove direct licenses'
    $detail = 'Remove every directly assigned license. Group-based licenses remain until the user leaves the licensing group. If this user has a mailbox, convert it to shared before this step when the mail must be kept.'
    if (-not $PSCmdlet.ShouldProcess($upn, $operation)) {
        $preview = Get-M365ActionResult -UserPrincipalName $upn -Operation $operation -Target $upn -Result WhatIf -Detail $detail
        [void]$results.Add((Write-M365Result -ActionResult $preview))
        continue
    }

    try {
        $directoryUser = Get-M365GraphUserByUpn -UserPrincipalName $upn
        if ($null -eq $directoryUser) {
            $failureCount++
            $missing = Get-M365ActionResult -UserPrincipalName $upn -Operation $operation -Target $upn -Result Failed -Detail 'User was not found.'
            [void]$results.Add((Write-M365Result -ActionResult $missing -Log -LogDirectory $LogDirectory))
            continue
        }

        $outcome = Invoke-M365LicenseRemoval -UserId $directoryUser.Id
        $recorded = Get-M365ActionResult -UserPrincipalName $upn -Operation $operation -Target $upn -Result $outcome.Status -Detail $outcome.Detail
        [void]$results.Add((Write-M365Result -ActionResult $recorded -Log -LogDirectory $LogDirectory))
        if ($outcome.Status -eq 'Failed') {
            $failureCount++
        }
    }
    catch {
        $failureCount++
        $failed = Get-M365ActionResult -UserPrincipalName $upn -Operation $operation -Target $upn -Result Failed -Detail $_.Exception.Message
        [void]$results.Add((Write-M365Result -ActionResult $failed -Log -LogDirectory $LogDirectory))
    }
}

foreach ($item in $results) {
    Write-Output $item
}

if ($failureCount -gt 0) {
    throw "Finished with $failureCount failed action(s). Review the log in $LogDirectory."
}
