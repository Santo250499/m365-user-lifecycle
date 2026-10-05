#Requires -Version 7.0

<#
.SYNOPSIS
    Creates Entra ID users from a CSV using Microsoft Graph.

.DESCRIPTION
    Portfolio and lab onboarding script. Validates the CSV locally, then creates each
    user with New-MgUser, assigns an optional license, adds optional group memberships,
    and sets an optional manager. -WhatIf previews the plan and does not contact Graph.
    A generated temporary password is returned only on the create result of a real run.
    It is not written to the log.

    Graph calls live in M365UserLifecycle.Graph.ps1 and run only after ShouldProcess
    returns true. Connect with Connect-M365Graph.ps1 before a real run.

.PARAMETER CsvPath
    Path to an onboarding CSV. See samples/users-onboard.sample.csv.

.PARAMETER UsageLocation
    ISO country code used when a row does not set UsageLocation. The default is AU.

.PARAMETER LogDirectory
    Directory for the applied-change log. Defaults to the logs folder in this repo.
    Relative paths are resolved from the repository root. -WhatIf does not write a log file.

.EXAMPLE
    ./src/New-M365UserFromCsv.ps1 -CsvPath ./samples/users-onboard.sample.csv -WhatIf

.EXAMPLE
    ./src/Connect-M365Graph.ps1
    ./src/New-M365UserFromCsv.ps1 -CsvPath ./samples/users-onboard.sample.csv

.NOTES
    ConfirmImpact is Medium, so a real run does not prompt unless you pass -Confirm
    or lower $ConfirmPreference. Preview with -WhatIf first.
    Use a lab tenant. This script has not been run against a production directory.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({
        if (Test-Path -LiteralPath $_ -PathType Leaf) { return $true }
        throw "CSV file not found: $_"
    })]
    [string]$CsvPath,

    [ValidatePattern('^[A-Z]{2}$')]
    [string]$UsageLocation = 'AU',

    [string]$LogDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Helpers.psm1') -Force

$repoRoot = Split-Path -Parent $PSScriptRoot
$LogDirectory = Resolve-M365LogDirectory -LogDirectory $LogDirectory -RepoRoot $repoRoot
$validation = Test-M365OnboardCsv -Path $CsvPath -UsageLocation $UsageLocation
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
    $plan = @(Get-M365OnboardPlan -Row $row -UsageLocation $UsageLocation)
    $userId = $null
    $createFailed = $false
    if (-not $previewOnly) {
        $existing = Get-M365GraphUserByUpn -UserPrincipalName $row.UserPrincipalName
        if ($null -ne $existing) {
            $userId = [string]$existing.Id
        }
    }

    foreach ($action in $plan) {
        if (-not $PSCmdlet.ShouldProcess($action.Target, $action.Operation)) {
            $preview = Get-M365ActionResult -UserPrincipalName $row.UserPrincipalName -Operation $action.Operation -Target $action.Target -Result WhatIf -Detail $action.Detail
            [void]$results.Add((Write-M365Result -ActionResult $preview))
            continue
        }

        if ($createFailed -and $action.Name -ne 'CreateUser') {
            $skipped = Get-M365ActionResult -UserPrincipalName $row.UserPrincipalName -Operation $action.Operation -Target $action.Target -Result Skipped -Detail 'Skipped because creating the user failed.'
            [void]$results.Add((Write-M365Result -ActionResult $skipped -Log -LogDirectory $LogDirectory))
            continue
        }

        try {
            $outcome = $null
            $temporaryPassword = $null
            switch ($action.Name) {
                'CreateUser' {
                    if (-not [string]::IsNullOrWhiteSpace($userId)) {
                        $outcome = [pscustomobject]@{
                            Status = 'Skipped'
                            Detail = 'User already exists. Creation was skipped. License, group, and manager steps still run.'
                        }
                    }
                    else {
                        $temporaryPassword = Get-M365TemporaryPassword -UserPrincipalName $row.UserPrincipalName -DisplayName $row.DisplayName
                        $body = ConvertTo-M365OnboardUserBody -Row $row -Password $temporaryPassword -UsageLocation $UsageLocation
                        $created = Invoke-M365UserCreation -Body $body
                        $userId = [string]$created.Id
                        $outcome = [pscustomobject]@{
                            Status = 'Applied'
                            Detail = "Created user id $userId. The temporary password is on this result only. The user must change it at next sign-in."
                        }
                    }
                }
                'AssignLicense' {
                    if ([string]::IsNullOrWhiteSpace($userId)) {
                        throw 'User id is not available, so the license was not assigned.'
                    }

                    $outcome = Invoke-M365LicenseAssignment -UserId $userId -SkuPartNumber $action.SkuPartNumber -UsageLocation $row.UsageLocation
                }
                'AddGroup' {
                    if ([string]::IsNullOrWhiteSpace($userId)) {
                        throw 'User id is not available, so the group was not updated.'
                    }

                    $outcome = Invoke-M365GroupAdd -UserId $userId -GroupId $action.GroupId
                }
                'SetManager' {
                    if ([string]::IsNullOrWhiteSpace($userId)) {
                        throw 'User id is not available, so the manager was not set.'
                    }

                    $outcome = Invoke-M365ManagerAssignment -UserId $userId -ManagerUserPrincipalName $action.ManagerUserPrincipalName
                }
                default {
                    throw "Unknown onboarding action '$($action.Name)'."
                }
            }

            $resultParams = @{
                UserPrincipalName = $row.UserPrincipalName
                Operation         = $action.Operation
                Target            = $action.Target
                Result            = $outcome.Status
                Detail            = $outcome.Detail
            }
            if ($null -ne $temporaryPassword) {
                $resultParams['TemporaryPassword'] = $temporaryPassword
            }

            $recorded = Get-M365ActionResult @resultParams
            [void]$results.Add((Write-M365Result -ActionResult $recorded -Log -LogDirectory $LogDirectory))
            if ($outcome.Status -eq 'Failed') {
                $failureCount++
                if ($action.Name -eq 'CreateUser') {
                    $createFailed = $true
                }
            }
        }
        catch {
            $failureCount++
            if ($action.Name -eq 'CreateUser') {
                $createFailed = $true
            }

            $failed = Get-M365ActionResult -UserPrincipalName $row.UserPrincipalName -Operation $action.Operation -Target $action.Target -Result Failed -Detail $_.Exception.Message
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
