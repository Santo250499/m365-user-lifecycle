#Requires -Version 7.0

<#
.SYNOPSIS
    Exports a read-only snapshot of selected Entra users to CSV.

.DESCRIPTION
    Reads account status, direct license part numbers, and group object ids for each
    user principal name. This script does not change the directory. It does contact
    Graph, so connect with Connect-M365Graph.ps1 first. The Audit scope profile is
    enough for this export.

    Output is written outside the samples and config folders. Use -Force to replace
    an existing file.

.PARAMETER CsvPath
    CSV with a UserPrincipalName column. Extra columns are ignored.

.PARAMETER UserPrincipalName
    One user instead of a CSV.

.PARAMETER OutputPath
    Destination CSV. Refuses to write into samples/ or config/.

.PARAMETER Force
    Replace OutputPath when it already exists.

.EXAMPLE
    ./src/Connect-M365Graph.ps1 -ScopeProfile Audit
    ./src/Export-M365UserAudit.ps1 -CsvPath ./samples/users-offboard.sample.csv -OutputPath ./out/user-audit.csv

.NOTES
    A row whose user does not exist is written with Status NotFound. The script then
    exits with an error so a partial export is visible and the failure is not silent.
#>
[CmdletBinding(DefaultParameterSetName = 'Csv')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [ValidateScript({
        if (Test-Path -LiteralPath $_ -PathType Leaf) { return $true }
        throw "CSV file not found: $_"
    })]
    [string]$CsvPath,

    [Parameter(Mandatory, ParameterSetName = 'Single')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory)]
    [string]$OutputPath,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Helpers.psm1') -Force

$repoRoot = Split-Path -Parent $PSScriptRoot
$fullOutput = [System.IO.Path]::GetFullPath($OutputPath)
foreach ($protectedName in @('samples', 'config')) {
    $protected = [System.IO.Path]::GetFullPath((Join-Path -Path $repoRoot -ChildPath $protectedName))
    $prefix = $protected.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    if ($fullOutput.StartsWith($prefix, [System.StringComparison]::Ordinal) -or $fullOutput -eq $protected) {
        throw "Refusing to write audit output into the $protectedName directory."
    }
}

if ((Test-Path -LiteralPath $fullOutput) -and -not $Force) {
    throw "Output file already exists: $fullOutput. Pass -Force to replace it."
}

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

. (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Graph.ps1')
Import-M365GraphSdk
$null = Assert-M365GraphConnection

$records = [System.Collections.Generic.List[object]]::new()
$missing = 0
foreach ($upn in $upns) {
    $record = Get-M365GraphAuditRecord -UserPrincipalName $upn
    [void]$records.Add($record)
    if ($record.Status -eq 'NotFound') {
        $missing++
    }
}

$parent = Split-Path -Parent $fullOutput
if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
    $null = New-Item -ItemType Directory -Path $parent -Force
}

$records | Export-Csv -LiteralPath $fullOutput -NoTypeInformation -Encoding utf8
Write-Information -MessageData "Wrote $($records.Count) audit row(s) to $fullOutput." -InformationAction Continue
Write-Output ([pscustomobject]@{
    OutputPath = $fullOutput
    Rows       = $records.Count
    NotFound   = $missing
})

if ($missing -gt 0) {
    throw "Audit export finished with $missing user(s) not found. The CSV was still written."
}
