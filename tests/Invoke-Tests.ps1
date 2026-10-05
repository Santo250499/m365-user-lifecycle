#Requires -Version 7.0

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$pester = Get-Module -ListAvailable -Name Pester |
    Where-Object { $_.Version -ge [version]'5.0.0' } |
    Sort-Object Version -Descending |
    Select-Object -First 1
if ($null -eq $pester) {
    throw 'Pester 5 or newer is required. Run: Install-Module Pester -MinimumVersion 5.0.0 -Scope CurrentUser'
}

Import-Module -Name $pester.Path -Force
$result = Invoke-Pester -Path @(
    (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Helpers.Tests.ps1')
    (Join-Path -Path $PSScriptRoot -ChildPath 'smoke-tests.ps1')
) -Output Detailed -PassThru

if ($null -eq $result -or $result.PassedCount -eq 0 -or $result.FailedCount -gt 0) {
    exit 1
}
