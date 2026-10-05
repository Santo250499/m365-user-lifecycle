#Requires -Version 7.0

<#
.SYNOPSIS
    Connects to Microsoft Graph for the lifecycle scripts.

.DESCRIPTION
    Interactive sign-in is the default and is the simplest option for a lab tenant.
    Certificate and client-secret modes read the tenant id, client id, and secret
    material from a gitignored settings file and from environment variables.
    The client secret is never read from the settings file and is never written to the log.

.PARAMETER AuthMode
    Interactive, Certificate, or ClientSecret. When omitted, the settings file decides.

.PARAMETER SettingsPath
    Path to a PowerShell data file. Defaults to config/settings.psd1 when that file
    exists, otherwise config/settings.example.psd1.

.PARAMETER Scopes
    Graph scopes to request. Overrides the settings file and -ScopeProfile.

.PARAMETER ScopeProfile
    Lifecycle requests the write scopes. Audit requests the read scopes used by the export script.

.EXAMPLE
    ./src/Connect-M365Graph.ps1

.EXAMPLE
    ./src/Connect-M365Graph.ps1 -ScopeProfile Audit

.EXAMPLE
    ./src/Connect-M365Graph.ps1 -AuthMode Certificate -SettingsPath ./config/settings.psd1

.NOTES
    Point this at a personal Microsoft 365 developer tenant or another lab tenant.
    Do not store the client secret in git.
#>
[CmdletBinding()]
param(
    [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
    [string]$AuthMode,

    [string]$SettingsPath,

    [string[]]$Scopes,

    [ValidateSet('Lifecycle', 'Audit')]
    [string]$ScopeProfile = 'Lifecycle'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Helpers.psm1') -Force
. (Join-Path -Path $PSScriptRoot -ChildPath 'M365UserLifecycle.Graph.ps1')

function Connect-M365GraphWithClientSecret {
    [CmdletBinding()]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingConvertToSecureStringWithPlainText',
        '',
        Justification = 'The client secret is read from the process environment at runtime and is not stored in the repository.')]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$ClientId
    )

    $secret = [Environment]::GetEnvironmentVariable('M365_CLIENT_SECRET')
    if ([string]::IsNullOrWhiteSpace($secret)) {
        throw 'M365_CLIENT_SECRET is not set. Load it from a local .env file that is not committed.'
    }

    try {
        $secure = ConvertTo-SecureString -String $secret -AsPlainText -Force
        $credential = [pscredential]::new($ClientId, $secure)
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome -ErrorAction Stop
    }
    finally {
        $secret = $null
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$resolvedSettingsPath = Get-M365SettingsPath -RepoRoot $repoRoot -RequestedPath $SettingsPath
$settings = Import-M365SettingsFile -Path $resolvedSettingsPath

if ($resolvedSettingsPath.EndsWith('settings.example.psd1', [System.StringComparison]::OrdinalIgnoreCase)) {
    Write-Warning 'Using config/settings.example.psd1. Copy it to config/settings.psd1 before certificate or client-secret sign-in.'
}

$environment = @{}
foreach ($name in @('M365_TENANT_ID', 'M365_CLIENT_ID', 'M365_CERT_THUMBPRINT', 'M365_CLIENT_SECRET')) {
    $value = [Environment]::GetEnvironmentVariable($name)
    if ($null -ne $value) {
        $environment[$name] = $value
    }
}

$resolved = Resolve-M365ConnectSetting -Settings $settings -AuthMode $AuthMode -Environment $environment
$readiness = Test-M365ConnectReadiness -Resolved $resolved
if (-not $readiness.IsValid) {
    throw ($readiness.Problems -join ' ')
}

$selectedScopes = @($resolved.Scopes)
if ($PSBoundParameters.ContainsKey('Scopes')) {
    $selectedScopes = @($Scopes)
}
elseif ($PSBoundParameters.ContainsKey('ScopeProfile')) {
    $selectedScopes = @(Get-M365DefaultGraphScope -ScopeProfile $ScopeProfile)
}

if ($selectedScopes.Count -eq 0) {
    throw 'No Graph scopes were selected.'
}

Import-M365GraphSdk

switch ($resolved.AuthMode) {
    'Interactive' {
        $connectParams = @{
            Scopes    = $selectedScopes
            NoWelcome = $true
        }
        if (-not (Test-M365PlaceholderIdentifier -Value $resolved.TenantId)) {
            $connectParams['TenantId'] = $resolved.TenantId
        }

        Connect-MgGraph @connectParams -ErrorAction Stop
    }
    'Certificate' {
        Connect-MgGraph -TenantId $resolved.TenantId -ClientId $resolved.ClientId -CertificateThumbprint $resolved.CertificateThumbprint -NoWelcome -ErrorAction Stop
    }
    'ClientSecret' {
        Connect-M365GraphWithClientSecret -TenantId $resolved.TenantId -ClientId $resolved.ClientId
    }
}

$context = Assert-M365GraphConnection
$logDirectory = Resolve-M365LogDirectory -LogDirectory $settings.LogDirectory -RepoRoot $repoRoot
$account = ''
$accountProperty = $context.PSObject.Properties['Account']
if ($null -ne $accountProperty -and $null -ne $accountProperty.Value) {
    $account = [string]$accountProperty.Value
}

$null = Write-M365LifecycleLog -Level INFO -Message "Connected. AuthMode=$($resolved.AuthMode) TenantId=$($context.TenantId) Account=$account" -LogDirectory $logDirectory

$contextScopes = @()
$scopeProperty = $context.PSObject.Properties['Scopes']
if ($null -ne $scopeProperty -and $null -ne $scopeProperty.Value) {
    $contextScopes = @($scopeProperty.Value)
}

Write-Output ([pscustomobject]@{
    AuthMode = $resolved.AuthMode
    TenantId = [string]$context.TenantId
    Account  = $account
    AppId    = [string]$context.ClientId
    Scopes   = [string[]]$contextScopes
})
