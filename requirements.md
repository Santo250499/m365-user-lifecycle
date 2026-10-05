# Requirements

This file lists PowerShell modules and Microsoft Graph permissions. It is not a Python requirements file.

## Runtime

- PowerShell 7.0 or newer (`pwsh`)
- A lab tenant, such as a Microsoft 365 developer tenant
- The Microsoft Graph PowerShell SDK 2.x modules below

```powershell
Install-Module -Name @(
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Users.Actions'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Identity.DirectoryManagement'
) -Scope CurrentUser
```

`Install-Module Microsoft.Graph` also works. That meta-module installs every service module and is a much larger download.

## Permissions

See [docs/permissions.md](docs/permissions.md). The lifecycle scope set is:

- `User.ReadWrite.All`
- `GroupMember.ReadWrite.All`
- `Organization.Read.All`
- `LicenseAssignment.ReadWrite.All`
- `User.RevokeSessions.All`

The read-only audit export can use `User.Read.All`, `GroupMember.Read.All`, and `Organization.Read.All`.

## Tests and lint

These are only needed to run the local checks. They do not talk to a tenant.

```powershell
Install-Module -Name Pester -MinimumVersion 5.0.0 -Scope CurrentUser
Install-Module -Name PSScriptAnalyzer -Scope CurrentUser
```

```powershell
./tests/Invoke-Tests.ps1
Invoke-ScriptAnalyzer -Path ./src -Recurse -Severity Warning,Error
```
