# m365-user-lifecycle

PowerShell toolkit for Entra ID / Microsoft 365 user onboarding and offboarding via Microsoft Graph (WhatIf-safe, sample data only).

It gives an IT support or systems administrator a repeatable way to create lab users from a CSV and to disable them again: block sign-in, revoke sessions, remove direct licenses, and remove group memberships. Every change can be previewed with `-WhatIf` before anything is written to the directory.

## Disclaimer

This is a portfolio and lab starter. It was written to show how these joiner and leaver steps can be automated with the Microsoft Graph PowerShell SDK. It has not been used in production. It is not a copy of an employer script, and the sample users are fictional (`@contoso.com` only).

Do not point it at a production tenant. Use a personal Microsoft 365 developer tenant, or another lab tenant you are allowed to change, and review `-WhatIf` output before a real run.

## Features

- CSV onboarding with `New-MgUser`, optional license assignment, optional group membership, and an optional manager
- Offboarding that disables sign-in, revokes sessions, removes direct licenses, and removes group memberships
- A written note for the Exchange shared-mailbox step, which this repo does not perform
- `-WhatIf` and `-Confirm` on every script that changes the directory
- Local CSV checks that run with no tenant and no Graph modules
- Sample CSVs and an example settings file with placeholder ids
- Pester tests for the validation, UPN, password, and plan logic

## Prerequisites

- PowerShell 7 or newer
- The Microsoft Graph PowerShell SDK modules listed in [requirements.md](requirements.md)
- Permission to change users in a **lab** tenant, plus the Graph scopes below

`-WhatIf` against the sample CSVs does not need the Graph modules or a sign-in. A real run does.

## Graph permissions

| Scope | Used for |
| --- | --- |
| `User.ReadWrite.All` | Create users, block sign-in, set the manager |
| `GroupMember.ReadWrite.All` | Add and remove group members, list memberships |
| `Organization.Read.All` | Resolve a license SkuPartNumber such as `SPE_E3` |
| `LicenseAssignment.ReadWrite.All` | Assign and remove direct licenses |
| `User.RevokeSessions.All` | Revoke refresh tokens and session cookies |

The read-only export can use `User.Read.All`, `GroupMember.Read.All`, and `Organization.Read.All` instead. Delegated and application permission names match for these calls. Details, including what is intentionally not requested, are in [docs/permissions.md](docs/permissions.md).

## Setup

Install the modules:

```powershell
Install-Module -Name @(
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Users.Actions'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Identity.DirectoryManagement'
) -Scope CurrentUser
```

For interactive sign-in you can start from the example settings file. The connect script falls back to it and ignores the all-zero tenant id.

For certificate or client-secret sign-in, copy the example and replace the placeholders locally. `config/settings.psd1` is gitignored:

```powershell
Copy-Item ./config/settings.example.psd1 ./config/settings.psd1
```

A client secret does not belong in the data file. Put it in an untracked `.env` file copied from [.env.example](.env.example), then load it in your shell. Prefer a certificate. `M365_TENANT_ID`, `M365_CLIENT_ID`, `M365_CERT_THUMBPRINT`, and `M365_CLIENT_SECRET` override the data file when they are set. The secret is not copied into the settings object and is not written to the log.

## Usage examples

Preview onboarding. This does not contact Graph:

```powershell
./src/New-M365UserFromCsv.ps1 -CsvPath ./samples/users-onboard.sample.csv -WhatIf
```

After you have replaced the sample UPNs with lab UPNs and connected:

```powershell
./src/Connect-M365Graph.ps1
./src/New-M365UserFromCsv.ps1 -CsvPath ./samples/users-onboard.sample.csv
```

The create result for a real run includes `TemporaryPassword`. The user must change it at next sign-in. That value is not written to `logs/`. Redirect or store the pipeline output somewhere private, then clear it.

Preview offboarding:

```powershell
./src/Disable-M365User.ps1 -CsvPath ./samples/users-offboard.sample.csv -WhatIf
```

One fictional user, including the mailbox note and without removing licenses:

```powershell
./src/Disable-M365User.ps1 -UserPrincipalName alex.example@contoso.com -AddSharedMailboxNote -RemoveLicenses:$false -WhatIf
```

Offboarding uses `ConfirmImpact High`, so a real run asks before each change. Onboarding uses `ConfirmImpact Medium`, so it does not ask unless you pass `-Confirm` or lower `$ConfirmPreference`.

License removal and group changes are also high impact:

```powershell
./src/Remove-M365UserLicenses.ps1 -UserPrincipalName jamie.sample@contoso.com -WhatIf
./src/Set-M365UserGroups.ps1 -CsvPath ./out/groups.csv -Action Add -WhatIf
```

Read-only export after an audit-scoped sign-in:

```powershell
./src/Connect-M365Graph.ps1 -ScopeProfile Audit
./src/Export-M365UserAudit.ps1 -CsvPath ./samples/users-offboard.sample.csv -OutputPath ./out/user-audit.csv
```

Applied changes are appended to `logs/m365-lifecycle-yyyyMMdd.log`. `-WhatIf` prints the plan and does not create that file.

## Sample CSV format

`samples/users-onboard.sample.csv`

| Column | Required | Notes |
| --- | --- | --- |
| `DisplayName` | Yes | 1–256 characters |
| `UserPrincipalName` | Yes | Must be a real lab domain when you leave the sample. Maximum 113 characters |
| `Department` | Column required, value optional | Maximum 64 characters |
| `JobTitle` | Column required, value optional | Maximum 128 characters |
| `LicenseSku` | Column required, value optional | SkuPartNumber such as `SPE_E3`. Blank skips license assignment |
| `GroupIds` | Column required, value optional | Object ids separated by `;`, `,`, or `\|`. Blank skips groups. The all-zero GUID is rejected |
| `GivenName`, `Surname` | No | Filled from `DisplayName` when omitted |
| `ManagerUserPrincipalName` | No | Must already exist on a real run. Cannot be the user's own UPN |
| `UsageLocation` | No | Two-letter country code. Falls back to `-UsageLocation`, default `AU` |

`samples/users-offboard.sample.csv`

| Column | Meaning when blank |
| --- | --- |
| `UserPrincipalName` | Required |
| `DisableSignIn` | `true` |
| `RevokeSessions` | `true` |
| `RemoveLicenses` | `true` |
| `RemoveGroups` | `true` |
| `AddSharedMailboxNote` | `false` |

`true`, `false`, `yes`, `no`, `1`, and `0` are accepted. The sample leaves licenses on Alex Example because that row asks for the shared-mailbox note. Jamie Sample and Riley Placeholder are full removal rows. All three names are fictional.

Group CSV for `Set-M365UserGroups.ps1`:

```csv
UserPrincipalName,GroupIds
alex.example@contoso.com,11111111-2222-3333-4444-555555555555
```

Replace that id with a group object id from the lab tenant before a real run.

## Safety

- No employer data. Sample users are `@contoso.com`. Do not commit real UPNs, tenant names, or screenshots from a work tenant.
- No secrets in git. `config/settings.psd1`, `.env`, certificates, and `logs/` are ignored. Ship only `settings.example.psd1` and `.env.example`.
- No real tenant ids. The example settings use `00000000-0000-0000-0000-000000000000`. Certificate and client-secret modes refuse to connect while that placeholder is still in place.
- Directory changes go through the Microsoft Graph PowerShell SDK, not the legacy MSOnline or AzureAD modules.
- Every mutating script supports `-WhatIf` and `-Confirm` through `SupportsShouldProcess`.
- `-WhatIf` validates the CSV and prints the plan. It does not load the Graph modules and it does not sign in.
- A real run still needs you to connect first, and it should target a lab tenant.
- Logs record the action and the UPN. They do not record the temporary password.
- If a user already exists, onboarding does not recreate that user. License, group, and manager steps still run.
- Removing a direct license does not remove a license inherited from a group. Leaving the group does. Removing the license from a mailbox user can schedule the mailbox for deletion.
- `AddSharedMailboxNote` only writes this manual step: `Set-Mailbox -Identity '<upn>' -Type Shared`. Run that in Exchange Online yourself, in the lab, before you remove the license if the mail must be kept. This repo does not connect to Exchange.
- Dynamic groups can refuse member removal. That failure is reported and the other actions still return a result.
- The audit script refuses to write into `samples/` or `config/`.
- Entra may still reject a generated password if it hits the tenant banned-password list. The script reports that Graph error and does not hide it.

## Tests

The tests do not need a tenant:

```powershell
./tests/Invoke-Tests.ps1
```

`tests/M365UserLifecycle.Helpers.Tests.ps1` covers CSV validation, UPN checks, password rules, plans, settings, and log redaction. `tests/smoke-tests.ps1` runs the sample CSVs with `-WhatIf` and checks that the preview does not create a log file.

Lint the scripts with:

```powershell
Invoke-ScriptAnalyzer -Path ./src -Recurse -Severity Warning,Error
```

## Roadmap

Not implemented, and not claimed as done:

- Autopilot or group-based licensing tags
- A real Exchange Online mailbox conversion, instead of the note
- A ticket-system hook for the joiner and leaver request

## Layout

```text
src/Connect-M365Graph.ps1          Sign-in
src/New-M365UserFromCsv.ps1        Onboarding
src/Disable-M365User.ps1           Offboarding
src/Remove-M365UserLicenses.ps1    Direct license removal
src/Set-M365UserGroups.ps1         Group add or remove
src/Export-M365UserAudit.ps1       Read-only CSV export
src/M365UserLifecycle.Helpers.psm1 CSV, password, and plan helpers
src/M365UserLifecycle.Graph.ps1    Graph cmdlets used after -WhatIf
```

## License

MIT. See [LICENSE](LICENSE).
