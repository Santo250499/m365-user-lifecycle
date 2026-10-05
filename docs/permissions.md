# Microsoft Graph permissions

These are the permissions the scripts call. Consent them on a lab app registration, or sign in with an account that can consent, in a personal Microsoft 365 developer tenant. Admin consent is required for all of them.

The scripts use the Microsoft Graph PowerShell SDK (`Microsoft.Graph.*`). They do not use the legacy MSOnline or AzureAD modules.

## Delegated and application

| Action | Cmdlet | Least delegated permission | Least application permission |
| --- | --- | --- | --- |
| Create a user | `New-MgUser` | `User.ReadWrite.All` | `User.ReadWrite.All` |
| Block sign-in | `Update-MgUser` | `User.ReadWrite.All` | `User.ReadWrite.All` |
| Read a user | `Get-MgUser` | `User.Read.All` | `User.Read.All` |
| Assign or remove a direct license | `Set-MgUserLicense` | `LicenseAssignment.ReadWrite.All` | `LicenseAssignment.ReadWrite.All` |
| List subscribed SKUs | `Get-MgSubscribedSku` | `Organization.Read.All` | `Organization.Read.All` |
| Read license details | `Get-MgUserLicenseDetail` | `User.Read.All` | `User.Read.All` |
| Revoke sessions | `Revoke-MgUserSignInSession` | `User.RevokeSessions.All` | `User.RevokeSessions.All` |
| Add a group member | `New-MgGroupMember` | `GroupMember.ReadWrite.All` | `GroupMember.ReadWrite.All` |
| Remove a group member | `Remove-MgGroupMemberDirectoryObjectByRef` | `GroupMember.ReadWrite.All` | `GroupMember.ReadWrite.All` |
| List a user's groups | `Get-MgUserMemberOf` | `GroupMember.Read.All` | `GroupMember.Read.All` |
| Set a manager | `Set-MgUserManagerByRef` | `User.ReadWrite.All` | `User.ReadWrite.All` |

`User.ReadWrite.All` covers the user reads used while creating and updating accounts. `GroupMember.ReadWrite.All` covers group membership reads. The audit export can therefore use the smaller read set.

## Scope sets in this repo

Lifecycle (onboarding, offboarding, licenses, groups), from `Get-M365DefaultGraphScope -ScopeProfile Lifecycle`:

- `User.ReadWrite.All`
- `GroupMember.ReadWrite.All`
- `Organization.Read.All`
- `LicenseAssignment.ReadWrite.All`
- `User.RevokeSessions.All`

Audit (read-only export), from `Get-M365DefaultGraphScope -ScopeProfile Audit`:

- `User.Read.All`
- `GroupMember.Read.All`
- `Organization.Read.All`

`Directory.ReadWrite.All` is not required. It is broader than these scripts need.

## What these permissions do not cover

Shared-mailbox conversion is an Exchange Online action (`Set-Mailbox -Type Shared`). This repo only writes a note telling you to do that yourself. It does not request Exchange permissions and it does not change a mailbox.

A signed-in user still needs an Entra role that is allowed to perform the action, such as User Administrator for user and license changes and Groups Administrator for membership changes. A permission on the app is not a substitute for that role.

## Application auth

Prefer a certificate over a client secret. If you use a client secret for a lab, put it in `M365_CLIENT_SECRET` in your shell or in an untracked `.env` file. Do not add it to `settings.psd1`.

Application `User.RevokeSessions.All` is supported. Keep the app registration in the lab tenant, and delete it when you are finished with the demo.
