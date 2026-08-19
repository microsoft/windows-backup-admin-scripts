# Windows Backup Admin Tooling

PowerShell scripts that help administrators view, export, and delete users'
Windows settings backup data through the Microsoft Graph `windowsSetting` APIs.
They provide a simpler alternative to calling the Graph APIs directly.

## Scripts

| Script | Purpose |
| --- | --- |
| `View-WindowsBackup.ps1` | View and optionally export backup data for yourself or, with the required admin role, another user. |
| `Delete-WindowsBackup.ps1` | Delete all backup data for a user. This is irreversible; the script shows what will be removed and prompts for confirmation. |

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7.
- The `Microsoft.Graph.Authentication` module (installed automatically on first run).
- Internet access to `login.microsoftonline.com` and `graph.microsoft.com` (port 443).

## Permissions and roles

| Action | Delegated scope(s) | Role required |
| --- | --- | --- |
| View your own backup | `UserWindowsSettings.Read` | None |
| View another user's backup | `UserWindowsSettings.Read.All`, `User.Read.All` | Microsoft 365 Backup Administrator |
| Delete a user's backup | `UserWindowsSettings.ReadWrite.All`, `User.Read.All` | Microsoft 365 Backup Administrator |

> The Microsoft 365 Backup Administrator role must be explicitly assigned.
> Global Administrator alone returns HTTP 403.

## Install

The recommended way to use the tooling is the `WindowsBackupAdmin` module from
the [PowerShell Gallery](https://www.powershellgallery.com/packages/WindowsBackupAdmin):

```powershell
Install-Module WindowsBackupAdmin -Scope CurrentUser
```

This also installs the required `Microsoft.Graph.Authentication` dependency. After
installation the commands load automatically; no `Import-Module` is needed.

## Usage (module)

View your own backup:

```powershell
Get-WindowsBackup
```

View another user's backup (requires the admin role and consented scopes):

```powershell
Get-WindowsBackup -UserId user@contoso.com
```

Delete a user's backup (irreversible; prompts for confirmation):

```powershell
Remove-WindowsBackup -UserId user@contoso.com
```

Use `-Fresh` on `Get-WindowsBackup` to force a clean sign-in and switch accounts.

The familiar names `View-WindowsBackup` and `Delete-WindowsBackup` are available
as aliases for `Get-WindowsBackup` and `Remove-WindowsBackup`.

## Usage (scripts directly)

You can also run the scripts from this repository without installing the module:

```powershell
# View your own backup
powershell -ExecutionPolicy Bypass -File .\View-WindowsBackup.ps1

# View another user's backup
powershell -ExecutionPolicy Bypass -File .\View-WindowsBackup.ps1 -UserId user@contoso.com

# Delete a user's backup (irreversible; prompts for confirmation)
powershell -ExecutionPolicy Bypass -File .\Delete-WindowsBackup.ps1 -UserId user@contoso.com
```

## Third-party code

This repository contains no third-party code. The only dependency,
`Microsoft.Graph.Authentication`, is a Microsoft-published module installed at
runtime from the PowerShell Gallery; it is not distributed in this repository.

## Notes

These scripts call Microsoft Graph beta APIs, which are subject to change.
Deletion is permanent and applies to all of a user's backup data.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

## Trademarks

This project may contain trademarks or logos for projects, products, or
services. Authorized use of Microsoft trademarks or logos is subject to and
must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must
not cause confusion or imply Microsoft sponsorship. Any use of third-party
trademarks or logos are subject to those third-party's policies.
