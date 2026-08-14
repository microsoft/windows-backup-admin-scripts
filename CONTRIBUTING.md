# Contributing to windows-backup-admin-scripts

Thank you for your interest in contributing. This repository contains PowerShell
scripts that help administrators view, export, and delete users' Windows settings
backup data through the Microsoft Graph `windowsSetting` APIs.

## Contributor License Agreement

This project welcomes contributions and suggestions. Most contributions require
you to agree to a Contributor License Agreement (CLA) declaring that you have
the right to, and actually do, grant us the rights to use your contribution. For
details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether
you need to provide a CLA and decorate the PR appropriately (for example, status
check, comment). Simply follow the instructions provided by the bot. You will
only need to do this once across all repos using our CLA.

## Code of Conduct

This project has adopted the
[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the
[Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any
additional questions or comments.

## Before you start

> [!IMPORTANT]
> `Delete-WindowsBackup.ps1` permanently and irreversibly removes **all** of a
> user's Windows settings backup data. `View-WindowsBackup.ps1` can export data
> that includes personal information. **Always test changes against a
> non-production ("test") tenant and a test user first.** Never validate changes
> against a production tenant or a real user's data.

## How to contribute

1. **Fork** the repository and create a topic branch from `main`.
2. Make your change. Keep pull requests small and focused, ideally one fix or
   improvement per PR.
3. Keep the copyright header at the top of every `.ps1` file:

   ```powershell
   # Copyright (c) Microsoft Corporation.
   # Licensed under the MIT License.
   ```

4. Test both scripts on **Windows PowerShell 5.1** and **PowerShell 7**, since
   both are supported.
5. Do not commit exported backup data or credential material. The `.gitignore`
   already excludes common patterns; double-check before committing.
6. Open a pull request describing what you changed and how you tested it.

## Reporting issues

Use [GitHub Issues](../../issues) to report bugs or request features. Please
search existing issues first to avoid duplicates. When filing a bug, include the
script name, the exact command, your PowerShell version, and the error message.
Redact any tenant IDs, object IDs, or user principal names.

## Trademarks

This project may contain trademarks or logos for projects, products, or
services. Authorized use of Microsoft trademarks or logos is subject to and
must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must
not cause confusion or imply Microsoft sponsorship. Any use of third-party
trademarks or logos are subject to those third-party's policies.
