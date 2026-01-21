# M365-Management

A collection of PowerShell scripts, automation runbooks, and administrative tools designed for the management and maintenance of Microsoft 365 tenants.

## Repository Contents

### 1. Exchange Management
* **Script:** `Monitor-FullMailboxes.ps1`
* **Description:** An Azure Automation runbook that monitors Exchange Online mailboxes for storage usage and alerts administrators when quotas exceed a defined threshold (default 90%).
* **Documentation:** [View detailed usage instructions](Exchange/Monitor-FullMailboxes/README.md)

## Getting Started

### Prerequisites
The scripts in this repository primarily rely on the **Exchange Online Management** and **Az** PowerShell modules. Ensure you have the necessary modules installed or imported into your Azure Automation account:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Az.Automation -Scope CurrentUser
```

### Usage
1. Clone this repository:
   ```bash
   git clone [https://github.com/jsorelcx/m365-management.git](https://github.com/jsorelcx/m365-management.git)
   ```
2. Navigate to the relevant directory (e.g., `Exchange/`).
3. Review the specific documentation (`README.md`) for the tool you intend to use before execution.

## Disclaimer
These scripts are provided "as-is" for use in administrative scenarios. **Always test** in a non-production environment before deploying to a live tenant.

## License
Licensed under GPL-3.0. See [LICENSE](LICENSE) for full details.
