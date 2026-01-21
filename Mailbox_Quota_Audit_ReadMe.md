# M365-Management: Automated Full Mailbox Alerting

This PowerShell script is designed to run in an **Azure Automation** environment. It monitors Exchange Online mailboxes for storage usage and sends an alert if a mailbox exceeds a defined capacity threshold (default: 90%).

## Overview

* **Technology:** Azure Automation (Cloud Job)
* **Language:** PowerShell 5.1
* **Authentication:** System-Assigned Managed Identity (Secure, no hardcoded credentials)
* **Permissions:** Read-only access to Exchange Online mailbox statistics.

## Prerequisites

Before deploying this script, ensure you have the following Azure resources configured:

1.  **Azure Automation Account:** A standard automation account.
2.  **Modules:** The `ExchangeOnlineManagement` module (v3.0 or later) must be imported into the Automation Account.
3.  **Managed Identity:**
    * The Automation Account must have **System Assigned Identity** enabled.
    * The Identity must be assigned the **Exchange Administrator** role in Microsoft Entra ID.

## Implementation Guide

Follow these steps to deploy the alerting solution:

### 1. Create the Runbook
1.  Navigate to your Azure Automation Account > **Runbooks**.
2.  Create a new Runbook named `Monitor-FullMailboxes`.
3.  Select **PowerShell** as the type and **5.1** as the Runtime version.

### 2. Configure the Script
1.  Copy the content of the script from this repository.
2.  Paste it into the Runbook editor.
3.  **IMPORTANT:** You must update the variables at the top of the script to match your environment:
    * `$AlertThresholdPercentage`: Set your desired alert limit (e.g., `90`).
    * `$AdminRecipientEmail`: The email address that should receive the alerts.
    * `$SenderEmail`: The address the alert should come from.
4.  Save and **Publish** the Runbook.

### 3. Schedule Execution
1.  In the Runbook menu, go to **Schedules**.
2.  Add a schedule to run the script daily (e.g., at 6:00 AM).
3.  Link the schedule to the `Monitor-FullMailboxes` Runbook.

## Troubleshooting

* **"Command not found":** Ensure the `ExchangeOnlineManagement` module is fully imported and has a status of "Available" in the Automation Account.
* **"Access Denied":** Verify that the Managed Identity has been assigned the correct role in Entra ID and that replication has occurred (this can take 15-30 minutes).

---
*Disclaimer: This script is provided as-is. Always test in a non-production environment before full deployment.*
