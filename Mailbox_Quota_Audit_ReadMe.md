# M365-Management: Mailbox Quota Audit & Full-Mailbox Alerting (Exchange Online)

This PowerShell runbook is designed to run in Azure Automation. It scans Exchange Online mailboxes, calculates storage usage as a percentage of the configured quota, and sends an email alert when a mailbox exceeds a defined threshold (default 90%).

* **No credentials are stored:** authentication is performed using a System-Assigned Managed Identity.
* **The notification email is sent using Microsoft Graph (`sendMail`).**

## Overview

* **Platform:** Azure Automation (Runbook / Cloud Job)
* **Language:** Windows PowerShell 5.1
* **Exchange authentication:** `Connect-ExchangeOnline -ManagedIdentity`
* **Email delivery:** Microsoft Graph `POST /v1.0/users/{id}/sendMail`
* **Mailbox types scanned:** `UserMailbox` + `SharedMailbox`
* **Matching method:** Statistics are keyed by `MailboxGuid` / `ExchangeGuid` for reliability

## What the Script Does

1.  **Connects** to Exchange Online using the Automation Account’s Managed Identity.
2.  **Retrieves** all mailboxes, excluding those with `DisplayName` beginning with `RESIGNED`.
3.  **Collects** mailbox statistics in batches (to avoid throttling), keyed by GUID.
4.  **Parses** size + quota values into GB, supporting MB, GB, and TB.
5.  **Calculates** % used and flags mailboxes at/above the threshold.
6.  **If any are flagged, it:**
    * Gets a Graph access token using Managed Identity.
    * Sends an HTML report email listing mailboxes above threshold.

## Configuration (Required)

Update these values at the top of the script:

```powershell
$AlertThresholdPercentage = 90
$AdminRecipientEmail      = "Insert mailbox to receive alert"
$SenderEmail              = "Insert mailbox to appear as sender" # Must be a valid mailbox
```

### Notes
* `$SenderEmail` must be a real mailbox in Exchange Online (user or shared) because Graph is sending “as that user”.
* `$AdminRecipientEmail` can be any recipient mailbox you want to receive the alert.
* The script uses `-Organization "insert tenant primary domain"` in `Connect-ExchangeOnline`. That value should be in quotes because it’s a string.

**Example:**
```powershell
Connect-ExchangeOnline -ManagedIdentity -Organization "contoso.onmicrosoft.com" -ShowBanner:$false
```

## Prerequisites

### 1) Azure Automation Account
A standard Azure Automation Account.

### 2) PowerShell Module
Import the **ExchangeOnlineManagement** module into the Automation Account.
* **Recommended:** latest stable version available in Automation.
* **Required cmdlets:**
    * `Connect-ExchangeOnline`
    * `Get-EXOMailbox`
    * `Get-EXOMailboxStatistics`

### 3) System-Assigned Managed Identity
Enable **System Assigned Managed Identity** on the Automation Account.

### 4) Permissions (Exchange)
The Managed Identity must be granted rights to read mailbox + statistics.
* **Minimum practical approach:** assign the Managed Identity an Exchange/Entra role that permits mailbox stats queries.
* **Common option:** Exchange Administrator (broad).
* **Least-privilege:** Use a role assignment that allows `Get-EXOMailbox` and `Get-EXOMailboxStatistics` for your scope.
* *Permission changes can take time to propagate (often 10–30 minutes).*

### 5) Permissions (Graph)
The Managed Identity must be allowed to send mail using Graph as the sender mailbox.
The script calls:
`POST https://graph.microsoft.com/v1.0/users/{senderObjectId}/sendMail`

Ensure the identity has the appropriate Graph permissions/consent for sending mail in your tenant context.

## Deployment Steps

### Step 1 — Create the Runbook
1.  Navigate to **Azure Portal** → **Automation Accounts** → **[Your Account]** → **Runbooks**.
2.  Create a runbook:
    * **Name:** `MailboxQuotaAudit` (or your preferred name)
    * **Runbook type:** PowerShell
    * **Runtime version:** 5.1

### Step 2 — Paste the Script
1.  Open the runbook editor.
2.  Paste the full script content from this repository.
3.  Update the configuration variables at the top:
    * `$AlertThresholdPercentage`
    * `$AdminRecipientEmail`
    * `$SenderEmail`
    * `-Organization "insert tenant primary domain"`

### Step 3 — Publish
Click **Save** → **Publish**.

### Step 4 — Test
Click **Start** and observe the output. You should see:
* Mailboxes found
* Batch retrieval progress
* Parsed/skipped counts
* **If exceeded:** a Graph `sendMail` URI + token length + confirmation.

### Step 5 — Schedule Execution
1.  Go to **Runbook** → **Schedules**.
2.  Add a schedule (e.g., daily at 06:00).
3.  Link the schedule to the runbook.

## Output & Interpretation

The runbook prints progress and a summary like:

```text
Parsed (size+quota): <count>
Skipped (no guid): <count> | (no stats): <count> | (Under 1MB content): <count> | (bad quota): <count>
Max usage percent seen: <value>%
```

### Skipped: “Under 1MB content”
These are mailboxes whose `TotalItemSize` returns a value below MB, e.g., in KB (very small / nearly empty mailboxes). This is expected and acceptable if you only care about mailboxes approaching quota.

## Troubleshooting

* **“Command not found”**
    * Ensure `ExchangeOnlineManagement` is imported and shows as **Available** in Azure Automation modules.
* **“Access Denied” / “Insufficient privileges”**
    * Verify the Managed Identity has appropriate Exchange permissions to read mailbox + stats.
    * Verify the Managed Identity has appropriate Graph permissions to send mail.
    * Allow time for role assignment propagation.
* **Token errors / 401 Unauthorized**
    * Usually indicates the token request failed or Graph permissions are missing. The script checks token validity (length) before sending the request.
* **Throttling / transient failures**
    * Mailbox statistics are fetched in batches with retries. Batch size defaults to 200 with retries up to 3 times with increasing delay. You can tune this via `$batchSize = 200`.

## Security Notes

* **No secrets are stored** in the script.
* Identity authentication is performed via **Managed Identity endpoints**.
* Ensure the sender mailbox and recipient mailbox align with your security policies.

---

### Disclaimer
*This script is provided as-is. Always test in a non-production environment first, and confirm permissions align with your organization’s governance standards.*
