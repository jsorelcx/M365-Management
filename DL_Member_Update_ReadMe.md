# M365 Distribution List Member Sync

A PowerShell script to safely synchronize Microsoft 365 Distribution List members against a provided list of email addresses. 

Unlike simple "wipe and replace" scripts, this tool performs a **differential sync**:
1. It looks up current members and **removes** only those who are not in your new list.
2. It looks at your new list and **adds** only those who are missing from the DL.
3. It preserves existing members (and their object links) if they are already in the list.

## Features
- **Safe Sync Logic:** Reduces the risk of breaking permissions or object links by only modifying differences.
- **Verbose Logging:** Color-coded console output shows exactly who is being added `[+]`, removed `[-]`, or verified as existing `[v]`.
- **Duplicate Protection:** Automatically handles whitespace trimming and checks for existence before adding.

## Prerequisites
1. **PowerShell 5.1** or later (or PowerShell Core).
2. **Exchange Online PowerShell Module**:
   ```powershell
   Install-Module -Name ExchangeOnlineManagement
   ```
3. **Permissions:** You must be a Global Admin or Exchange Administrator in the target M365 tenant.

## Usage

1. Open the script file (e.g., `Sync-DLMembers.ps1`).
2. Edit the **Configuration** section at the top of the file:
   - Set `$DistributionListEmail` to the email address of the Distribution List you want to modify.
   - Paste your list of users into the `$NewUserList` array.

```powershell
# --- Configuration ---
$DistributionListEmail = "allstaff@yourdomain.com"

$NewUserList = @(
    "john.doe@yourdomain.com",
    "jane.smith@yourdomain.com",
    "bob.jones@yourdomain.com"
)
# ---------------------
```

3. Run the script. It will prompt for M365 credentials if you are not already connected.

## Excel Helper: Generating the Array
If you have a list of emails in Excel (e.g., Column A) and want to format them for the PowerShell array, use this formula to wrap them in quotes and add a comma:

**For a single cell (A1):**
```excel
=CHAR(34) & A1 & CHAR(34) & ","
```

**For a dynamic list (Spill array in Office 365):**
Use this formula to generate the entire formatted block in one cell, which you can copy/paste directly between the `@(` and `)` in the script:
```excel
=TEXTJOIN(CHAR(10), TRUE, CHAR(34) & A1:A100 & CHAR(34) & ",")
```

## How It Works
The script follows a two-phase process:

1.  **Phase 1 (Cleanup):** It fetches the *current* members of the DL. It iterates through them and checks if they exist in your local `$NewUserList`. If they are missing from your local list, they are removed from the cloud DL.
2.  **Phase 2 (Addition):** It iterates through your local `$NewUserList`. It checks if the user exists in the *original snapshot* of the DL. If they are missing, they are added.

## License
This script is provided "as is" without warranty of any kind. Please test in a non-production environment before running on critical distribution lists.
