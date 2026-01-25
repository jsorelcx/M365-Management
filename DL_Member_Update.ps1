# --- Configuration ---
$DistributionListEmail = "DL@domain.com"

# Paste your array here exactly as requested
$NewUserList = @(
"user@domain.com",
"admin@domain.com"
)
# ---------------------

# 1. Connect to Exchange Online if not already connected
if (-not (Get-Module -Name ExchangeOnlineManagement)) {
    Write-Host "Please install the ExchangeOnlineManagement module." -ForegroundColor Red
    return
}
try {
    Get-ExoMailbox -ResultSize 1 -ErrorAction SilentlyContinue | Out-Null
}
catch {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline
}

$ErrorActionPreference = "Stop"

try {
    Write-Host "`n--- STARTING SYNC FOR: $DistributionListEmail ---" -ForegroundColor Cyan

    # 1. PREPARE DATA
    # We simply trim the existing array to ensure no accidental spaces exist
    $TargetMembers = $NewUserList | ForEach-Object { $_.Trim() }
    
    # Get the current members of the DL
    Write-Host "Fetching current members..." -ForegroundColor Gray
    $CurrentMembers = Get-DistributionGroupMember -Identity $DistributionListEmail -ResultSize Unlimited
    
    # Create a list of just the email addresses for easy comparison
    $CurrentMemberEmails = $CurrentMembers.PrimarySmtpAddress

    # ---------------------------------------------------------
    # 2. STEP ONE: REMOVE MEMBERS
    # Look at existing members -> If NOT in the new array -> Remove
    # ---------------------------------------------------------
    Write-Host "`n[PHASE 1] Checking for members to REMOVE..." -ForegroundColor Magenta
    
    if ($CurrentMembers) {
        foreach ($Member in $CurrentMembers) {
            $MemberEmail = $Member.PrimarySmtpAddress

            if ($MemberEmail -notin $TargetMembers) {
                # Member exists currently but is NOT in your new array
                Write-Host "[-] REMOVING: $MemberEmail (Not found in new list)" -ForegroundColor Red
                Remove-DistributionGroupMember -Identity $DistributionListEmail -Member $Member.Identity -Confirm:$false
            }
            else {
                # Member exists and IS in the new array
                Write-Host "[=] KEEPING:  $MemberEmail (Match found)" -ForegroundColor DarkGray
            }
        }
    }
    else {
        Write-Host "Current list is already empty. Nothing to remove." -ForegroundColor Yellow
    }

    # ---------------------------------------------------------
    # 3. STEP TWO: ADD MEMBERS
    # Look at new array -> If NOT in existing members -> Add
    # ---------------------------------------------------------
    Write-Host "`n[PHASE 2] Checking for members to ADD..." -ForegroundColor Magenta

    foreach ($NewEmail in $TargetMembers) {
        # Check if this email was in our original snapshot of current members
        if ($NewEmail -notin $CurrentMemberEmails) {
            # It wasn't there before, so we add it
            Write-Host "[+] ADDING:   $NewEmail (Missing from current list)" -ForegroundColor Green
            try {
                Add-DistributionGroupMember -Identity $DistributionListEmail -Member $NewEmail
            }
            catch {
                Write-Host "    ! ERROR adding $NewEmail : $_" -ForegroundColor Red
            }
        }
        else {
            # It was already there. 
            Write-Host "[v] VERIFIED: $NewEmail is already a member." -ForegroundColor DarkGray
        }
    }

    Write-Host "`n--- SYNC COMPLETE ---" -ForegroundColor Cyan
}
catch {
    Write-Host "`nFATAL ERROR: $_" -ForegroundColor Red
}
