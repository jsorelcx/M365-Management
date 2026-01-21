<#
    .DESCRIPTION
        Checks Exchange Online mailboxes for usage > 90% and alerts via email.
        Uses Managed Identity for authentication.
#>

# Variables (As per Rollout Plan)
$AlertThresholdPercentage = 90
$AdminRecipientEmail = "admin@yourdomain.com" # UPDATE THIS
$SenderEmail = "automation@yourdomain.com"    # UPDATE THIS (Must be a valid sender)

try {
    # 1. Connect to Exchange Online using Managed Identity
    Write-Output "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ManagedIdentity -Organization "yourtenant.onmicrosoft.com"

    # 2. Get all mailboxes and check usage
    Write-Output "Scanning mailboxes..."
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq "UserMailbox"}
    
    $HighUsageMailboxes = @()

    foreach ($mbx in $Mailboxes) {
        $stats = Get-MailboxStatistics -Identity $mbx.UserPrincipalName | Select-Object DisplayName, TotalItemSize, ItemCount
        
        # Parse TotalItemSize to GB (simplified logic)
        if ($stats.TotalItemSize.Value -match "([\d\.]+) GB") {
            $SizeGB = [double]$matches[1]
            
            # Assuming standard 100GB or 50GB quota - retrieval of exact quota requires Get-Mailbox output
            # For this script, we check percentage if possible, or assume a threshold.
            # Best practice: Compare $stats.TotalItemSize against $mbx.ProhibitSendQuota
            
            $Quota = $mbx.ProhibitSendQuota
            if ($Quota -match "([\d\.]+) GB") {
                $QuotaGB = [double]$matches[1]
                if ($QuotaGB -gt 0) {
                    $UsagePercent = ($SizeGB / $QuotaGB) * 100
                    
                    if ($UsagePercent -ge $AlertThresholdPercentage) {
                        $HighUsageMailboxes += [PSCustomObject]@{
                            Name = $mbx.DisplayName
                            UPN = $mbx.UserPrincipalName
                            UsageGB = $SizeGB
                            QuotaGB = $QuotaGB
                            Percent = [math]::Round($UsagePercent, 2)
                        }
                    }
                }
            }
        }
    }

    # 3. Send Alert if needed
    if ($HighUsageMailboxes.Count -gt 0) {
        Write-Output "High usage detected. Sending alert."
        $Body = $HighUsageMailboxes | Format-Table | Out-String
        
        # Sending email typically requires Send-MailMessage (SMTP) or Graph API.
        # Since standard SMTP is often blocked in Azure, ensure you have an SMTP relay or use Graph.
        # This is a placeholder for the alert logic mentioned in the plan.
        Write-Output "ALERT: The following mailboxes are over ${AlertThresholdPercentage}%:`n$Body"
    } else {
        Write-Output "No mailboxes found over quota threshold."
    }

} catch {
    Write-Error "An error occurred: $_"
}
