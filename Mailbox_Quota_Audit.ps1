$AlertThresholdPercentage = 90
$AdminRecipientEmail      = "recipient@domain.com"
$SenderEmail              = "Insert mailbox to appear as sender" # Must be valid mailbox
$Organization             = "<tenant.onmicrosoft.com or primary domain>"

$ErrorActionPreference = "Stop"
$VerbosePreference     = "Continue"   # set to "SilentlyContinue" once stable

function Get-ManagedIdentityToken {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Resource
    )

    # Azure Automation Managed Identity (IDENTITY_ENDPOINT + IDENTITY_HEADER)
    if ($env:IDENTITY_ENDPOINT -and $env:IDENTITY_HEADER) {

        $endpoint = $env:IDENTITY_ENDPOINT.Trim()
        $ub = [System.UriBuilder]$endpoint

        # Parse existing query (if any) into a hashtable
        $existing = @{}
        $rawQuery = ($ub.Query.TrimStart('?'))
        if ($rawQuery) {
            foreach ($pair in $rawQuery -split '&') {
                if (-not $pair) { continue }
                $kv = $pair -split '=', 2
                $k = [uri]::UnescapeDataString($kv[0])
                $v = if ($kv.Count -gt 1) { [uri]::UnescapeDataString($kv[1]) } else { "" }
                if ($k) { $existing[$k] = $v }
            }
        }

        # Force required parameters
        $existing["resource"]    = $Resource
        $existing["api-version"] = "2019-08-01"

        # Rebuild query string (escaped)
        $qs = ($existing.GetEnumerator() | ForEach-Object {
            "{0}={1}" -f [uri]::EscapeDataString($_.Key), [uri]::EscapeDataString([string]$_.Value)
        }) -join "&"

        $ub.Query = $qs
        $tokenUri = $ub.Uri.AbsoluteUri

        Write-Verbose "Token endpoint: $tokenUri"

        $resp = Invoke-RestMethod -Method GET -Uri $tokenUri -Headers @{
            "X-IDENTITY-HEADER" = $env:IDENTITY_HEADER
        } -ErrorAction Stop

        return $resp.access_token
    }

    # IMDS fallback
    $imdsUri = "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=$([uri]::EscapeDataString($Resource))"
    Write-Verbose "Token endpoint (IMDS): $imdsUri"

    $resp = Invoke-RestMethod -Method GET -Uri $imdsUri -Headers @{ Metadata="true" } -ErrorAction Stop
    return $resp.access_token
}

function Parse-SizeToGB {
    param([Parameter(Mandatory=$true)] $Value)

    $text = $Value.ToString()

    # Accepts: "12.34 GB (...)" / "850 MB (...)" / "1.2 TB (...)" / "12 GB"
    if ($text -match '([\d\.,]+)\s*(TB|GB|MB)') {
        $num  = $matches[1]
        $unit = $matches[2]

        if ($num.Contains(',') -and -not $num.Contains('.')) {
            $num = $num.Replace(',', '.')
        } else {
            $num = $num.Replace(',', '')
        }

        $val = [double]::Parse($num, [System.Globalization.CultureInfo]::InvariantCulture)

        switch ($unit) {
            'TB' { return $val * 1024.0 }
            'GB' { return $val }
            'MB' { return $val / 1024.0 }
        }
    }

    return $null
}

try {
    Write-Output "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ManagedIdentity -Organization $Organization -ShowBanner:$false

    Write-Output "Retrieving mailbox list (excluding 'RESIGNED' users)..."
    $Mailboxes = Get-EXOMailbox -ResultSize Unlimited `
        -RecipientTypeDetails UserMailbox,SharedMailbox `
        -Properties ProhibitSendQuota,ExchangeGuid `
        -Filter "DisplayName -notlike 'RESIGNED*'"

    $TotalCount = $Mailboxes.Count
    Write-Output "Found $TotalCount mailboxes to scan."

    Write-Output "Retrieving mailbox statistics in batches (keyed by GUID)..."
    $StatsByGuid = @{}

    $ids = @(
        $Mailboxes |
        Where-Object { $_.ExchangeGuid -and $_.ExchangeGuid -ne [guid]::Empty } |
        ForEach-Object { $_.ExchangeGuid.ToString() }
    )

    $batchSize    = 200
    $totalBatches = [math]::Ceiling($ids.Count / $batchSize)

    for ($b = 0; $b -lt $totalBatches; $b++) {
        $start = $b * $batchSize
        $end   = [math]::Min($start + $batchSize - 1, $ids.Count - 1)
        $batch = $ids[$start..$end]

        Write-Output "Fetching stats batch $($b+1)/$totalBatches ($($batch.Count) mailboxes)..."

        $tries = 0
        while ($true) {
            try {
                $batchStats = $batch | Get-EXOMailboxStatistics -Properties TotalItemSize
                break
            } catch {
                $tries++
                if ($tries -ge 4) { throw }
                Start-Sleep -Seconds (5 * $tries)
            }
        }

        foreach ($s in $batchStats) {
            if ($s.MailboxGuid) {
                $StatsByGuid[$s.MailboxGuid.ToString()] = $s
            }
        }
    }

    Write-Output "Stats collected for $($StatsByGuid.Count) mailboxes."

    $HighUsageMailboxes = @()

    $Counter         = 0
    $ParsedCount     = 0
    $SkippedNoGuid   = 0
    $SkippedNoStats  = 0
    $SkippedBadSize  = 0
    $SkippedBadQuota = 0
    $MaxPercent      = 0.0

    foreach ($mbx in $Mailboxes) {
        $Counter++
        if ($Counter % 200 -eq 0) {
            Write-Output "Processed $Counter of $TotalCount mailboxes..."
        }

        if (-not $mbx.ExchangeGuid -or $mbx.ExchangeGuid -eq [guid]::Empty) { $SkippedNoGuid++; continue }

        $guidKey = $mbx.ExchangeGuid.ToString()
        $stats   = $StatsByGuid[$guidKey]
        if (-not $stats) { $SkippedNoStats++; continue }

        $SizeGB = Parse-SizeToGB -Value $stats.TotalItemSize
        if ($null -eq $SizeGB) { $SkippedBadSize++; continue }

        $QuotaText = $mbx.ProhibitSendQuota.ToString()
        if ($QuotaText -eq "Unlimited") {
            $QuotaGB = 100.0
        } else {
            $QuotaGB = Parse-SizeToGB -Value $mbx.ProhibitSendQuota
        }

        if ($null -eq $QuotaGB -or $QuotaGB -le 0) { $SkippedBadQuota++; continue }

        $ParsedCount++
        $UsagePercent = ($SizeGB / $QuotaGB) * 100.0
        if ($UsagePercent -gt $MaxPercent) { $MaxPercent = $UsagePercent }

        if ($UsagePercent -ge $AlertThresholdPercentage) {
            $HighUsageMailboxes += [PSCustomObject]@{
                Name    = $mbx.DisplayName
                UPN     = $mbx.UserPrincipalName
                UsageGB = [math]::Round($SizeGB, 2)
                QuotaGB = [math]::Round($QuotaGB, 2)
                Percent = [math]::Round($UsagePercent, 2)
            }
        }
    }

    Write-Output "Parsed (size+quota): $ParsedCount"
    Write-Output "Skipped (no guid): $SkippedNoGuid | (no stats): $SkippedNoStats | (nparsed size): $SkippedBadSize | (bad quota): $SkippedBadQuota"
    Write-Output ("Max usage percent seen: {0:N2}%" -f $MaxPercent)

    if ($HighUsageMailboxes.Count -gt 0) {
        Write-Output "High usage detected ($($HighUsageMailboxes.Count) mailboxes). Sending email..."

        $HighUsageMailboxes = $HighUsageMailboxes | Sort-Object Percent -Descending

        $TableRows = ""
        foreach ($row in $HighUsageMailboxes) {
            $TableRows += "<tr><td>$($row.Name)</td><td>$($row.UPN)</td><td>$($row.UsageGB) GB</td><td>$($row.QuotaGB) GB</td><td><b>$($row.Percent)%</b></td></tr>"
        }

        $HtmlBody = @"
<h3>Mailbox Storage Alert</h3>
<p>The following mailboxes have exceeded the <b>$AlertThresholdPercentage%</b> capacity threshold.</p>
<table border='1' cellpadding='5' style='border-collapse:collapse;'>
<tr><th>Name</th><th>Email</th><th>Usage</th><th>Quota</th><th>% Full</th></tr>
$TableRows
</table>
"@

        $payload = @{
            message = @{
                subject = "Alert: Mailbox Storage Capacity Report"
                body = @{ contentType = "HTML"; content = $HtmlBody }
                toRecipients = @(@{ emailAddress = @{ address = $AdminRecipientEmail } })
            }
            saveToSentItems = $false
        } | ConvertTo-Json -Depth 10

        $senderMbx     = Get-EXOMailbox -Identity $SenderEmail -Properties ExternalDirectoryObjectId
        $senderObjectId = $senderMbx.ExternalDirectoryObjectId
        if (-not $senderObjectId) { throw "Could not resolve ExternalDirectoryObjectId for sender '$SenderEmail'" }

        $sendMailUri = "https://graph.microsoft.com/v1.0/users/$senderObjectId/sendMail"
        Write-Verbose "Graph sendMail URI: $sendMailUri"

        $token = Get-ManagedIdentityToken -Resource "https://graph.microsoft.com/"
        if (-not ($token -is [string]) -or $token.Length -lt 200) {
            throw "Managed Identity token looks invalid. Type=$($token.GetType().FullName) Length=$($token.Length)"
        }

        Write-Output "Token acquired (length=$($token.Length)). Sending POST..."

        Invoke-RestMethod -Method POST -Uri $sendMailUri -Headers @{
            Authorization  = "Bearer $token"
            "Content-Type" = "application/json"
        } -Body $payload

        Write-Output "Email sent successfully."
    }
    else {
        Write-Output "All mailboxes are within safe limits."
    }
}
catch {
    $resp = $_.Exception.Response
    if ($resp) {
        $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
        $body   = $reader.ReadToEnd()
        Write-Error "CRITICAL ERROR: $($_.Exception.Message) | Response: $body"
    }
    else {
        Write-Error "CRITICAL ERROR: $_"
    }
}
