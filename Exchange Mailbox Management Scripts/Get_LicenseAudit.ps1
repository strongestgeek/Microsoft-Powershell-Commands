# 1. Connect to Services
Connect-ExchangeOnline
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"

# 2. Setup License Mapping
$skuMap = @{
    "AAD_PREMIUM"             = "Microsoft Entra ID P1"
    "AAD_PREMIUM_P2"          = "Microsoft Entra ID P2"
    "ENTERPRISEPREMIUM"       = "Microsoft 365 E5"
    "ENTERPRISEPACK"          = "Office 365 E3"
    "STANDARDPACK"            = "Office 365 E1"
    "DESKLESSPACK"            = "Microsoft 365 F3"
    "M365_BUSINESS_PREMIUM"   = "Microsoft 365 Business Premium"
    "M365_BUSINESS_STANDARD"  = "Microsoft 365 Business Standard"
    "SPE_E3_NOPSTN"           = "Microsoft 365 E3 (no Teams)"
    "EMS"                     = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"              = "Enterprise Mobility + Security E5"
    "OFFICESUBSCRIPTION"      = "Microsoft 365 Apps for Enterprise"
}

# 3. Gather Data
Write-Host "Gathering tenant data..." -ForegroundColor Yellow
$tenantSkus = Get-MgSubscribedSku -All
$allUsers = Get-MgUser -All -Property UserPrincipalName, DisplayName, AssignedLicenses, UserType
$sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

# Fast Lookups
$upnArray = [string[]]($sharedMailboxes.UserPrincipalName)
$sharedSet = New-Object System.Collections.Generic.HashSet[string]($upnArray, [System.StringComparer]::OrdinalIgnoreCase)

# 4. Process Users
$total = $allUsers.Count
$counter = 0

$results = foreach ($u in $allUsers) {
    $counter++
    
    # Skip unlicensed or guest users
    if ($null -eq $u.AssignedLicenses -or $u.AssignedLicenses.Count -eq 0 -or $u.UserType -ne 'Member') { 
        continue 
    }

    $upn = $u.UserPrincipalName
    Write-Progress -Activity "Auditing Microsoft 365 Licenses" -Status "Processing $upn" -PercentComplete (($counter / $total) * 100)

    $isShared = $sharedSet.Contains($upn)

    # Map SKUs to friendly names
    $friendlyLics = foreach ($lic in $u.AssignedLicenses) {
        $sku = $tenantSkus | Where-Object { $_.SkuId -eq $lic.SkuId }
        if ($sku) { 
            if ($skuMap.ContainsKey($sku.SkuPartNumber)) { 
                $skuMap[$sku.SkuPartNumber] 
            } else { 
                $sku.SkuPartNumber 
            }
        }
    }

    # Get mailbox details
    $sizeGB = 0
    $hasArchive = "No"
    
    $m = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
    if ($m) {
        # Check for archive
        $hasArchive = if ($m.ArchiveGuid -ne [Guid]::Empty -and $null -ne $m.ArchiveGuid) { "Yes" } else { "No" }
        
        # Get mailbox size - handle deserialized objects properly
        $stats = Get-MailboxStatistics -Identity $upn -ErrorAction SilentlyContinue
        if ($null -ne $stats -and $null -ne $stats.TotalItemSize) {
            try {
                # Convert to string, extract bytes, calculate GB
                $sizeString = $stats.TotalItemSize.Value.ToString()
                # Match pattern like "1.234 GB (1,234,567,890 bytes)"
                if ($sizeString -match '\(([0-9,]+) bytes\)') {
                    $bytes = [int64]($matches[1] -replace ',', '')
                    $sizeGB = [math]::Round(($bytes / 1GB), 2)
                }
            } catch {
                Write-Warning "Could not parse size for $upn"
            }
        }
    }

    # Audit Logic
    $status = "OK"
    $notes = ""
    
    if ($isShared) {
        if ($sizeGB -gt 50 -and $friendlyLics -notmatch "Plan 2|E3|E5") { 
            $status = "Action Required"
            $notes = "Shared mailbox >50GB needs Plan 2 or E3/E5 license" 
        }
        elseif ($friendlyLics -match "E3|E5|Business Premium") { 
            $status = "Optimization"
            $notes = "Consider removing full license (shared mailbox doesn't require one under 50GB)" 
        }
    } else {
        if ($friendlyLics -contains "Microsoft 365 E5" -and $friendlyLics -contains "Office 365 E3") { 
            $status = "Redundant"
            $notes = "Remove E3 license (E5 includes all E3 features)" 
        }
        elseif ($friendlyLics -contains "Microsoft 365 E3 (no Teams)" -and $friendlyLics -notmatch "Teams") { 
            $status = "Warning"
            $notes = "User has E3 without Teams - verify Teams licensing" 
        }
    }

    # Output object
    [PSCustomObject]@{
        User     = $upn
        Name     = $u.DisplayName
        Type     = if ($isShared) { "Shared Mailbox" } else { "User Mailbox" }
        Licenses = ($friendlyLics | Select-Object -Unique) -join ", "
        SizeGB   = $sizeGB
        Archive  = $hasArchive
        Status   = $status
        Notes    = $notes
    }
}

Write-Progress -Activity "Auditing Microsoft 365 Licenses" -Completed

# 5. Export Results
$outputPath = ".\M365_Audit_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

# Summary
$summary = $results | Group-Object Status | Select-Object Name, Count
Write-Host "`nAudit Complete!" -ForegroundColor Green
Write-Host "Total objects processed: $($results.Count)" -ForegroundColor Cyan
Write-Host "`nStatus Summary:" -ForegroundColor Yellow
$summary | Format-Table -AutoSize
Write-Host "Report saved to: $outputPath" -ForegroundColor Cyan
