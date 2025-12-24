# 1. Connect to Services
Connect-ExchangeOnline
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"

# 2. Setup License Mapping
$skuMap = @{
    "AAD_PREMIUM"           = "Entra ID P1"; "AAD_PREMIUM_P2" = "Entra ID P2"
    "ENTERPRISEPREMIUM"     = "M365 E5"; "ENTERPRISEPACK"     = "O365 E3"
    "M365_BUSINESS_PREMIUM" = "Business Premium"; "SPE_E3_NOPSTN" = "M365 E3 (No Teams)"
}

# 3. Gather Data
$tenantSkus = Get-MgSubscribedSku -All
$allUsers = Get-MgUser -All -Property UserPrincipalName, DisplayName, AssignedLicenses, UserType
$sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

# Create a fast lookup for user licenses
$licenseLookup = @{}
foreach ($u in $allUsers) { $licenseLookup[$u.UserPrincipalName.ToLower()] = $u }

# Create a fast lookup to identify shared mailboxes
$sharedSet = New-Object System.Collections.Generic.HashSet[string] ([string[]]$sharedMailboxes.UserPrincipalName, [System.StringComparer]::OrdinalIgnoreCase)

# 4. Process Everything
$results = foreach ($u in $allUsers) {
    # Skip guests or things that aren't licensed
    if ($null -eq $u.AssignedLicenses -or $u.UserType -ne 'Member') { continue }
    
    $isShared = $sharedSet.Contains($u.UserPrincipalName)
    $upn = $u.UserPrincipalName
    
    # Get Friendly License Names
    $friendlyLics = foreach ($lic in $u.AssignedLicenses) {
        $sku = $tenantSkus | Where-Object { $_.SkuId -eq $lic.SkuId }
        if ($sku) { $skuMap[$sku.SkuPartNumber] ?? $sku.SkuPartNumber }
    }

    # Get Mailbox Size & Archive Status
    $sizeGB = 0; $hasArchive = "No"
    $m = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
    if ($m) {
        $hasArchive = if ($m.ArchiveGuid -ne [Guid]::Empty) { "Yes" } else { "No" }
        $stats = Get-MailboxStatistics -Identity $upn -ErrorAction SilentlyContinue
        if ($stats.TotalItemSize.Value) { $sizeGB = [math]::Round($stats.TotalItemSize.Value.ToBytes() / 1GB, 2) }
    }

    # Audit Logic
    $status = "OK"; $notes = ""
    if ($isShared) {
        if ($sizeGB -gt 50 -and $friendlyLics -notmatch "Plan 2|E3|E5") { $status = "Action Required"; $notes = "Shared > 50GB needs License" }
        if ($friendlyLics -match "E3|E5|Business Premium") { $status = "Optimization"; $notes = "Remove full license from Shared" }
    }

    [PSCustomObject]@{
        User        = $upn
        DisplayName = $u.DisplayName
        Type        = if ($isShared) { "Shared" } else { "User" }
        Licenses    = $friendlyLics -join ", "
        SizeGB      = $sizeGB
        Archive     = $hasArchive
        Status      = $status
        Notes       = $notes
    }
}


    # UPDATED USER ACCOUNT LOGIC
    $hasE5 = $friendlyLics -contains "M365 E5"
    $hasE3 = $friendlyLics -contains "O365 E3" -or $friendlyLics -contains "M365 E3 (No Teams)"
    
    # 1. Redundant Bundle Check
    if ($hasE5 -and $hasE3) {
        $status = "Redundant"; $notes = "Remove E3 (E5 covers everything)"
    }
    # 2. Add-on Overlap Check (e.g., Entra ID P1/P2 or Intune)
    elseif (($hasE5 -or $hasE3) -and ($friendlyLics -match "Entra ID|Intune|AIP")) {
        $status = "Optimization"; $notes = "Check for redundant standalone Security/Identity add-ons"
    }
    # 3. Teams Essentials Redundancy
    elseif ($friendlyLics -contains "Microsoft Teams Essentials" -and ($friendlyLics -match "Business|E3|E5")) {
        $status = "Redundant"; $notes = "Teams Essentials is included in the main bundle"
    }



# 5. Export and Close
$results | Export-Csv -Path ".\M365_Audit.csv" -NoTypeInformation
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
Write-Host "Done! Report saved to M365_Audit.csv" -ForegroundColor Cyan
