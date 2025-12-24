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

# Optimization: HashSets and HashTables for speed with 1,200+ items
$licenseLookup = @{}
foreach ($u in $allUsers) { $licenseLookup[$u.UserPrincipalName.ToLower()] = $u }
$sharedSet = New-Object System.Collections.Generic.HashSet[string] ([string[]]$sharedMailboxes.UserPrincipalName, [System.StringComparer]::OrdinalIgnoreCase)

# 4. Process Everything
$results = foreach ($u in $allUsers) {
    if ($null -eq $u.AssignedLicenses -or $u.UserType -ne 'Member') { continue }
    
    $upn = $u.UserPrincipalName
    $isShared = $sharedSet.Contains($upn)
    
    # Map Friendly Names
    $friendlyLics = foreach ($lic in $u.AssignedLicenses) {
        $sku = $tenantSkus | Where-Object { $_.SkuId -eq $lic.SkuId }
        if ($sku) { $skuMap[$sku.SkuPartNumber] ?? $sku.SkuPartNumber }
    }

    # Get Mailbox Size & Archive
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
        elseif ($friendlyLics -match "E3|E5|Business Premium") { $status = "Optimization"; $notes = "Remove full license from Shared" }
    } else {
        # USER ACCOUNT LOGIC
        if ($friendlyLics -contains "M365 E5" -and $friendlyLics -contains "O365 E3") {
            $status = "Redundant"; $notes = "User has E5 and E3; remove E3 to save cost"
        }
        elseif ($friendlyLics -contains "M365 E3 (No Teams)" -and $friendlyLics -notmatch "Teams") {
            $status = "Warning"; $notes = "Missing Teams license for E3 (No Teams) SKU"
        }
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

# 5. Export and Close
$results | Export-Csv -Path ".\M365_Audit.csv" -NoTypeInformation
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
Write-Host "Done! Report saved to M365_Audit.csv" -ForegroundColor Cyan
