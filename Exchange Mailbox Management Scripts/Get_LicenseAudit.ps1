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

# 5. Export and Close
$results | Export-Csv -Path ".\M365_Audit.csv" -NoTypeInformation
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
Write-Host "Done! Report saved to M365_Audit.csv" -ForegroundColor Cyan
