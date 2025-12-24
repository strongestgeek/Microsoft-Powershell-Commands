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

# 3. Gather Data
$tenantSkus = Get-MgSubscribedSku -All
$allUsers = Get-MgUser -All -Property UserPrincipalName, DisplayName, AssignedLicenses, UserType
$sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

# Fast Lookups
$licenseLookup = @{}
foreach ($u in $allUsers) { $licenseLookup[$u.UserPrincipalName.ToLower()] = $u }
$sharedSet = New-Object System.Collections.Generic.HashSet[string] ([string[]]$sharedMailboxes.UserPrincipalName, [System.StringComparer]::OrdinalIgnoreCase)

# 4. Process
$results = foreach ($u in $allUsers) {
    if ($null -eq $u.AssignedLicenses -or $u.UserType -ne 'Member') { continue }
    
    $upn = $u.UserPrincipalName
    $isShared = $sharedSet.Contains($upn)
    
    $friendlyLics = foreach ($lic in $u.AssignedLicenses) {
        $sku = $tenantSkus | Where-Object { $_.SkuId -eq $lic.SkuId }
        if ($sku) { $skuMap[$sku.SkuPartNumber] ?? $sku.SkuPartNumber }
    }

    $sizeGB = 0; $hasArchive = "No"
    $m = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
    if ($m) {
        $hasArchive = if ($m.ArchiveGuid -ne [Guid]::Empty) { "Yes" } else { "No" }
        $stats = Get-MailboxStatistics -Identity $upn -ErrorAction SilentlyContinue
        # THE FIX: Cast to [int64] to prevent op_Division error
        if ($null -ne $stats.TotalItemSize -and $null -ne $stats.TotalItemSize.Value) {
            $sizeGB = [math]::Round(([int64]$stats.TotalItemSize.Value / 1GB), 2)
        }
    }

    $status = "OK"; $notes = ""
    if ($isShared) {
        if ($sizeGB -gt 50 -and $friendlyLics -notmatch "Plan 2|E3|E5") { $status = "Action Required"; $notes = "Shared > 50GB needs License" }
        elseif ($friendlyLics -match "E3|E5|Business Premium") { $status = "Optimization"; $notes = "Remove full license from Shared" }
    } else {
        if ($friendlyLics -contains "M365 E5" -and $friendlyLics -contains "O365 E3") { $status = "Redundant"; $notes = "Remove E3" }
        elseif ($friendlyLics -contains "M365 E3 (No Teams)" -and $friendlyLics -notmatch "Teams") { $status = "Warning"; $notes = "Missing Teams" }
    }

    [PSCustomObject]@{
        User = $upn; Name = $u.DisplayName; Type = if($isShared){"Shared"}else{"User"}
        Licenses = $friendlyLics -join ", "; SizeGB = $sizeGB; Archive = $hasArchive
        Status = $status; Notes = $notes
    }
}

# 5. Export
$results | Export-Csv -Path ".\M365_Audit_Report.csv" -NoTypeInformation
Write-Host "Done! Report saved to M365_Audit_Report.csv" -ForegroundColor Cyan