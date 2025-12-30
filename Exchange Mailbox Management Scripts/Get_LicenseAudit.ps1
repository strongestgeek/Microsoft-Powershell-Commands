#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Microsoft 365 License Compliance Audit Script
.DESCRIPTION
    Audits M365 licenses against organizational policies:
    - Shared mailboxes: No license unless >50GB or has archive (then Exchange Plan 2)
    - Standard users: Microsoft 365 E3 + Teams + Defender Suite + Defender VM Add-on
    - Contractors: Exchange Plan 1 or Office 365 E1
.NOTES
    Version: 4.0
    Requires: ExchangeOnlineManagement, Microsoft.Graph modules
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = ".\M365_License_Compliance_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter()]
    [switch]$SkipSharedMailboxCheck,
    
    [Parameter()]
    [int]$SharedMailboxThresholdGB = 50
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Write-ColorOutput {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Connect-Services {
    try {
        Write-ColorOutput "Connecting to Exchange Online..." "Yellow"
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        Write-ColorOutput "Connecting to Microsoft Graph..." "Yellow"
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
        
        Write-ColorOutput "Successfully connected to all services`n" "Green"
    }
    catch {
        Write-ColorOutput "Failed to connect to services: $_" "Red"
        throw
    }
}

# License Mapping (Only licenses you care about)
$skuMap = @{
    # Exchange
    "EXCHANGESTANDARD"                    = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE"                  = "Exchange Online Plan 2"
    "EXCHANGEARCHIVE_ADDON"               = "Exchange Online Archiving"
    
    # Microsoft 365
    "SPE_E3"                              = "Microsoft 365 E3"
    "SPE_E3_NOPSTN"                       = "Microsoft 365 E3 (No Teams)"
    
    # Office 365
    "STANDARDPACK"                        = "Office 365 E1"
    
    # Defender
    "ATP_ENTERPRISE"                      = "Microsoft Defender for Office 365 Plan 1"
    "ATP_ENTERPRISE_FACULTY"              = "Microsoft Defender for Office 365 Plan 1"
    "THREAT_INTELLIGENCE"                 = "Microsoft Defender for Office 365 Plan 2"
    "DEFENDER_SUITE"                      = "Microsoft Defender Suite"
    "DEFENDER_ENDPOINT_P2"                = "Microsoft Defender for Endpoint P2"
    "MDATP_XPLAT"                         = "Microsoft Defender for Endpoint"
    "DEFENDER_VULNERABILITY_MANAGEMENT"   = "Microsoft Defender Vulnerability Management Add-on"
    
    # Teams
    "TEAMS_ENTERPRISE"                    = "Microsoft Teams Enterprise"
    "MCOSTANDARD"                         = "Microsoft Teams"
    "TEAMS1"                              = "Microsoft Teams"
    
    # Ignored (Visio, Project, etc.)
    "VISIOCLIENT"                         = "Visio Plan 2"
    "PROJECTPROFESSIONAL"                 = "Project Plan 3"
    "POWER_BI_PRO"                        = "Power BI Pro"
}

# Licenses to ignore in compliance checks
$ignoredLicenses = @(
    "Visio Plan 2",
    "Project Plan 3",
    "Power BI Pro"
)

# Expected licenses for different user types
$standardUserRequired = @{
    "Microsoft 365 E3 or E3 (No Teams)" = $true
    "Microsoft Defender Suite" = $true
    "Microsoft Defender Vulnerability Management Add-on" = $true
    "Microsoft Teams" = $true  # If they have E3 No Teams
}

$contractorAllowed = @(
    "Exchange Online Plan 1",
    "Office 365 E1"
)

function Get-MailboxSizeInGB {
    param([string]$UserPrincipalName)
    
    try {
        $stats = Get-MailboxStatistics -Identity $UserPrincipalName -ErrorAction SilentlyContinue
        if ($null -ne $stats -and $null -ne $stats.TotalItemSize) {
            $sizeString = $stats.TotalItemSize.Value.ToString()
            if ($sizeString -match '\(([0-9,]+) bytes\)') {
                $bytes = [int64]($matches[1] -replace ',', '')
                return [math]::Round(($bytes / 1GB), 2)
            }
        }
    }
    catch {
        Write-Verbose "Could not retrieve mailbox size for $UserPrincipalName"
    }
    return 0
}

function Get-ArchiveUsageInGB {
    param([string]$UserPrincipalName)
    
    try {
        $stats = Get-MailboxStatistics -Identity $UserPrincipalName -Archive -ErrorAction SilentlyContinue
        if ($null -ne $stats -and $null -ne $stats.TotalItemSize) {
            $sizeString = $stats.TotalItemSize.Value.ToString()
            if ($sizeString -match '\(([0-9,]+) bytes\)') {
                $bytes = [int64]($matches[1] -replace ',', '')
                return [math]::Round(($bytes / 1GB), 2)
            }
        }
    }
    catch {
        Write-Verbose "Could not retrieve archive size for $UserPrincipalName"
    }
    return 0
}

function Test-LicenseCompliance {
    param(
        [array]$Licenses,
        [bool]$IsSharedMailbox,
        [double]$MailboxSizeGB,
        [bool]$HasArchive,
        [string]$UserPrincipalName,
        [double]$ArchiveSizeGB,
        [bool]$LitigationHoldEnabled,
        [string]$LicenseAssignmentMethod
    )
    
    # Remove ignored licenses from compliance checks
    $relevantLicenses = $Licenses | Where-Object { $_ -notin $ignoredLicenses }
    
    $issues = @()
    $warnings = @()
    $status = "Compliant"
    
    # ===== LITIGATION HOLD CHECK (Critical) =====
    if ($LitigationHoldEnabled) {
        $warnings += "⚠ LITIGATION HOLD ENABLED - DO NOT REMOVE LICENSE"
    }
    
    # ===== DUPLICATE EXCHANGE LICENSE CHECK =====
    $hasStandaloneExchange = $relevantLicenses -match "^Exchange Online Plan"
    $hasM365Bundle = $relevantLicenses -match "Microsoft 365 E3"
    
    if ($hasStandaloneExchange -and $hasM365Bundle -and -not $IsSharedMailbox) {
        $status = "Non-Compliant"
        $issues += "Duplicate Exchange license detected (M365 E3 includes Exchange - remove standalone Exchange license)"
    }
    
    # ===== SHARED MAILBOX CHECKS =====
    if ($IsSharedMailbox) {
        $needsLicense = ($MailboxSizeGB -gt $SharedMailboxThresholdGB) -or ($HasArchive -eq "Yes")
        
        if ($needsLicense) {
            # Should have Exchange Plan 2
            if ($relevantLicenses -notcontains "Exchange Online Plan 2") {
                $status = "Non-Compliant"
                $reason = if ($HasArchive -eq "Yes") { "has archive" } else { "size >$($SharedMailboxThresholdGB)GB" }
                $issues += "Shared mailbox $reason but missing Exchange Online Plan 2 license"
            }
            
            # Should NOT have full E3 licenses
            if ($relevantLicenses -match "Microsoft 365 E3") {
                $status = "Non-Compliant"
                $issues += "Shared mailbox has full M365 E3 license (should use Exchange Plan 2 only)"
            }
            
            # Check if archive is being used
            if ($HasArchive -eq "Yes" -and $ArchiveSizeGB -lt 1) {
                $warnings += "Archive enabled but empty (<1GB) - consider if archive is needed"
            }
        }
        else {
            # Should NOT have any licenses
            if ($relevantLicenses.Count -gt 0) {
                $status = "Non-Compliant"
                $issues += "Shared mailbox under $($SharedMailboxThresholdGB)GB without archive should not have licenses (current: $($relevantLicenses -join ', '))"
            }
        }
        
        return @{
            Status = $status
            Issues = $issues
            Warnings = $warnings
            UserType = "Shared Mailbox"
            LicenseAssignment = $LicenseAssignmentMethod
        }
    }
    
    # ===== CONTRACTOR CHECKS =====
    # Detect contractors by their base license
    $isContractor = ($relevantLicenses -match "Exchange Online Plan 1|Office 365 E1") -and 
                    ($relevantLicenses -notmatch "Microsoft 365 E3")
    
    if ($isContractor) {
        $hasValidContractorLicense = $false
        foreach ($lic in $contractorAllowed) {
            if ($relevantLicenses -contains $lic) {
                $hasValidContractorLicense = $true
                break
            }
        }
        
        if (-not $hasValidContractorLicense) {
            $status = "Non-Compliant"
            $issues += "Contractor without valid license (Exchange Plan 1 or Office 365 E1)"
        }
        
        # Check if contractor needs upgrade (mailbox size)
        if ($MailboxSizeGB -gt 50 -and $relevantLicenses -contains "Exchange Online Plan 1") {
            $warnings += "Contractor mailbox >50GB with Exchange Plan 1 - consider upgrading to Plan 2"
        }
        
        return @{
            Status = $status
            Issues = $issues
            Warnings = $warnings
            UserType = "Contractor"
            LicenseAssignment = $LicenseAssignmentMethod
        }
    }
    
    # ===== STANDARD USER CHECKS =====
    # Check for Microsoft 365 E3 (with or without Teams)
    $hasE3 = $relevantLicenses -contains "Microsoft 365 E3"
    $hasE3NoTeams = $relevantLicenses -contains "Microsoft 365 E3 (No Teams)"
    
    if (-not $hasE3 -and -not $hasE3NoTeams) {
        $status = "Non-Compliant"
        $issues += "Missing Microsoft 365 E3 license"
    }
    
    # If they have E3 No Teams, they need the Teams add-on
    if ($hasE3NoTeams) {
        $hasTeamsAddon = $relevantLicenses -match "Microsoft Teams"
        if (-not $hasTeamsAddon) {
            $status = "Non-Compliant"
            $issues += "Has M365 E3 (No Teams) but missing Microsoft Teams add-on license"
        }
    }
    
    # Check for Microsoft Defender Suite
    $hasDefenderSuite = $relevantLicenses -contains "Microsoft Defender Suite"
    if (-not $hasDefenderSuite) {
        $status = "Non-Compliant"
        $issues += "Missing Microsoft Defender Suite"
    }
    
    # Check for Microsoft Defender Vulnerability Management Add-on
    $hasDefenderVM = $relevantLicenses -contains "Microsoft Defender Vulnerability Management Add-on"
    if (-not $hasDefenderVM) {
        $status = "Non-Compliant"
        $issues += "Missing Microsoft Defender Vulnerability Management Add-on"
    }
    
    # Check for Defender for Office 365 (email protection)
    $hasDefenderForOffice = $relevantLicenses -match "Microsoft Defender for Office 365"
    if (-not $hasDefenderForOffice) {
        $warnings += "No explicit Defender for Office 365 license - verify email protection (may be included in Defender Suite)"
    }
    
    # Check for duplicate base licenses
    if ($hasE3 -and $hasE3NoTeams) {
        $status = "Non-Compliant"
        $issues += "Has both M365 E3 and M365 E3 (No Teams) - remove one"
    }
    
    # Check if archive is being used
    if ($HasArchive -eq "Yes" -and $ArchiveSizeGB -lt 1) {
        $warnings += "Archive enabled but empty (<1GB) - consider if archive is needed"
    }
    
    # License assignment method warning
    if ($LicenseAssignmentMethod -eq "Direct") {
        $warnings += "License assigned directly (recommend group-based licensing for easier management)"
    }
    
    return @{
        Status = $status
        Issues = $issues
        Warnings = $warnings
        UserType = "Standard User"
        LicenseAssignment = $LicenseAssignmentMethod
    }
}

# Main Script Execution
try {
    Write-ColorOutput "`n=== Microsoft 365 License Compliance Audit ===" "Cyan"
    Write-ColorOutput "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" "Gray"
    
    Connect-Services
    
    # Gather tenant data
    Write-ColorOutput "Gathering tenant license information..." "Yellow"
    $tenantSkus = Get-MgSubscribedSku -All
    
    Write-ColorOutput "Retrieving all users..." "Yellow"
    $allUsers = Get-MgUser -All -Property UserPrincipalName, DisplayName, AssignedLicenses, UserType, AccountEnabled
    
    # Get shared mailboxes if not skipped
    $sharedSet = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    if (-not $SkipSharedMailboxCheck) {
        Write-ColorOutput "Identifying shared mailboxes..." "Yellow"
        $sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
        $upnArray = [string[]]($sharedMailboxes.UserPrincipalName)
        $sharedSet = New-Object System.Collections.Generic.HashSet[string]($upnArray, [System.StringComparer]::OrdinalIgnoreCase)
    }
    
    # Process users
    Write-ColorOutput "`nProcessing users and licenses...`n" "Yellow"
    $total = $allUsers.Count
    $counter = 0
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    foreach ($user in $allUsers) {
        $counter++
        
        # Skip guests and disabled users
        if ($user.UserType -ne 'Member' -or -not $user.AccountEnabled) { 
            continue 
        }
        
        $upn = $user.UserPrincipalName
        $percentComplete = [math]::Round(($counter / $total) * 100, 1)
        Write-Progress -Activity "Auditing License Compliance" -Status "Processing $upn ($counter of $total)" -PercentComplete $percentComplete
        
        $isShared = $sharedSet.Contains($upn)
        
        # Map licenses to friendly names
        $friendlyLicenses = @()
        if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
            foreach ($lic in $user.AssignedLicenses) {
                $sku = $tenantSkus | Where-Object { $_.SkuId -eq $lic.SkuId }
                if ($sku) { 
                    if ($skuMap.ContainsKey($sku.SkuPartNumber)) { 
                        $friendlyLicenses += $skuMap[$sku.SkuPartNumber]
                    }
                    else {
                        # Include unmapped licenses for visibility
                        $friendlyLicenses += "$($sku.SkuPartNumber) (unmapped)"
                    }
                }
            }
        }
        
        # Get mailbox details
        $sizeGB = 0
        $hasArchive = "No"
        $archiveSizeGB = 0
        $hasMailbox = $false
        $litigationHold = $false
        $licenseAssignment = "Direct"
        
        try {
            $mailbox = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
            if ($mailbox) {
                $hasMailbox = $true
                $litigationHold = $mailbox.LitigationHoldEnabled
                $hasArchive = if ($mailbox.ArchiveGuid -ne [Guid]::Empty -and $null -ne $mailbox.ArchiveGuid) { "Yes" } else { "No" }
                $sizeGB = Get-MailboxSizeInGB -UserPrincipalName $upn
                
                # Get archive size if archive exists
                if ($hasArchive -eq "Yes") {
                    $archiveSizeGB = Get-ArchiveUsageInGB -UserPrincipalName $upn
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve mailbox for $upn"
        }
        
        # Determine license assignment method
        if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
            # Check if any license has AssignedByGroup property
            $hasGroupAssignment = $false
            foreach ($lic in $user.AssignedLicenses) {
                # In Graph API, if there's no direct assignment, it's group-based
                # This is a simplified check - you might need to enhance based on your Graph API version
                if ($lic.PSObject.Properties.Name -contains 'AssignedByGroup') {
                    $hasGroupAssignment = $true
                    break
                }
            }
            $licenseAssignment = if ($hasGroupAssignment) { "Group-Based" } else { "Direct" }
        }
        
        # Run compliance checks
        $uniqueLicenses = $friendlyLicenses | Select-Object -Unique
        $compliance = Test-LicenseCompliance `
            -Licenses $uniqueLicenses `
            -IsSharedMailbox $isShared `
            -MailboxSizeGB $sizeGB `
            -HasArchive $hasArchive `
            -UserPrincipalName $upn `
            -ArchiveSizeGB $archiveSizeGB `
            -LitigationHoldEnabled $litigationHold `
            -LicenseAssignmentMethod $licenseAssignment
        
        # Create result object
        $results.Add([PSCustomObject]@{
            UserPrincipalName    = $upn
            DisplayName          = $user.DisplayName
            UserType             = $compliance.UserType
            HasMailbox           = $hasMailbox
            MailboxSizeGB        = $sizeGB
            HasArchive           = $hasArchive
            ArchiveSizeGB        = $archiveSizeGB
            LitigationHold       = $litigationHold
            LicenseAssignment    = $compliance.LicenseAssignment
            Licenses             = ($uniqueLicenses | Where-Object { $_ -notin $ignoredLicenses }) -join "; "
            AllLicenses          = ($uniqueLicenses -join "; ")
            Status               = $compliance.Status
            Issues               = ($compliance.Issues -join " | ")
            Warnings             = ($compliance.Warnings -join " | ")
        })
    }
    
    Write-Progress -Activity "Auditing License Compliance" -Completed
    
    # Export results
    Write-ColorOutput "`nExporting results..." "Yellow"
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    
    # Generate summary
    Write-ColorOutput "`n=== Compliance Summary ===" "Cyan"
    Write-ColorOutput "Total users processed: $($results.Count)" "White"
    
    # Status breakdown
    $compliant = ($results | Where-Object { $_.Status -eq "Compliant" }).Count
    $nonCompliant = ($results | Where-Object { $_.Status -eq "Non-Compliant" }).Count
    
    Write-ColorOutput "`nCompliance Status:" "Yellow"
    Write-ColorOutput "  Compliant:     $compliant ($([math]::Round(($compliant / $results.Count) * 100, 1))%)" "Green"
    Write-ColorOutput "  Non-Compliant: $nonCompliant ($([math]::Round(($nonCompliant / $results.Count) * 100, 1))%)" "Red"
    
    # User type breakdown
    $userTypes = $results | Group-Object UserType | Select-Object @{N='UserType';E={$_.Name}}, Count | Sort-Object Count -Descending
    Write-ColorOutput "`nUser Type Breakdown:" "Yellow"
    $userTypes | Format-Table -AutoSize | Out-String | Write-Host
    
    # Show non-compliant users
    if ($nonCompliant -gt 0) {
        Write-ColorOutput "`n=== Non-Compliant Users (First 10) ===" "Red"
        $results | Where-Object { $_.Status -eq "Non-Compliant" } | 
            Select-Object -First 10 UserPrincipalName, UserType, Issues | 
            Format-Table -Wrap | Out-String | Write-Host
        
        # Issue frequency
        Write-ColorOutput "Most Common Issues:" "Yellow"
        $allIssues = $results | Where-Object { $_.Status -eq "Non-Compliant" } | 
            ForEach-Object { $_.Issues -split '\|' | ForEach-Object { $_.Trim() } } |
            Where-Object { $_ }
        
        $issueFrequency = $allIssues | Group-Object | 
            Select-Object @{N='Issue';E={$_.Name}}, Count | 
            Sort-Object Count -Descending | 
            Select-Object -First 5
        
        $issueFrequency | Format-Table -Wrap | Out-String | Write-Host
    }
    else {
        Write-ColorOutput "`nAll users are compliant! ✓" "Green"
    }
    
    # Critical warnings (Litigation Hold)
    $litigationHoldUsers = ($results | Where-Object { $_.LitigationHold -eq $true }).Count
    if ($litigationHoldUsers -gt 0) {
        Write-ColorOutput "`n⚠ CRITICAL: $litigationHoldUsers users have Litigation Hold enabled" "Magenta"
        Write-ColorOutput "   These users MUST retain their licenses to preserve mailbox data" "Magenta"
    }
    
    # Archive usage insights
    $archivesEnabled = ($results | Where-Object { $_.HasArchive -eq "Yes" }).Count
    $emptyArchives = ($results | Where-Object { $_.HasArchive -eq "Yes" -and $_.ArchiveSizeGB -lt 1 }).Count
    if ($emptyArchives -gt 0) {
        Write-ColorOutput "`n💡 Optimization: $emptyArchives of $archivesEnabled archives are empty (<1GB)" "Cyan"
        Write-ColorOutput "   Consider disabling unused archives to simplify management" "Cyan"
    }
    
    # License assignment method
    $directAssignments = ($results | Where-Object { $_.LicenseAssignment -eq "Direct" }).Count
    $groupAssignments = ($results | Where-Object { $_.LicenseAssignment -eq "Group-Based" }).Count
    Write-ColorOutput "`nLicense Assignment Methods:" "Yellow"
    Write-ColorOutput "  Direct:      $directAssignments" "White"
    Write-ColorOutput "  Group-Based: $groupAssignments" "White"
    if ($directAssignments -gt ($groupAssignments * 0.5)) {
        Write-ColorOutput "  💡 Tip: Consider migrating to group-based licensing for easier management" "Cyan"
    }
    
    Write-ColorOutput "`nReport saved to: $OutputPath" "Green"
    Write-ColorOutput "Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "Gray"
}
catch {
    Write-ColorOutput "`nError during execution: $_" "Red"
    Write-ColorOutput $_.ScriptStackTrace "Red"
    exit 1
}
finally {
    # Optionally disconnect (commented out to allow review)
    # Disconnect-MgGraph
    # Disconnect-ExchangeOnline -Confirm:$false
}
