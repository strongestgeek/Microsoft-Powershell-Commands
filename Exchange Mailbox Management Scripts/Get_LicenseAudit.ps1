<#
.SYNOPSIS
    Audits Microsoft 365 licenses for both user and shared mailboxes.

.DESCRIPTION
    This script connects to Microsoft Graph and Exchange Online to generate a comprehensive license audit report.
    It identifies redundant licenses, missing required add-ons, and inappropriate license assignments for shared mailboxes
    based on their size and archive status.

    The script requires the 'ExchangeOnlineManagement' and 'Microsoft.Graph' modules.
    Permissions needed:
    - Exchange Online: Exchange Administrator role.
    - Microsoft Graph: User.Read.All, Directory.Read.All.

.OUTPUTS
    - A CSV file containing the detailed license audit.
    - A transcript log file capturing the script's execution details and any errors.

.NOTES
    Version: 1.0
    Author: Gemini
    Date: 2025-09-05
#>

#requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement

#=======================================================================================================================
# SCRIPT CONFIGURATION
#=======================================================================================================================

#region Script Parameters
Param(
    [string]$CsvOutputPath = ".\M365_License_Audit_Report_$(Get-Date -Format 'yyyy-MM-dd').csv",
    [string]$LogFilePath = ".\M365_License_Audit_Log_$(Get-Date -Format 'yyyy-MM-dd').log"
)
#endregion

#=======================================================================================================================
# SETUP AND FUNCTIONS
#=======================================================================================================================

#region Logging Setup
# Start a transcript to log all console output to a file.
try {
    Start-Transcript -Path $LogFilePath -Append -ErrorAction Stop
}
catch {
    Write-Host "FATAL: Could not start transcript at '$LogFilePath'. Please check permissions." -ForegroundColor Red
    exit 1
}

# Custom logging function to write to both console and the transcript log.
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO" # INFO, WARN, ERROR
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd"
    $formattedMessage = "[$timestamp] [$Level] $Message"
    
    $colorMap = @{
        "INFO"  = "White"
        "WARN"  = "Yellow"
        "ERROR" = "Red"
    }
    
    Write-Host $formattedMessage -ForegroundColor $colorMap[$Level]
}
#endregion

#region License Mapping and Audit Logic
# A dictionary to map SKU Part Numbers (from Microsoft) to user-friendly product names.
# This list covers many common SKUs but can be extended.
$global:skuPartNumberMapping = @{
    "AAD_PREMIUM"                     = "Azure Active Directory Premium P1"
    "AAD_PREMIUM_P2"                  = "Azure Active Directory Premium P2"
    "ENTERPRISEPREMIUM"               = "Microsoft 365 E5"
    "ENTERPRISEPACK"                  = "Office 365 E3"
    "STANDARDPACK"                    = "Office 365 E1"
    "DESKLESSPACK"                    = "Microsoft 365 F3" # Or Office 365 F3
    "M365_BUSINESS_PREMIUM"           = "Microsoft 365 Business Premium"
    "M365_BUSINESS_STANDARD"          = "Microsoft 365 Business Standard"
    "MCOTEAMS_ESSENTIALS"             = "Microsoft Teams Essentials (AAD Identity)"
    "TEAMS_EXPLORATORY"               = "Microsoft Teams Exploratory"
    "TEAMS_ENTERPRISE"                = "Microsoft Teams Enterprise"
    "EXCHANGE_S_ENTERPRISE"           = "Exchange Online (Plan 2)"
    "EXCHANGE_S_PLAN1"                = "Exchange Online (Plan 1)"
    "EXCHANGESTANDARD"                = "Exchange Online Kiosk"
    "INTUNE_A"                        = "Microsoft Intune Plan 1"
    "M365_E5_COMPLIANCE"              = "Microsoft 365 E5 Compliance"
    "M365_E5_SECURITY"                = "Microsoft 365 E5 Security"
    "RIGHTSMANAGEMENT_STANDARD"       = "Azure Information Protection Premium P1"
    "POWER_BI_PRO"                    = "Power BI Pro"
    "PROJECTPROFESSIONAL"             = "Project Plan 3"
    "VISIO_PLAN2"                     = "Visio Plan 2"
    "WIN_DEF_VULN_MGMT_ADDON"         = "Microsoft Defender Vulnerability Management Add-on"
    "SPE_E3_NOPSTN"                   = "Microsoft 365 E3 (no Teams)"
}

# This function determines the audit status and provides notes based on the mailbox and license data.
function Get-LicenseAuditStatus {
    param(
        [string]$Type,
        [array]$AssignedLicenses,
        [double]$MailboxSizeGB,
        [bool]$ArchiveEnabled
    )
    
    $status = "OK"
    $notes = "No action needed"

    if ($Type -eq 'User') {
        if ($AssignedLicenses -contains "Microsoft 365 E5" -and $AssignedLicenses -contains "Office 365 E3") {
            $status = "Redundant license detected"
            $notes = "Remove Office 365 E3"
        }
        elseif ($AssignedLicenses -contains "Microsoft 365 E3 (no Teams)") {
            $status = "Missing Microsoft Teams Enterprise"
            $notes = "Assign Microsoft Teams Enterprise for full functionality."
        }
        elseif ($AssignedLicenses -contains "Microsoft 365 E3" -and -not ($AssignedLicenses -contains "Microsoft 365 E5 Security" -or $AssignedLicenses -contains "Microsoft Defender Vulnerability Management Add-on")) {
            # This is an example of a custom business rule. You can modify this.
            # $status = "Missing required add-on for Microsoft 365 E3"
            # $notes = "Consider assigning Microsoft 365 E5 Security and/or Microsoft Defender Vulnerability Management Add-on."
        }
    }
    elseif ($Type -eq 'SharedMailbox') {
        $hasLicense = $AssignedLicenses.Count -gt 0
        $hasEOP1 = $AssignedLicenses -contains "Exchange Online (Plan 1)"
        $hasEOP2 = $AssignedLicenses -contains "Exchange Online (Plan 2)"
        $hasFullLicense = $AssignedLicenses -match "E3|E5|Business Premium"
        
        $issues = @()
        
        # Check for size > 50GB
        if ($MailboxSizeGB -gt 50 -and !$hasEOP2) {
            $issues += "Missing license for size >50 GB"
        }
        
        # Check for archive enabled
        if ($ArchiveEnabled -and !$hasEOP2) {
            # Note: A Plan 1 license allows a 50GB archive, but for simplicity and future growth, Plan 2 is the standard recommendation for archiving.
            $issues += "Missing license for archiving"
        }
        
        # Check for unnecessary full licenses
        if ($hasFullLicense) {
            $issues += "Unnecessary license assigned"
        }

        if ($issues.Count -gt 0) {
            $status = ($issues -join '; ')
            $notes = if ($issues -match "Missing license") { "Assign Exchange Online (Plan 2)" }
            elseif ($issues -match "Unnecessary license") { "Remove $($AssignedLicenses -match "E3|E5|Business Premium" | Out-String -Stream). Trim() for cost savings" }
            else { "Review license assignment."}
        }
    }
    
    return @{ Status = $status; Notes = $notes }
}

#endregion

#=======================================================================================================================
# MAIN SCRIPT BODY
#=======================================================================================================================

Write-Log "Starting M365 License Audit Script."
$auditResults = [System.Collections.Generic.List[PSObject]]::new()

try {
    #region Connect to Services
    Write-Log "Connecting to services..."
    # Connect to Exchange Online
    try {
        Write-Log "Attempting to connect to Exchange Online..."
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Log "Successfully connected to Exchange Online."
    }
    catch {
        throw "Failed to connect to Exchange Online. Please ensure you have the ExchangeOnlineManagement module installed and appropriate permissions."
    }

    # Connect to Microsoft Graph
    try {
        Write-Log "Attempting to connect to Microsoft Graph..."
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -ErrorAction Stop
        Write-Log "Successfully connected to Microsoft Graph."
    }
    catch {
        throw "Failed to connect to Microsoft Graph. Please ensure you have the Microsoft.Graph module installed and appropriate permissions."
    }
    #endregion

    #region Data Gathering
    Write-Log "Gathering data from Microsoft 365 tenant. This may take a while..."

    # 1. Fetch Tenant's specific SKUs to supplement the built-in map
    Write-Log "Fetching tenant subscribed SKUs..."
    $tenantSkus = Get-MgSubscribedSku -All
    foreach ($sku in $tenantSkus) {
        if (-not $global:skuPartNumberMapping.ContainsKey($sku.SkuPartNumber)) {
            $global:skuPartNumberMapping[$sku.SkuPartNumber] = $sku.SkuPartNumber # Use the PartNumber itself as a fallback friendly name
            Write-Log "Discovered new SKU '$($sku.SkuPartNumber)' and added it to the mapping." -Level WARN
        }
    }
    
    # 2. Get all Shared Mailboxes and store their UPNs for filtering later
    Write-Log "Fetching all Shared Mailboxes from Exchange Online..."
    $sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Select-Object UserPrincipalName, DisplayName, ArchiveGuid
    $sharedMailboxUpns = $sharedMailboxes.UserPrincipalName | ForEach-Object { $_.ToLower() } | Select-Object -Unique
    Write-Log "Found $($sharedMailboxes.Count) shared mailboxes."

    # 3. Get all licensed users from Microsoft Graph, excluding guests and shared mailboxes
    Write-Log "Fetching all licensed users from Microsoft Graph..."
    $graphUsers = Get-MgUser -All -Filter "userType eq 'Member'" -Property Id, UserPrincipalName, DisplayName, AssignedLicenses | Where-Object { $_.AssignedLicenses -and ($sharedMailboxUpns -notcontains $_.UserPrincipleName.ToLower()) }
    Write-Log "Found $($graphUsers.Count) licensed users (excluding shared mailboxes)."

    # Combine Users and Shared Mailboxes into one list for processing
    $allMailboxesToProcess = @()
    $allMailboxesToProcess += $graphUsers | Select-Object @{N = 'Type'; E = { 'User' } }, UserPrincipalName, DisplayName, @{N = 'Licenses'; E = { $_.AssignedLicenses } }
    $allMailboxesToProcess += $sharedMailboxes | Select-Object @{N = 'Type'; E = { 'SharedMailbox' } }, UserPrincipalName, DisplayName, @{N = 'Licenses'; E = { 
            # For shared mailboxes, we must separately query Graph for any assigned licenses
            (Get-MgUser -UserId $_.UserPrincipalName -Property AssignedLicenses -ErrorAction SilentlyContinue).AssignedLicenses
        } }

    #endregion

    #region Main Processing Loop
    Write-Log "Processing $($allMailboxesToProcess.Count) total mailboxes. Fetching individual stats..."
    $total = $allMailboxesToProcess.Count
    $counter = 0

    foreach ($mailbox in $allMailboxesToProcess) {
        $counter++
        Write-Progress -Activity "Auditing Mailboxes" -Status "Processing $($mailbox.UserPrincipalName) ($counter of $total)" -PercentComplete (($counter / $total) * 100)
        
        $upn = $mailbox.UserPrincipalName
        $assignedLicensesFriendly = @()
        
        # Map assigned license SKUs to friendly names
        if ($mailbox.Licenses) {
            foreach ($license in $mailbox.Licenses) {
                $skuId = $license.SkuId
                $subscribedSku = $tenantSkus | Where-Object { $_.SkuId -eq $skuId }
                if ($subscribedSku) {
                    $friendlyName = $global:skuPartNumberMapping[$subscribedSku.SkuPartNumber]
                    if ($friendlyName) {
                        $assignedLicensesFriendly += $friendlyName
                    }
                    else {
                        $assignedLicensesFriendly += $subscribedSku.SkuPartNumber # Fallback
                    }
                }
            }
        }
        
        # Get Mailbox Size and Archive Status
        $mailboxSizeGB = $null
        $archiveEnabled = $false
        try {
            # Get-Mailbox is needed for ArchiveGuid which reliably shows if archive is provisioned
            $exoMailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
            $archiveEnabled = $null -ne $exoMailbox.ArchiveGuid

            $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop
            # Robustly parse size, avoiding ToBytes errors on strings
            if ($stats.TotalItemSize.Value) {
                $mailboxSizeGB = [math]::Round((($stats.TotalItemSize).Value.ToBytes() / 1GB), 2)
            } else {
                 $mailboxSizeGB = 0
            }
        }
        catch {
            Write-Log "Could not get statistics for '$upn'. It may not have a provisioned mailbox. Error: $($_.Exception.Message)" -Level WARN
            continue # Skip to the next mailbox
        }

        # Perform the audit
        $audit = Get-LicenseAuditStatus -Type $mailbox.Type -AssignedLicenses $assignedLicensesFriendly -MailboxSizeGB $mailboxSizeGB -ArchiveEnabled $archiveEnabled
        
        # Create the output object
        $outputObject = [PSCustomObject]@{
            Type             = $mailbox.Type
            Identity         = $upn
            DisplayName      = $mailbox.DisplayName
            AssignedLicenses = $assignedLicensesFriendly -join ';'
            MailboxSizeGB    = if ($null -ne $mailboxSizeGB) { "$mailboxSizeGB" } else { "" }
            ArchiveEnabled   = if ($archiveEnabled) { "Yes" } else { "No" }
            Status           = $audit.Status
            Notes            = $audit.Notes
        }
        
        $auditResults.Add($outputObject)
    }
    #endregion
    
    #region Export Results
    if ($auditResults.Count -gt 0) {
        Write-Log "Processing complete. Exporting $($auditResults.Count) results to '$CsvOutputPath'..."
        $auditResults | Export-Csv -Path $CsvOutputPath -NoTypeInformation -Encoding UTF8
        Write-Log "Successfully exported the report."
    }
    else {
        Write-Log "No mailboxes were processed or found. The report is empty." -Level WARN
    }
    #endregion
}
catch {
    Write-Log "A critical error occurred: $($_.Exception.Message)" -Level ERROR
}
finally {
    #region Disconnect and Cleanup
    Write-Log "Disconnecting from all services."
    if (Get-ConnectionInformation -Module "ExchangeOnlineManagement" -SessionType "Exchange") {
        Disconnect-ExchangeOnline -Confirm:$false
    }
    if (Get-MgProfile) {
        Disconnect-MgGraph
    }
    
    Write-Log "Script finished. Log file saved at: $LogFilePath"
    Stop-Transcript
    #endregion
}

