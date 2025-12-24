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
try {
    Start-Transcript -Path $LogFilePath -Append -ErrorAction Stop
}
catch {
    Write-Host "FATAL: Could not start transcript at '$LogFilePath'. Check permissions." -ForegroundColor Red
    exit 1
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedMessage = "[$timestamp] [$Level] $Message"
    $colorMap = @{"INFO"="White"; "WARN"="Yellow"; "ERROR"="Red"}
    Write-Host $formattedMessage -ForegroundColor $colorMap[$Level]
}
#endregion

#region License Mapping and Audit Logic
$global:skuPartNumberMapping = @{
    "AAD_PREMIUM"             = "Microsoft Entra ID P1"
    "AAD_PREMIUM_P2"          = "Microsoft Entra ID P2"
    "ENTERPRISEPREMIUM"       = "Microsoft 365 E5"
    "ENTERPRISEPACK"          = "Office 365 E3"
    "STANDARDPACK"            = "Office 365 E1"
    "DESKLESSPACK"            = "Microsoft 365 F3"
    "M365_BUSINESS_PREMIUM"   = "Microsoft 365 Business Premium"
    "M365_BUSINESS_STANDARD"  = "Microsoft 365 Business Standard"
    "SPE_E3_NOPSTN"           = "Microsoft 365 E3 (no Teams)"
    # Add other SKUs as needed
}

function Get-LicenseAuditStatus {
    param([string]$Type, [array]$AssignedLicenses, [double]$MailboxSizeGB, [bool]$ArchiveEnabled)

    $status = "OK"
    $notes = "No action needed"

    if ($Type -eq 'User') {
        if ($AssignedLicenses -contains "Microsoft 365 E5" -and $AssignedLicenses -contains "Office 365 E3") {
            $status = "Redundant license detected"; $notes = "Remove Office 365 E3"
        }
    }
    elseif ($Type -eq 'SharedMailbox') {
        $issues = @()
        if ($MailboxSizeGB -gt 50 -and $AssignedLicenses -notmatch "Plan 2|E3|E5") { $issues += "Size >50GB needs Plan 2" }
        if ($ArchiveEnabled -and $AssignedLicenses -notmatch "Plan 2|E3|E5") { $issues += "Archive needs Plan 2" }
        if ($AssignedLicenses -match "E3|E5|Business Premium") { $issues += "Unnecessary full license" }

        if ($issues.Count -gt 0) {
            $status = $issues -join '; '
            $notes = "Review licensing for cost/compliance."
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
    Connect-ExchangeOnline -ErrorAction Stop
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -ErrorAction Stop
    #endregion

    #region Data Gathering (Optimized)
    Write-Log "Gathering tenant data..."
    $tenantSkus = Get-MgSubscribedSku -All
    
    Write-Log "Pre-fetching license data from Graph..."
    $allGraphUsers = Get-MgUser -All -Property Id, UserPrincipalName, DisplayName, AssignedLicenses, UserType
    $licenseLookup = $allGraphUsers | Group-Object -Property UserPrincipalName -AsHashTable -IgnoreCase

    Write-Log "Fetching Shared Mailboxes..."
    $sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    $sharedMailboxUpns = New-Object System.Collections.Generic.HashSet[string] ([string[]]($sharedMailboxes.UserPrincipalName), [System.StringComparer]::OrdinalIgnoreCase)

    # Combine lists
    $allToProcess = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($u in $allGraphUsers) {
        if ($u.AssignedLicenses -and $u.UserType -eq 'Member' -and -not $sharedMailboxUpns.Contains($u.UserPrincipalName)) {
            $allToProcess.Add([PSCustomObject]@{Type='User'; UPN=$u.UserPrincipalName; Name=$u.DisplayName; Licenses=$u.AssignedLicenses})
        }
    }
    foreach ($sm in $sharedMailboxes) {
        $allToProcess.Add([PSCustomObject]@{Type='SharedMailbox'; UPN=$sm.UserPrincipalName; Name=$sm.DisplayName; Licenses=$licenseLookup[$sm.UserPrincipalName].AssignedLicenses})
    }
    #endregion

    #region Main Processing Loop
    $total = $allToProcess.Count
    $counter = 0

    foreach ($entry in $allToProcess) {
        $counter++
        Write-Progress -Activity "Auditing" -Status "$($entry.UPN)" -PercentComplete (($counter/$total)*100)

        $assignedFriendly = foreach ($lic in $entry.Licenses) {
            $sku = $tenantSkus | Where-Object { $_.SkuId -eq $lic.SkuId }
            if ($sku) { $global:skuPartNumberMapping[$sku.SkuPartNumber] ?? $sku.SkuPartNumber }
        }

        # SAFE MAILBOX STATS CHECK
        $sizeGB = 0; $archive = $false
        try {
            $m = Get-Mailbox -Identity $entry.UPN -ErrorAction Stop
            $archive = $null -ne $m.ArchiveGuid
            $stats = Get-MailboxStatistics -Identity $entry.UPN -ErrorAction SilentlyContinue
            
            # The "Null-Valued Expression" Fix:
            if ($null -ne $stats -and $null -ne $stats.TotalItemSize -and $null -ne $stats.TotalItemSize.Value) {
                $sizeGB = [math]::Round(($stats.TotalItemSize.Value.ToBytes() / 1GB), 2)
            }
        } catch { Write-Log "No mailbox found for $($entry.UPN)" -Level WARN }

        $audit = Get-LicenseAuditStatus -Type $entry.Type -AssignedLicenses $assignedFriendly -MailboxSizeGB $sizeGB -ArchiveEnabled $archive

        $auditResults.Add([PSCustomObject]@{
            Type             = $entry.Type
            Identity         = $entry.UPN
            DisplayName      = $entry.Name
            AssignedLicenses = $assignedFriendly -join ';'
            MailboxSizeGB    = $sizeGB
            ArchiveEnabled   = if($archive){"Yes"}else{"No"}
            Status           = $audit.Status
            Notes            = $audit.Notes
        })
    }

    $auditResults | Export-Csv -Path $CsvOutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Report exported to $CsvOutputPath"
    #endregion
}
catch {
    Write-Log "Critical Error: $($_.Exception.Message)" -Level ERROR
}
finally {
    Write-Log "Closing connections..."
    if (Get-Module -Name "ExchangeOnlineManagement") { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
    if (Get-MgContext) { Disconnect-MgGraph }
    Stop-Transcript
}


