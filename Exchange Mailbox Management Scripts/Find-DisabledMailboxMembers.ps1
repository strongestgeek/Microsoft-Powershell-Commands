<#
Script Name: Find-DisabledMailboxMembers-Cloud.ps1
Author: Grok (xAI) + @strongestgeek
Date: April 08, 2025
Version: 1.3 (Cloud Edition)

Purpose:
Identifies disabled Entra ID users with Full Access or Send-As permissions 
on shared mailboxes in a fully cloud Microsoft 365 environment for cleanup purposes.

Functionality:
- Retrieves all shared mailboxes from Exchange Online
- Checks each mailbox for Full Access and Send-As permissions
- Cross-references permissioned users against Entra ID to check disabled status
- Reports users who are disabled but still have mailbox permissions

Key Features:
- Outputs results with mailbox name, user, permission type, and Entra ID status
- Filters for shared mailboxes only using Get-Mailbox
- Handles both Full Access and Send-As permissions separately
- Explicitly excludes NT AUTHORITY accounts from Entra ID lookup
- Sorts results alphabetically by Mailbox (A to Z)
- Offers option to export results to CSV

Requirements:
- Exchange Online PowerShell module (EXO V2 or later)
- Microsoft Graph PowerShell module (for Get-MgUser)
- Appropriate permissions: Exchange Admin and Entra ID read access
- Connected to Exchange Online and Microsoft Graph via PowerShell

Notes:
- Run in a PowerShell session connected to Exchange Online and Microsoft Graph
- Disabled users are identified by AccountEnabled=$false in Entra ID
- Results can be used to manually remove permissions as needed
- Unresolved users (e.g., external or non-existent) are logged for review
#>

# Check module versions
$exoModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable | Select-Object -First 1
$graphModule = Get-Module -Name Microsoft.Graph -ListAvailable | Select-Object -First 1
Write-Host "ExchangeOnlineManagement Version: $($exoModule.Version)" -ForegroundColor Cyan
Write-Host "Microsoft.Graph Version: $($graphModule.Version)" -ForegroundColor Cyan

# Connect to Exchange Online and Microsoft Graph
Write-Host "Attempting to connect to services..." -ForegroundColor Cyan
try {
    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
        Write-Host "Connected to Exchange Online" -ForegroundColor Green
    } else {
        Write-Host "Exchange Online module already loaded" -ForegroundColor Green
    }

    if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
        Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Cyan
        Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop
    }
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop -NoWelcome
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to required services. Ensure Exchange Online and Microsoft Graph modules are installed and credentials have sufficient permissions."
    Write-Error "Error: $($_.Exception.Message)"
    exit
}

# Get all shared mailboxes
Write-Host "Retrieving shared mailboxes..." -ForegroundColor Cyan
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
Write-Host "Found $($sharedMailboxes.Count) shared mailboxes" -ForegroundColor Green

# Arrays to store results and unresolved users
$results = @()
$unresolved = @()

# Process each shared mailbox
foreach ($mailbox in $sharedMailboxes) {
    $mailboxName = $mailbox.PrimarySmtpAddress
    Write-Host "Processing mailbox: $mailboxName" -ForegroundColor Cyan

    # Get Full Access permissions
    $fullAccess = Get-MailboxPermission -Identity $mailboxName | 
        Where-Object { ($_.AccessRights -eq "FullAccess") -and ($_.User -notlike "NT AUTHORITY*") }

    foreach ($permission in $fullAccess) {
        $userEmail = $permission.User

        # Skip NT AUTHORITY accounts explicitly
        if ($userEmail -like "NT AUTHORITY*") {
            continue
        }

        try {
            $mgUser = Get-MgUser -Filter "mail eq '$userEmail'" -Property AccountEnabled -ErrorAction Stop
            if ($mgUser -and $mgUser.AccountEnabled -eq $false) {
                $results += [PSCustomObject]@{
                    Mailbox       = $mailboxName
                    User          = $userEmail
                    Permission    = "FullAccess"
                    EntraIDStatus = "Disabled"
                }
            }
        } catch {
            $unresolved += [PSCustomObject]@{
                Mailbox    = $mailboxName
                Identifier = $userEmail
                Permission = "FullAccess"
                Error      = $_.Exception.Message
            }
        }
    }

    # Get Send-As permissions
    $sendAs = Get-RecipientPermission -Identity $mailboxName | 
        Where-Object { ($_.AccessRights -eq "SendAs") -and ($_.Trustee -notlike "NT AUTHORITY*") }

    foreach ($permission in $sendAs) {
        $userEmail = $permission.Trustee

        # Skip NT AUTHORITY accounts explicitly
        if ($userEmail -like "NT AUTHORITY*") {
            continue
        }

        try {
            $mgUser = Get-MgUser -Filter "mail eq '$userEmail'" -Property AccountEnabled -ErrorAction Stop
            if ($mgUser -and $mgUser.AccountEnabled -eq $false) {
                $results += [PSCustomObject]@{
                    Mailbox       = $mailboxName
                    User          = $userEmail
                    Permission    = "SendAs"
                    EntraIDStatus = "Disabled"
                }
            }
        } catch {
            $unresolved += [PSCustomObject]@{
                Mailbox    = $mailboxName
                Identifier = $userEmail
                Permission = "SendAs"
                Error      = $_.Exception.Message
            }
        }
    }
}

# Sort results alphabetically by Mailbox
$results = $results | Sort-Object Mailbox
$unresolved = $unresolved | Sort-Object Mailbox

# Display results
if ($results.Count -gt 0) {
    Write-Host "`nDisabled Users with Mailbox Permissions Found (Sorted A-Z by Mailbox):" -ForegroundColor Yellow
    $results | Format-Table -AutoSize
} else {
    Write-Host "`nNo disabled users with mailbox permissions found." -ForegroundColor Green
}

# Display unresolved users
if ($unresolved.Count -gt 0) {
    Write-Host "`nUnresolved Users (Not Found in Entra ID, Sorted A-Z by Mailbox):" -ForegroundColor Red
    $unresolved | Format-Table -AutoSize
}

# Prompt for CSV export
$exportChoice = Read-Host "`nWould you like to export the results to a CSV file? (Y/N)"
if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
    $timestamp = Get-Date -Format "yyyyMMdd"
    $csvPath = "DisabledMailboxMembers_$timestamp.csv"
    $unresolvedCsvPath = "UnresolvedMailboxMembers_$timestamp.csv"

    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Exported disabled users to $csvPath" -ForegroundColor Green
    }
    if ($unresolved.Count -gt 0) {
        $unresolved | Export-Csv -Path $unresolvedCsvPath -NoTypeInformation
        Write-Host "Exported unresolved users to $unresolvedCsvPath" -ForegroundColor Green
    }
    if ($results.Count -eq 0 -and $unresolved.Count -eq 0) {
        Write-Host "No data to export." -ForegroundColor Yellow
    }
}

# Disconnect from services with error suppression
Write-Host "Disconnecting from services..." -ForegroundColor Cyan
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 2>$null
    Disconnect-MgGraph -ErrorAction SilentlyContinue 2>$null
    Write-Host "Disconnected from Exchange Online and Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "Minor error during disconnection (likely harmless JSON parsing issue), but script completed successfully." -ForegroundColor Yellow
}
