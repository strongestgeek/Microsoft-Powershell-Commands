<#
Script Name: Remove-DisabledMailboxMembers.ps1
Author: Grok (xAI) + @strongestgeek
Date: April 08, 2025
Version: 1.3 (Cloud Edition)

Purpose:
Removes disabled Entra ID users from shared mailbox permissions based on a CSV file 
generated by Find-DisabledMailboxMembers.ps1 in a fully cloud Microsoft 365 environment.

Functionality:
- Imports a CSV file containing disabled users with mailbox permissions
- Iterates through each entry to remove Full Access or Send-As permissions
- Logs actions and any errors encountered during the process

Key Features:
- Supports removal of Full Access and Send-As permissions
- Validates CSV input for required columns (Mailbox, User, Permission)
- Provides detailed feedback on successful removals and failures
- Works in a fully cloud Exchange Online environment

Requirements:
- Exchange Online PowerShell module (EXO V2 or later)
- Appropriate permissions: Exchange Admin role to modify mailbox permissions
- Connected to Exchange Online via PowerShell
- A valid CSV file from Find-DisabledMailboxMembers-Cloud.ps1

Notes:
- Run in a PowerShell session connected to Exchange Online
- Ensure the CSV file path is correct and contains the expected data
- Only processes entries where EntraIDStatus is "Disabled"
- Errors (e.g., permission already removed) are logged but do not halt execution
#>

# Prompt for CSV file path
$csvPath = Read-Host "Enter the path to the DisabledMailboxMembers CSV file (e.g., DisabledMailboxMembers_20250408.csv)"

# Validate CSV file exists
if (-not (Test-Path $csvPath)) {
    Write-Error "CSV file not found at $csvPath. Please provide a valid path."
    exit
}

# Connect to Exchange Online
Write-Host "Attempting to connect to Exchange Online..." -ForegroundColor Cyan
try {
    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
        Write-Host "Connected to Exchange Online" -ForegroundColor Green
    } else {
        Write-Host "Exchange Online module already loaded" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to connect to Exchange Online. Ensure the module is installed and credentials have sufficient permissions."
    Write-Error "Error: $($_.Exception.Message)"
    exit
}

# Import the CSV file
Write-Host "Importing data from $csvPath..." -ForegroundColor Cyan
$disabledMembers = Import-Csv -Path $csvPath
Write-Host "Found $($disabledMembers.Count) entries in the CSV" -ForegroundColor Green

# Validate CSV headers
$requiredHeaders = @("Mailbox", "User", "Permission", "EntraIDStatus")
$csvHeaders = ($disabledMembers | Get-Member -MemberType NoteProperty).Name
foreach ($header in $requiredHeaders) {
    if ($header -notin $csvHeaders) {
        Write-Error "CSV file is missing required column: $header. Please use a valid CSV from Find-DisabledMailboxMembers-Cloud.ps1."
        exit
    }
}

# Array to store removal results
$removalResults = @()

# Process each disabled member
foreach ($member in $disabledMembers) {
    # Only process if the user is disabled
    if ($member.EntraIDStatus -ne "Disabled") {
        continue
    }

    $mailbox = $member.Mailbox
    $user = $member.User
    $permission = $member.Permission

    Write-Host "Processing: $user on $mailbox ($permission)" -ForegroundColor Cyan

    try {
        if ($permission -eq "FullAccess") {
            Remove-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
            $removalResults += [PSCustomObject]@{
                Mailbox    = $mailbox
                User       = $user
                Permission = "FullAccess"
                Status     = "Removed"
                Error      = $null
            }
            Write-Host "Successfully removed Full Access for $user from $mailbox" -ForegroundColor Green
        } elseif ($permission -eq "SendAs") {
            Remove-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            $removalResults += [PSCustomObject]@{
                Mailbox    = $mailbox
                User       = $user
                Permission = "SendAs"
                Status     = "Removed"
                Error      = $null
            }
            Write-Host "Successfully removed Send-As for $user from $mailbox" -ForegroundColor Green
        } else {
            $removalResults += [PSCustomObject]@{
                Mailbox    = $mailbox
                User       = $user
                Permission = $permission
                Status     = "Skipped"
                Error      = "Invalid permission type"
            }
            Write-Host "Skipped $user on $mailbox - invalid permission type: $permission" -ForegroundColor Yellow
        }
    } catch {
        $removalResults += [PSCustomObject]@{
            Mailbox    = $mailbox
            User       = $user
            Permission = $permission
            Status     = "Failed"
            Error      = $_.Exception.Message
        }
        Write-Host "Failed to remove $permission for $user from $mailbox - Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Display removal summary
Write-Host "`nRemoval Summary:" -ForegroundColor Cyan
$removalResults | Format-Table -AutoSize

# Export removal results to CSV
$exportChoice = Read-Host "`nWould you like to export the removal results to a CSV file? (Y/N)"
if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $resultsCsvPath = "RemovalResults_$timestamp.csv"
    $removalResults | Export-Csv -Path $resultsCsvPath -NoTypeInformation
    Write-Host "Exported removal results to $resultsCsvPath" -ForegroundColor Green
}

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 2>$null
    Write-Host "Disconnected from Exchange Online" -ForegroundColor Green
} catch {
    Write-Host "Minor error during disconnection (likely harmless JSON parsing issue), but script completed successfully." -ForegroundColor Yellow
}