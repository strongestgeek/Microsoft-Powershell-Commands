<#
Script Name: Get-MailboxMembers.ps1
Author: Grok (xAI)
Date: April 08, 2025
Version: 1.0

Purpose:
Lists users with Full Access permissions for a specified shared mailbox.

Functionality:
- Prompts for a mailbox email address input
- Retrieves Full Access permissions using Get-MailboxPermission
- Filters out system accounts and displays user permissions

Key Features:
- Displays results in console with color-coded sections
- Excludes NT AUTHORITY system accounts from output

Requirements:
- Exchange Management Shell or Exchange Online PowerShell module
- Appropriate administrative permissions for Exchange environment
- Active connection to Exchange server or Exchange Online

Notes:
- Input must be a valid mailbox email address
- Run in an Exchange-connected PowerShell session
- May need adjustment for specific Exchange versions or configurations
#>

$mailbox = Read-Host "Enter the mailbox email address"

# Get Full Access permissions
Write-Host "`nUsers with Full Access:" -ForegroundColor Green
Get-MailboxPermission -Identity $mailbox | 
    Where-Object { ($_.AccessRights -eq "FullAccess") -and ($_.User -notlike "NT AUTHORITY*") } | 
    Select-Object User, AccessRights
