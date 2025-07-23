<#
.SYNOPSIS
    Get-MailboxAccessAndForwarding.ps1 - Analyzes Exchange Online mailbox permissions and forwarding settings.

.DESCRIPTION
    This PowerShell script retrieves access permissions and email forwarding details for a specified mailbox. It:
    - Lists users with FullAccess or ReadPermission, excluding system accounts.
    - Checks mailbox-level forwarding settings (internal/external addresses and delivery options).
    - Identifies inbox rules (including hidden) that forward or redirect emails, with rule status.
    Results are displayed in the console and exported to a CSV file. Ideal for mailbox audits, security reviews, or compliance checks.

.PREREQUISITES
    - Requires ExchangeOnlineManagement module: Install-Module -Name ExchangeOnlineManagement -Force
	- Connect to Exhcnge Online with: connect-ExchangeOnline -UserPrincipalName <admin@email.com>
    - Account must have Exchange Admin or Global Admin permissions.

.USAGE
    Run the script, enter the target mailbox email address when prompted, and review the console output and CSV file (MailboxAccessAndForwarding_<email>.csv).

.NOTES
    Author: Grok, created by xAI
    Date: July 23, 2025
#>

# Prompt for the email address of the mailbox
$emailAddress = Read-Host "Enter the email address of the mailbox"

# Output file path
$exportPath = ".\MailboxAccessAndForwarding_$($emailAddress -replace '[^\w\.]','_').csv"

# Initialize an array to store results
$results = @()

try {
    # Get mailbox information
    $mailbox = Get-Mailbox -Identity $emailAddress -ErrorAction Stop

    # Get mailbox permissions (FullAccess and Read permissions)
    $permissions = Get-MailboxPermission -Identity $mailbox.UserPrincipalName | 
        Where-Object { 
            $_.AccessRights -match "FullAccess|ReadPermission" -and 
            $_.User -notlike "NT AUTHORITY\SELF" -and 
            $_.User -notlike "S-1-5-*" 
        }

    # Process mailbox permissions
    foreach ($perm in $permissions) {
        $results += [PSCustomObject]@{
            Mailbox       = $mailbox.PrimarySmtpAddress
            Type          = "Permission"
            AccessType    = $perm.AccessRights
            User          = $perm.User
            ForwardTo     = $null
            RuleName      = $null
            RuleEnabled   = $null
        }
    }

    # Check mailbox-level forwarding settings
    if ($mailbox.ForwardingSmtpAddress -or $mailbox.ForwardingAddress) {
        $results += [PSCustomObject]@{
            Mailbox       = $mailbox.PrimarySmtpAddress
            Type          = "Mailbox Forwarding"
            AccessType    = $null
            User          = $null
            ForwardTo     = $mailbox.ForwardingSmtpAddress ? $mailbox.ForwardingSmtpAddress : $mailbox.ForwardingAddress
            RuleName      = $null
            RuleEnabled   = $mailbox.DeliverToMailboxAndForward ? "True" : "False"
        }
    }

    # Get inbox rules for forwarding
    $inboxRules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName -IncludeHidden |
        Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo }

    # Process inbox rules
    foreach ($rule in $inboxRules) {
        $forwardTo = @($rule.ForwardTo, $rule.ForwardAsAttachmentTo, $rule.RedirectTo) -join ", "
        $results += [PSCustomObject]@{
            Mailbox       = $mailbox.PrimarySmtpAddress
            Type          = "Inbox Rule"
            AccessType    = $null
            User          = $null
            ForwardTo     = $forwardTo
            RuleName      = $rule.Name
            RuleEnabled   = $rule.Enabled ? "True" : "False"
        }
    }

    # If no results, add a placeholder
    if ($results.Count -eq 0) {
        $results += [PSCustomObject]@{
            Mailbox       = $mailbox.PrimarySmtpAddress
            Type          = "None"
            AccessType    = "None"
            User          = "No permissions or forwarding rules found"
            ForwardTo     = $null
            RuleName      = $null
            RuleEnabled   = $null
        }
    }

    # Export results to CSV
    $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to $exportPath" -ForegroundColor Green

    # Display results in console
    $results | Format-Table -AutoSize
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}