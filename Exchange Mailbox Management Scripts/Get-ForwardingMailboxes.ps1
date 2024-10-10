<# 
    Get-ForwardingMailboxes.ps1

    Description:
    This script retrieves all mailboxes from an Exchange environment and checks if they 
    have forwarding configured. It identifies both SMTP forwarding addresses and recipients 
    with mailbox-level forwarding rules. The results include the UserPrincipalName, 
    forwarding SMTP address, and resolved forwarding email address (if applicable).

    Usage:
    - Run the script in an Exchange Management Shell or PowerShell session with the 
      appropriate permissions to access mailbox properties.
    - The results will be displayed in a formatted table within the PowerShell window.
    - Optionally, you can export the results to a CSV file by modifying or keeping the 
      path in the `Export-Csv` command (e.g., ".\ForwardingMailboxes.csv").

    Requirements:
    - Exchange Management Shell or the necessary PowerShell modules for Exchange Online 
      or on-premise Exchange Server.

    Output:
    - A list of mailboxes with forwarding addresses or recipients, displayed in the 
      PowerShell session, and optionally exported to a CSV file.
#>

# Retrieve all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Initialize an empty array to hold the results
$forwardingMailboxes = @()

# Loop through each mailbox to check for forwarding
foreach ($mailbox in $mailboxes) {
    $forwardingAddress = $mailbox.ForwardingSMTPAddress
    $forwardingRecipient = $mailbox.ForwardingAddress

    if ($forwardingAddress -or $forwardingRecipient) {
        # If ForwardingAddress is set, resolve it to a full email address
        $forwardingEmail = if ($forwardingRecipient) {
            (Get-Recipient $forwardingRecipient).PrimarySmtpAddress
        } else {
            $null
        }

        $forwardingMailboxes += [PSCustomObject]@{
            Mailbox           = $mailbox.UserPrincipalName
            ForwardingAddress = $forwardingAddress
            ForwardingEmail   = $forwardingEmail
        }
    }
}

# Output the results
$forwardingMailboxes | Format-Table -AutoSize

# Optionally, export the results to a CSV file
$forwardingMailboxes | Export-Csv -Path ".\ForwardingMailboxes.csv" -NoTypeInformation