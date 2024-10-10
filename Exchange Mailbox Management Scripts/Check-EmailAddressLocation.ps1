<#
    Check-EmailAddressLocation.ps1

    Description:
    This script checks the location of a specified email address within an Exchange 
    environment. It determines whether the email address belongs to a mailbox, 
    shared mailbox, or distribution group and displays the relevant information. 
    Additionally, it checks if the email address is an alias for another recipient 
    and outputs the associated details.

    Usage:
    - Set the `$emailAddress` variable to the email address you want to search (e.g., 
      "user@email.com").
    - Run the script in an Exchange Management Shell or a PowerShell session with the 
      necessary permissions to query recipient details.
    - The script will display information about the recipient if found, and will also 
      indicate if the email address is an alias of another recipient.

    Requirements:
    - Exchange Management Shell or the appropriate PowerShell modules for Exchange Online 
      or on-premise Exchange Server.
    - Sufficient permissions to run `Get-Recipient` and `Get-Mailbox`.

    Output:
    - Information about the location of the email address (mailbox, shared mailbox, 
      or distribution group) and whether it is an alias for another recipient, displayed 
      in the PowerShell console.
#>

# Define the email address you want to search
$emailAddress = "user@email.com"

# Find the location (mailbox or distribution group) of the email address
$recipient = Get-Recipient -Identity $emailAddress -ErrorAction SilentlyContinue

if ($recipient -ne $null) {
    Write-Host "Email address '$emailAddress' is located in:"
    $recipient | Format-List Name, RecipientType, RecipientTypeDetails, PrimarySmtpAddress
} else {
    Write-Host "Email address '$emailAddress' not found."
}

# Check if the email address is an alias of another recipient
$aliasRecipient = Get-Mailbox -RecipientTypeDetails Mailbox,SharedMailbox -Filter {EmailAddresses -like "*$emailAddress*"} -ErrorAction SilentlyContinue

if ($aliasRecipient -ne $null) {
    Write-Host "Email address '$emailAddress' is an alias of:"
    $aliasRecipient | Format-List Name, PrimarySmtpAddress, EmailAddresses
} else {
    Write-Host "Email address '$emailAddress' is not an alias."
}