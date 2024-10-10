<#
    Check-SendOnBehalfPermissions.ps1

    Description:
    This script retrieves all mailboxes in an Exchange environment and checks each one for 
    any "Send on Behalf" permissions that are configured. It lists mailboxes where users 
    have been granted "Send on Behalf" rights and displays the details of the users 
    who have those permissions.

    Usage:
    - Run the script in an Exchange Management Shell or PowerShell session with the 
      necessary permissions to query mailbox properties and recipient information.
    - The results will be displayed in the PowerShell window, showing each mailbox and 
      the corresponding users who have "Send on Behalf" permissions.

    Requirements:
    - Exchange Management Shell or the required PowerShell modules for managing Exchange 
      Online or on-premise Exchange Server.
    - Sufficient privileges to run `Get-Mailbox` and `Get-Recipient` cmdlets.

    Output:
    - A list of mailboxes and users who have "Send on Behalf" permissions, displayed 
      in the PowerShell session.
#>

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Loop through each mailbox and check the "Send on Behalf" permission
foreach ($mailbox in $mailboxes) {
    $sendOnBehalf = $mailbox.GrantSendOnBehalfTo

    # If any users have "Send on Behalf" permissions, display the results
    if ($sendOnBehalf -ne $null) {
        Write-Host "Mailbox: $($mailbox.UserPrincipalName)"
        foreach ($user in $sendOnBehalf) {
            # Display the user's PrimarySmtpAddress or Name
            $userDetails = Get-Recipient -Identity $user
            Write-Host "  Has Send on Behalf permission: $($userDetails.PrimarySmtpAddress)"
        }
    }
}