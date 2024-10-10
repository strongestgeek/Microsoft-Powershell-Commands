<#
Grant-CalendarPermissions.ps1

Description:
This script grants "Editor" access to specific user accounts on the calendar folders 
of a list of room mailboxes. It loops through a predefined list of room mailboxes 
and assigns the specified permissions to User1 and User2 for each mailbox's calendar.

Usage:
- Modify the `$mailboxes` array to include the email addresses of the room mailboxes 
  you want to grant access to.
- Replace "user1@email.com" and "user2@email.com" with the user accounts that should 
  receive "Editor" rights.
- Run the script in an Exchange Management Shell or a PowerShell session with the 
  necessary Exchange Online or on-premise Exchange permissions.

Requirements:
- Exchange Management Shell or appropriate PowerShell modules for managing Exchange 
  Online or on-premise Exchange Server.
#>

# Define the list of rooms and their mailboxes
$mailboxes = @(
    "meetingroom@email.com",
	"boardroom@email.com"
)

# Iterate through each mailbox and grant "Editor" access to User1 and User2
foreach ($mailbox in $mailboxes) {
    Add-MailboxFolderPermission -Identity "${mailbox}:\calendar" -User user1@email.com -AccessRights Editor
    Add-MailboxFolderPermission -Identity "${mailbox}:\calendar" -User user2@email.com -AccessRights Editor
}