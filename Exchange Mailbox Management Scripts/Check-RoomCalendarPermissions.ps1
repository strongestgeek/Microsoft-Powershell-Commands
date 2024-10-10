<#
    CheckRoomCalendarPermissions.ps1

    Description:
    This script checks the calendar permissions for a list of room mailboxes in an Exchange environment.
    It iterates through each room's calendar and retrieves the current permissions assigned to it.

    Key Features:
    - Defines a list of room mailbox email addresses.
    - Uses `Get-MailboxFolderPermission` to retrieve permissions for each room's calendar folder.
    - Outputs the permissions for each room mailbox in a readable format.

    Usage:
    - Update the `$mailboxes` array with the room mailbox email addresses.
    - The script retrieves and displays the current permissions for the `\calendar` folder of each mailbox.

    Parameters:
    - `$mailboxes`: An array of room mailbox email addresses to check.

    Requirements:
    - Exchange Management Shell (EMS) or equivalent environment.
    - Appropriate permissions to view mailbox folder permissions.

    Output:
    - The script outputs a list of users and their access rights for each room mailbox's calendar.
#>

# Define the list of rooms and their mailboxes
$mailboxes = @(
    "meetingroom@email.com",
	"boardroom@email.com"
)

# Iterate through each mailbox and check calendar permissions
foreach ($mailbox in $mailboxes) {
    Write-Host "Permissions for $mailbox's calendar:"
    Get-MailboxFolderPermission -Identity "${mailbox}:\calendar"
    Write-Host "------------------------------------`n"
}
