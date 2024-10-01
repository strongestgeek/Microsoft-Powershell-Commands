# List_SharedMailbox_Access
List all shared mailboxes a user has access to.
DFW 06/06/2023

Hello and welcome to the ReadMe file for this PowerShell command.

All this does is list all shared mailboxes that a defined user has access too.
All you need to do is copy the .ps2 file to your local computer, edit it with notepad and add in the email of the user you want to find out what access they have.

For example, let's pretend we are looking at what access to shared mailboxes I have.
We would open up the .ps1 file and add my email address where it says "$UserPrincipalName"

So it would look like this:

$UserPrincipalName = "user.name@email.com"

Then save and run the command from PowerShell, please note it will take a few minutes before giving you the list.
