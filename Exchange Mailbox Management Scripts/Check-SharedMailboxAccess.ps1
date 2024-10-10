<#
    Check-SharedMailboxAccess.ps1

    Description:
    This script checks if a specified user has "Full Access" permissions to any shared 
    mailboxes in an Exchange environment. It retrieves shared mailboxes and their 
    permissions, displaying a list of those to which the user has access. If the user 
    has access to any shared mailboxes, the script prompts whether to save the results 
    to a CSV file.

    Usage:
    - Set the `$UserPrincipalName` variable to the email address of the user you want 
      to check (e.g., "user@email.com").
    - Run the script in an Exchange Management Shell or a PowerShell session with 
      the necessary permissions to query mailbox permissions.
    - The results will be displayed in a table format showing the name and email address 
      of each shared mailbox the user has access to.
    - The script will prompt whether to save the results as a CSV file.

    Requirements:
    - Exchange Management Shell or the appropriate PowerShell modules for Exchange Online 
      or on-premise Exchange Server.
    - Sufficient permissions to run `Get-Mailbox` and `Get-MailboxPermission`.

    Output:
    - A table of shared mailboxes the user has access to, and optionally a CSV file 
      saved in the current directory containing the same information.
#>

$UserPrincipalName = "user@email.com"
#Please enter the email address of the user you want to check above here.	
$CleanEmail = ($UserPrincipalName -replace '[^\w\s]','').Replace(' ','')

$SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize unlimited -ErrorAction SilentlyContinue | Get-MailboxPermission -ErrorAction SilentlyContinue | where {$_.user.tostring() -eq $UserPrincipalName -and $_.IsInherited -eq $false -and $_.AccessRights -match "FullAccess"} | Select-Object -Unique Identity -ErrorAction SilentlyContinue

if ($SharedMailboxes -eq $null) {
    Write-Host "$UserPrincipalName has no access to any shared mailboxes"
	#This just gives you a message so you know that the command didnt find any shared mailboxes rather than just ending the process.																															 
}
else {
    $MailboxInfo = $SharedMailboxes | ForEach-Object {
        $Mailbox = Get-Mailbox $_.Identity
        New-Object PSObject -Property @{
            Name = $Mailbox.Name #This gives you the name of the shared mailbox.
            EmailAddress = $Mailbox.PrimarySmtpAddress #And this gives you the email address of the shared mailbox.
        }
    }

    $MailboxInfo | ft Name, EmailAddress

    $SaveCsv = Read-Host "Would you like to save the results to a CSV file? (Y/N)"
	#If the command has found shared mailboxes then it will ask if you would also like to save the results as a .csv file.
    if ($SaveCsv.ToUpper() -eq "Y") {
        $CsvPath = ".\$CleanEmail.csv"
		$MailboxInfo | Export-Csv $CsvPath -NoTypeInformation
        Write-Host "Results saved to this working directory."
	#This saves the output as a .csv file, it will be in the same location that you ran this .ps1 from and it will be called 'the users email'.csv 
    }
	if ($SaveCsv.ToUpper() -eq "N") {
		Write-Host "K then hun, you do you." #For the lolz.
	}
}
