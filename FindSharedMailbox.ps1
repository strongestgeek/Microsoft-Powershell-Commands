#This command shows all shared mailboxes a single user has access too.
#I would recomend you copy this command to a location like C:\Temp and run is from there.
#Example commans: "C:\Temp\ThisCommand.ps1"	

$UserPrincipalName = "user.name@email.com"
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
