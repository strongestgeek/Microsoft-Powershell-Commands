# 1. Connect to Exchange Online
Connect-ExchangeOnline

# --- CONFIGURATION ---
# Enter the email address of the Distribution List here:
$dlIdentity = "distribution-list@email.com"
# Enter the path to your new CSV file:
$csvPath = ".\members.csv"
# ---------------------

# 2. SAFETY FIRST: Backup the CURRENT list before making changes
# This saves the old list to "backup_old_members.csv" in the same folder
Write-Host "Backing up current membership..." -ForegroundColor Cyan
Get-DistributionGroupMember -Identity $dlIdentity | Select-Object PrimarySmtpAddress | Export-Csv ".\backup_old_members.csv" -NoTypeInformation
Write-Host "Backup Complete. Saved to backup_old_members.csv" -ForegroundColor Green

# 3. Import the NEW list
$newData = Import-Csv $csvPath

# 4. OVERWRITE the Distribution List
# This command replaces the entire membership with the list from your CSV
Write-Host "Updating Distribution List..." -ForegroundColor Cyan
Update-DistributionGroupMember -Identity $dlIdentity -Members $newData.Members -Confirm:$false

# 5. Verify the count
$finalCount = (Get-DistributionGroupMember -Identity $dlIdentity).Count
Write-Host "Success! The Distribution List now has $finalCount members." -ForegroundColor Green
