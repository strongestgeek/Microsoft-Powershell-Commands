# Script Name: Export-CalendarPermissions.ps1
# Purpose: Retrieves calendar access rights for "Default" and "Anonymous" users for all mailboxes and exports to CSV

# Ensure the Exchange Online module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Error "ExchangeOnlineManagement module not installed. Install it using: Install-Module -Name ExchangeOnlineManagement"
    exit
}

# Connect to Exchange Online
try {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -UserPrincipalName admin@domain.com -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit
}

# Get all mailboxes
Write-Host "Retrieving mailboxes..." -ForegroundColor Cyan
try {
    $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox -ErrorAction Stop
}
catch {
    Write-Error "Failed to retrieve mailboxes: $_"
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

$totalMailboxes = $mailboxes.Count
$currentMailbox = 0
$results = @()

# Loop through each mailbox to check calendar permissions for Default and Anonymous
foreach ($mailbox in $mailboxes) {
    $currentMailbox++
    $percentComplete = [math]::Round(($currentMailbox / $totalMailboxes) * 100, 2)
    Write-Host "Processing $($mailbox.UserPrincipalName) ($currentMailbox of $totalMailboxes) [$percentComplete%]" -ForegroundColor Cyan

    $calendar = "$($mailbox.UserPrincipalName):\Calendar"
    try {
        $permissions = Get-MailboxFolderPermission -Identity $calendar -ErrorAction SilentlyContinue | 
                       Where-Object { $_.User -like "Default" -or $_.User -like "Anonymous" }
        
        if ($permissions) {
            foreach ($permission in $permissions) {
                $results += [PSCustomObject]@{
                    Mailbox       = $mailbox.UserPrincipalName
                    User          = $permission.User
                    AccessRights  = $permission.AccessRights -join ", "
                }
                
            }
        }
        else {
            
        }
    }
    catch {
        $errorMsg = "Error processing ${calendar}: $_"
        Write-Host $errorMsg
    }
}

# Export results to CSV
$csvFile = ".\CalendarPermissions_Default_Anonymous_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
if ($results) {
    Write-Host "Exporting results to $csvFile" -ForegroundColor Green
    $results | Export-Csv -Path $csvFile -NoTypeInformation
}
else {
    Write-Host "No calendars with Default or Anonymous permissions found." -ForegroundColor Yellow
}

# Display results in console
if ($results) {
    Write-Host "Results found:" -ForegroundColor Green
    $results | Format-Table -AutoSize
}

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Script completed. Check CSV at $csvFile (if results were found)." -ForegroundColor Green