<#
    List-SortedDistributionGroupMembers.ps1

    Description:
    This script connects to Exchange Online and retrieves the members of a specified distribution list, 
    then sorts and displays their email addresses. It's designed for use in environments such as Office 365 
    (Exchange Online).

    Key Features:
    - Connects to Exchange Online using the Exchange Online PowerShell module (if applicable).
    - Retrieves all members of the specified distribution list.
    - Sorts the list of members by their primary email address.
    - Displays sorted email addresses in a readable format.

    Parameters:
    - `$distributionList`: The name or email address of the distribution list to query.

    Requirements:
    - Exchange Online PowerShell module (`ExchangeOnlineManagement`).
    - Admin account access with sufficient permissions to query distribution lists.

    Example Usage:
    - Run the script and replace `"YourDistributionListNameOrEmail"` with the actual name or email address 
      of the distribution list you want to query.

    Notes:
    - Ensure you have the Exchange Online Management module installed (`Install-Module ExchangeOnlineManagement`).
    - Adjust and uncomment the connection block for Exchange Online if using Office 365.
    - Always disconnect from Exchange Online after completing the task to maintain security (`Disconnect-ExchangeOnline`).

#>

# Define the distribution list
$distributionList = "YourDistributionListNameOrEmail"

# Get the members of the distribution list and sort them by email address
$members = Get-DistributionGroupMember -Identity $distributionList | Sort-Object PrimarySmtpAddress

# Output only the sorted email addresses of the members
if ($members) {
    Write-Host "Sorted Email Addresses of Members in Distribution List: $distributionList" -ForegroundColor Green
    foreach ($member in $members) {
        Write-Host "$($member.PrimarySmtpAddress)" -ForegroundColor Yellow
    }
} else {
    Write-Host "No members found in the distribution list: $distributionList" -ForegroundColor Red
}

# Disconnect from Exchange Online (if using Office 365)
# Disconnect-ExchangeOnline -Confirm:$false
