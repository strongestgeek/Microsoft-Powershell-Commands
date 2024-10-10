<#
    Get-MailboxDistributionGroups.ps1

    Description:
    This script retrieves a list of distribution groups that a specified mailbox is a member of 
    in an Exchange environment. It checks each distribution group for the given mailbox's 
    membership and outputs the group's name and primary SMTP address.

    Usage:
    - Replace `$SharedMailbox` with the email address of the mailbox you want to check.
    - Run the script in an Exchange Management Shell or a PowerShell session with the necessary 
      permissions to query distribution group membership.
    - If the mailbox is found in any distribution groups, the group names and primary SMTP 
      addresses will be displayed. Otherwise, a message indicating no groups were found 
      will be shown.

    Requirements:
    - Exchange Management Shell or the appropriate PowerShell modules for Exchange Online 
      or on-premise Exchange Server.
    - Sufficient permissions to run `Get-DistributionGroup` and `Get-DistributionGroupMember`.

    Output:
    - A list of distribution groups where the specified mailbox is a member, or a message 
      indicating no groups were found.
#>

# Replace 'Mailbox@domain.com' with your mailbox email address
$SharedMailbox = "Mailbox@domain.com"

# Get a list of distribution groups that the mailbox is a member of
$DistributionLists = @()
try {
    $DistributionLists = Get-DistributionGroup -ResultSize Unlimited | Where-Object { 
        $members = Get-DistributionGroupMember $_ -ErrorAction SilentlyContinue
        $members -and ($members.PrimarySmtpAddress -contains $SharedMailbox)
    }
} catch {
    Write-Host "An error occurred while fetching distribution groups: $_"
}

# Display the list of distribution groups
if ($DistributionLists.Count -gt 0) {
    $DistributionLists | Select-Object Name, PrimarySmtpAddress
} else {
    Write-Host "No distribution groups found that the mailbox is a member of."
}