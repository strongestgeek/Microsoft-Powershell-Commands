<#
    Get-FullDistributionListMembers.ps1

    Description:
    This script retrieves all members of a specified distribution list, including any 
    nested distribution lists, recursively. It collects the members of the top-level 
    distribution group and drills down into any nested groups to retrieve their members 
    as well. The script exports the results to a CSV file, including details such as 
    member name, email address, recipient type, and the parent group they belong to.

    Usage:
    - Set `$distributionList` to the alias or name of the distribution group you want to 
      start with.
    - Run the script in an Exchange Management Shell or a PowerShell session with the 
      necessary permissions to retrieve distribution group members.
    - The results will be exported to a CSV file named after the specified distribution 
      group (e.g., "YourDistributionList All Members.csv").

    Requirements:
    - Exchange Management Shell or the appropriate PowerShell modules for Exchange Online 
      or on-premise Exchange Server.
    - Sufficient permissions to run `Get-DistributionGroupMember` and query nested groups.

    Output:
    - A CSV file containing the name, email address, recipient type, and parent group of 
      each member retrieved from the distribution group and any nested groups.
#>


# Function to retrieve members of a distribution list recursively
function Get-DistributionGroupMembersRecursively {
    param (
        [string]$DistributionGroup
    )
    
    # Initialize an array to hold all members
    $allMembers = @()

    # Get all members of the distribution list
    $members = Get-DistributionGroupMember -Identity $DistributionGroup -ResultSize Unlimited

    foreach ($member in $members) {
        # Check if the member is another distribution list
        if ($member.RecipientType -eq 'MailUniversalDistributionGroup' -or $member.RecipientType -eq 'MailUniversalSecurityGroup') {
            # Recursively get members of the nested distribution list
            $nestedMembers = Get-DistributionGroupMembersRecursively -DistributionGroup $member.Alias
            $allMembers += $nestedMembers
        } else {
            # Add the non-distribution list member to the result
            $allMembers += New-Object PSObject -Property @{
                Name = $member.DisplayName
                EmailAddress = $member.PrimarySmtpAddress
                RecipientType = $member.RecipientType
                ParentGroup = $DistributionGroup
            }
        }
    }

    return $allMembers
}

# Define the distribution list to start with
$distributionList = "YourDistributionList" # Enter your target distribution list here

# Call the recursive function to get all members
$allMembers = Get-DistributionGroupMembersRecursively -DistributionGroup $distributionList

# Define the CSV file name with the distribution list name
$csvFileName = "$distributionList All Members.csv"

# Export to CSV
$allMembers | Select-Object Name, EmailAddress, RecipientType, ParentGroup | Export-Csv -Path $csvFileName -NoTypeInformation

Write-Host "Export completed. Members are saved in $csvFileName"
