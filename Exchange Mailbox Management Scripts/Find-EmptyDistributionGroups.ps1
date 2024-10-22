<#
    Find-EmptyDistributionGroups.ps1

    Description:
    This script identifies and exports all distribution groups in the environment that currently have no members. 
    It connects to Exchange Online (or an on-premises Exchange environment) and scans each distribution group to 
    determine if it is empty.

    Key Features:
    - Connects to the Exchange environment to retrieve all distribution groups.
    - Iterates through each group to check if it has any members.
    - Collects details of groups with zero members.
    - Exports the results to a CSV file (`EmptyDistributionGroups.csv`).

    Parameters:
    - `$emptyGroups`: An array that stores information about distribution groups that have no members.
    - `$grp`: Each distribution group retrieved from the Exchange environment.

    Requirements:
    - Exchange Online or on-premises Exchange environment with sufficient permissions to access 
      distribution group details.
    - Exchange Online PowerShell module (`ExchangeOnlineManagement`) if using Exchange Online.

    Example Usage:
    - Ensure the necessary module is imported and authenticated if using Exchange Online.
    - Run the script, and it will generate a CSV file (`EmptyDistributionGroups.csv`) in the same 
      directory where the script is executed.
    
    Notes:
    - The CSV file contains columns: `DisplayName`, `PrimarySMTPAddress`, and `DistinguishedName`.
    - Adjust the export path if you wish to save the file in a different directory.
    - The script is optimized to handle environments with a large number of distribution groups.

#>

$emptyGroups = foreach ($grp in Get-DistributionGroup -ResultSize Unlimited) {
    if (@(Get-DistributionGroupMember –Identity $grp.DistinguishedName -ResultSize Unlimited).Count –eq 0 ) {
        [PsCustomObject]@{
            DisplayName        = $grp.DisplayName
            PrimarySMTPAddress = $grp.PrimarySMTPAddress
            DistinguishedName  = $grp.DistinguishedName
        }
    }
}
$emptyGroups | Export-Csv '.\EmptyDistributionGroups.csv' -NoTypeInformation