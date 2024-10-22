<#
    Get-DistributionGroupUsage.ps1

    Description:
    This script retrieves the number of messages received by each distribution group 
    in the last 10 days within an Exchange environment. It collects message trace data 
    for all distribution groups and counts the number of emails each group received 
    during the specified time range. The results are then exported to a CSV file.

    Usage:
    - The script defines a time range of the last 10 days from the current date. You can 
      modify the `$StartDate` and `$EndDate` variables to adjust the time range.
    - Run the script in an Exchange Management Shell or a PowerShell session with the 
      necessary permissions to retrieve distribution group details and message trace data.
    - The output is saved as a CSV file named "DistributionGroupEmailCounts.csv" in the 
      current working directory. You can change the path by modifying the `Export-Csv` 
      command.

    Requirements:
    - Exchange Management Shell or the appropriate PowerShell modules for Exchange Online 
      or on-premise Exchange Server.
    - Sufficient permissions to run `Get-DistributionGroup` and `Get-MessageTrace`.

    Output:
    - A CSV file containing the distribution group name, email address, and the number 
      of messages received in the specified time range.
#>

# Define the time range for the last 10 days
$EndDate = (Get-Date)
$StartDate = $EndDate.AddDays(-10)

# Get all distribution groups
$DistributionGroups = Get-DistributionGroup -ResultSize Unlimited

# Initialize an array to hold the results
$Results = @()

foreach ($Group in $DistributionGroups) {
    # Get the email address of the distribution group
    $GroupEmail = $Group.PrimarySmtpAddress

    # Get message trace data for this distribution group within the last 10 days
    $MessageTraces = Get-MessageTrace -RecipientAddress $GroupEmail -StartDate $StartDate -EndDate $EndDate -PageSize 5000

    # Count the number of messages received by this distribution group
    $EmailCount = $MessageTraces.Count

    # Create a custom object to store the result
    $Result = [PSCustomObject]@{
        DistributionGroup = $Group.DisplayName
        EmailAddress      = $GroupEmail
        EmailCount        = $EmailCount
    }

    # Add the result to the array
    $Results += $Result
}

# Export the results to a CSV file
$Results | Export-Csv -Path ".\DistributionGroupEmailCounts.csv" -NoTypeInformation -Encoding UTF8