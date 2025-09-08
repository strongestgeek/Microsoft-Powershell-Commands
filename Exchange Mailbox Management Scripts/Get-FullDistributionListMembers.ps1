<#
    Get-FullDistributionListMembers.ps1

    Description:
    This script provides a comprehensive view of email flow within an organisation by:
    - Retrieving all members of distribution lists (including nested ones)
    - Identifying shared mailboxes and their delegates
    - Tracking forward rules (both internal and external)
    - Mapping the complete path an email would take

    Usage:
    - Run the script in Exchange Management Shell
    - Enter the target email address when prompted
    - Results will be exported to a CSV file

    Requirements:
    - Exchange Management Shell or Exchange Online PowerShell modules
    - Appropriate permissions to query distribution groups, mailboxes, and forward rules
#>

# Function to get mailbox type (User/Shared/Distribution List)
function Get-MailboxType {
    param (
        [string]$EmailAddress
    )
    
    try {
        $recipient = Get-Recipient -Identity $EmailAddress -ErrorAction Stop
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue

        if ($recipient.RecipientType -eq 'MailUniversalDistributionGroup' -or 
            $recipient.RecipientType -eq 'MailUniversalSecurityGroup') {
            return "Distribution List"
        }
        elseif ($mailbox.IsShared) {
            return "Shared Mailbox"
        }
        else {
            return "User Mailbox"
        }
    }
    catch {
        return "External Email"
    }
}

# Function to get mailbox delegates and forwards
function Get-MailboxDelegatesAndForwards {
    param (
        [string]$EmailAddress,
        [string]$PathSoFar,
        [System.Collections.ArrayList]$VisitedMailboxes
    )

    if ($null -eq $VisitedMailboxes) {
        $VisitedMailboxes = New-Object System.Collections.ArrayList
    }

    $results = @()
    
    # Skip if we've already processed this mailbox
    if ($VisitedMailboxes -contains $EmailAddress) {
        return $results
    }
    
    [void]$VisitedMailboxes.Add($EmailAddress)
    
    try {
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
        
        if ($mailbox) {
            # Get delegates (Full Access permissions)
            $delegates = Get-MailboxPermission -Identity $EmailAddress | 
                Where-Object {
                    $_.AccessRights -contains "FullAccess" -and 
                    $_.User -notlike "NT AUTHORITY\*" -and 
                    $_.User -ne "Organization Management"
                }
            
            foreach ($delegate in $delegates) {
                $delegateEmail = (Get-Recipient $delegate.User -ErrorAction SilentlyContinue).PrimarySmtpAddress
                if ($delegateEmail) {
                    $results += [PSCustomObject]@{
                        Name = (Get-Recipient $delegate.User).DisplayName
                        EmailAddress = $delegateEmail
                        Type = Get-MailboxType -EmailAddress $delegateEmail
                        Relationship = "Delegate"
                        Path = "$PathSoFar -> $delegateEmail (Delegate)"
                    }
                }
            }

            # Get forwarding rules
            if ($mailbox.ForwardingSmtpAddress) {
                $forwardEmail = $mailbox.ForwardingSmtpAddress -replace "SMTP:"
                $results += [PSCustomObject]@{
                    Name = $forwardEmail
                    EmailAddress = $forwardEmail
                    Type = Get-MailboxType -EmailAddress $forwardEmail
                    Relationship = "Forward"
                    Path = "$PathSoFar -> $forwardEmail (Forward)"
                }

                # Recursively check forwarding destination if not already visited
                if ($VisitedMailboxes -notcontains $forwardEmail) {
                    $subResults = Get-MailboxDelegatesAndForwards -EmailAddress $forwardEmail `
                        -PathSoFar "$PathSoFar -> $forwardEmail" `
                        -VisitedMailboxes $VisitedMailboxes
                    $results += $subResults
                }
            }
        }
    }
    catch {
        Write-Warning "Error processing mailbox $EmailAddress : $_"
    }

    return $results
}

# Function to process distribution group flow
function Get-DistributionFlow {
    param (
        [string]$DistributionGroup,
        [string]$PathSoFar,
        [System.Collections.ArrayList]$VisitedMailboxes
    )

    if ($null -eq $VisitedMailboxes) {
        $VisitedMailboxes = New-Object System.Collections.ArrayList
    }

    $results = @()
    
    try {
        $members = Get-DistributionGroupMember -Identity $DistributionGroup -ResultSize Unlimited
        
        foreach ($member in $members) {
            $currentPath = if ($PathSoFar) { "$PathSoFar -> $($member.PrimarySmtpAddress)" } 
                          else { $member.PrimarySmtpAddress }
            
            $memberType = Get-MailboxType -EmailAddress $member.PrimarySmtpAddress
            
            $results += [PSCustomObject]@{
                Name = $member.DisplayName
                EmailAddress = $member.PrimarySmtpAddress
                Type = $memberType
                Relationship = "Member"
                Path = $currentPath
            }

            # If it's a distribution group, process its members
            if ($memberType -eq "Distribution List") {
                $subResults = Get-DistributionFlow -DistributionGroup $member.PrimarySmtpAddress `
                    -PathSoFar $currentPath `
                    -VisitedMailboxes $VisitedMailboxes
                $results += $subResults
            }
            # If it's a mailbox, get delegates and forwards
            else {
                $delegateResults = Get-MailboxDelegatesAndForwards -EmailAddress $member.PrimarySmtpAddress `
                    -PathSoFar $currentPath `
                    -VisitedMailboxes $VisitedMailboxes
                $results += $delegateResults
            }
        }
    }
    catch {
        Write-Warning "Error processing distribution group $DistributionGroup : $_"
    }

    return $results
}

# Main script
try {
    # Get email address from user
    $startDL = Read-Host "Enter the email address to analyse"
    
    Write-Host "Analysing email flow for $startDL..."
    
    # Initialise visited mailboxes tracking
    $VisitedMailboxes = New-Object System.Collections.ArrayList

    # Get the initial type
    $initialType = Get-MailboxType -EmailAddress $startDL

    # Initialise results array with the starting point
    $allResults = @([PSCustomObject]@{
        Name = (Get-Recipient $startDL).DisplayName
        EmailAddress = $startDL
        Type = $initialType
        Relationship = "Starting Point"
        Path = $startDL
    })

    # Process based on type
    if ($initialType -eq "Distribution List") {
        $allResults += Get-DistributionFlow -DistributionGroup $startDL `
            -VisitedMailboxes $VisitedMailboxes
    }
    else {
        $allResults += Get-MailboxDelegatesAndForwards -EmailAddress $startDL `
            -PathSoFar $startDL `
            -VisitedMailboxes $VisitedMailboxes
    }

    # Export results to CSV
    $csvFileName = "$startDL - Email Flow Analysis.csv"
    $allResults | Select-Object Name, EmailAddress, Type, Relationship, Path | 
        Export-Csv -Path $csvFileName -NoTypeInformation

    Write-Host "Analysis completed. Results saved to $csvFileName"
    Write-Host "Total items processed: $($allResults.Count)"
}
catch {
    Write-Error "An error occurred: $_"
}
