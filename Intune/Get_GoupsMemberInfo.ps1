# Export Intune Security Group Members with Extended Details
# Requires Microsoft.Graph PowerShell modules

# Install required modules if not already installed
# Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
# Install-Module Microsoft.Graph.Groups -Scope CurrentUser
# Install-Module Microsoft.Graph.Users -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All"

# Specify the security group name or ID
$GroupName = "YourSecurityGroupName"  # Change this to your group name
# Or use Group ID directly:
# $GroupId = "your-group-id-here"

# Get the group
if ($GroupName) {
    $Group = Get-MgGroup -Filter "displayName eq '$GroupName'"
    if (-not $Group) {
        Write-Error "Group '$GroupName' not found"
        exit
    }
    $GroupId = $Group.Id
}

Write-Host "Retrieving members from group: $($Group.DisplayName)" -ForegroundColor Green

# Get all group members
$Members = Get-MgGroupMember -GroupId $GroupId -All

# Create array to store results
$Results = @()

# Process each member
$Counter = 0
foreach ($Member in $Members) {
    $Counter++
    Write-Progress -Activity "Processing Members" -Status "Member $Counter of $($Members.Count)" -PercentComplete (($Counter / $Members.Count) * 100)
    
    # Get detailed user information
    $User = Get-MgUser -UserId $Member.Id -Property "DisplayName,Mail,UserPrincipalName,JobTitle,Department,OfficeLocation,Manager"
    
    # Get manager details if manager exists
    $ManagerName = $null
    $ManagerJobTitle = $null
    $ManagerDepartment = $null
    
    if ($User.Manager) {
        try {
            $Manager = Get-MgUserManager -UserId $User.Id
            if ($Manager) {
                $ManagerDetails = Get-MgUser -UserId $Manager.Id -Property "DisplayName,JobTitle,Department"
                $ManagerName = $ManagerDetails.DisplayName
                $ManagerJobTitle = $ManagerDetails.JobTitle
                $ManagerDepartment = $ManagerDetails.Department
            }
        }
        catch {
            Write-Warning "Could not retrieve manager for $($User.DisplayName)"
        }
    }
    
    # Create custom object with all required information
    $UserInfo = [PSCustomObject]@{
        'Name'                    = $User.DisplayName
        'Email Address'           = if ($User.Mail) { $User.Mail } else { $User.UserPrincipalName }
        'Job Title'               = $User.JobTitle
        'Department'              = $User.Department
        'Line Manager'            = $ManagerName
        'Line Manager Job Title'  = $ManagerJobTitle
        'Line Manager Department' = $ManagerDepartment
    }
    
    $Results += $UserInfo
}

# Export to CSV
$ExportPath = ".\IntuneGroupMembers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "`nExport completed successfully!" -ForegroundColor Green
Write-Host "Total members exported: $($Results.Count)" -ForegroundColor Cyan
Write-Host "File saved to: $ExportPath" -ForegroundColor Cyan

# Disconnect from Microsoft Graph
Disconnect-MgGraph