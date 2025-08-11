#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Exchange Online Mailbox Storage Usage Report
.DESCRIPTION
    This script connects to Exchange Online and generates a comprehensive report of mailbox storage usage.
    It retrieves storage limits, current usage, calculates percentage full, and checks archive status.
    Results are sorted by percentage used and displays the top most full mailboxes.
.AUTHOR
    System Administrator
.VERSION
    1.1
#>


# Function to convert bytes to GB with proper rounding
function Convert-BytesToGB {
    param([string]$SizeString)
    
    if ([string]::IsNullOrEmpty($SizeString) -or $SizeString -eq "Unlimited") {
        return 0
    }
    
    # Extract numeric value from size string (e.g., "45.2 GB (48,547,123,456 bytes)")
    if ($SizeString -match '[\d,]+\s*bytes') {
        $BytesMatch = $SizeString -match '\(([0-9,]+)\s*bytes\)'
        if ($BytesMatch) {
            $Bytes = [double]($matches[1] -replace ',', '')
            return [math]::Round($Bytes / 1GB, 2)
        }
    }
    
    # Fallback: try to extract GB value directly
    if ($SizeString -match '(\d+\.?\d*)\s*GB') {
        return [math]::Round([double]$matches[1], 2)
    }
    
    return 0
}

# Function to extract storage limit in GB
function Get-StorageLimitGB {
    param([string]$LimitString)
    
    if ([string]::IsNullOrEmpty($LimitString) -or $LimitString -eq "Unlimited") {
        return 100  # Default assumption for unlimited mailboxes
    }
    
    # Extract GB value from limit string
    if ($LimitString -match '(\d+\.?\d*)\s*GB') {
        return [double]$matches[1]
    }
    
    return 50  # Default fallback
}

# Main script execution
Write-Host "Exchange Online Mailbox Storage Usage Report" -ForegroundColor Cyan

try {
    Write-Host "`nRetrieving mailbox information..." -ForegroundColor Yellow
    
    # Get all mailboxes with necessary properties
    Write-Host "Fetching mailbox list and basic properties..." -ForegroundColor Gray
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | 
                 Select-Object DisplayName, PrimarySmtpAddress, ProhibitSendReceiveQuota, ArchiveStatus
    
    Write-Host "Found $($Mailboxes.Count) mailboxes. Gathering detailed statistics..." -ForegroundColor Gray
    
    # Initialize results array
    $Results = @()
    $ProcessedCount = 0
    
    # Process each mailbox
    foreach ($Mailbox in $Mailboxes) {
        $ProcessedCount++
        $PercentComplete = [math]::Round(($ProcessedCount / $Mailboxes.Count) * 100, 1)
        
        Write-Progress -Activity "Processing Mailboxes" -Status "Processing: $($Mailbox.DisplayName)" -PercentComplete $PercentComplete
        
        try {
            # Get mailbox statistics
            $MailboxStats = Get-MailboxStatistics -Identity $Mailbox.PrimarySmtpAddress -ErrorAction Stop
            
            # Extract and convert values
            $StorageLimitGB = Get-StorageLimitGB -LimitString $Mailbox.ProhibitSendReceiveQuota.ToString()
            $CurrentSizeGB = Convert-BytesToGB -SizeString $MailboxStats.TotalItemSize.ToString()
            
            # Calculate percentage full
            if ($StorageLimitGB -gt 0) {
                $PercentageFull = [math]::Round(($CurrentSizeGB / $StorageLimitGB) * 100, 2)
            } else {
                $PercentageFull = 0
            }
            
            # Determine archive status
            $ArchiveStatus = if ($Mailbox.ArchiveStatus -eq "Active") { "Enabled" } else { "Disabled" }
            
            # Create result object
            $ResultObject = [PSCustomObject]@{
                'Display Name'        = $Mailbox.DisplayName
                'Email Address'       = $Mailbox.PrimarySmtpAddress.ToString()
                'Storage Limit (GB)'  = $StorageLimitGB
                'Current Size (GB)'   = $CurrentSizeGB
                'Percentage Full'     = $PercentageFull
                'Archive Status'      = $ArchiveStatus
            }
            
            $Results += $ResultObject
        }
        catch {
            Write-Warning "Failed to process mailbox: $($Mailbox.DisplayName) - $($_.Exception.Message)"
        }
    }
    
    Write-Progress -Activity "Processing Mailboxes" -Completed
    
    if ($Results.Count -eq 0) {
        Write-Warning "No mailbox data was successfully retrieved."
        return
    }
    
    Write-Host "`nProcessing complete. Generating report..." -ForegroundColor Yellow
    
    # Sort by percentage full (descending) and get mailboxes with >=80% usage
    $TopMailboxes = $Results | Where-Object { $_.'Percentage Full' -ge 80 } | Sort-Object 'Percentage Full' -Descending
    
	if ($TopMailboxes.Count -eq 0) {
        Write-Warning "No mailboxes found with storage usage >= 80%."
    }
    else {
	
    # Display results
    Write-Host "`n" -NoNewline
    Write-Host "TOP 25 MAILBOXES BY STORAGE USAGE" -ForegroundColor Cyan
    Write-Host "Report generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    Write-Host "Total mailboxes processed: $($Results.Count)" -ForegroundColor Gray
    Write-Host ""
    
    # Format and display the table
    $TopMailboxes | Format-Table -AutoSize -Wrap @{
        Expression = 'Display Name'; Width = 25; Alignment = 'Left'
    }, @{
        Expression = 'Email Address'; Width = 35; Alignment = 'Left' 
    }, @{
        Expression = 'Storage Limit (GB)'; Width = 12; Alignment = 'Right'
    }, @{
        Expression = 'Current Size (GB)'; Width = 12; Alignment = 'Right'
    }, @{
        Expression = 'Percentage Full'; Width = 12; Alignment = 'Right'
    }, @{
        Expression = 'Archive Status'; Width = 12; Alignment = 'Center'
    }
}
    # Display summary statistics
    Write-Host "SUMMARY STATISTICS" -ForegroundColor Cyan
    
    $HighUsage = ($Results | Where-Object { $_.'Percentage Full' -ge 80 }).Count
    $MediumUsage = ($Results | Where-Object { $_.'Percentage Full' -ge 50 -and $_.'Percentage Full' -lt 80 }).Count
    $LowUsage = ($Results | Where-Object { $_.'Percentage Full' -lt 50 }).Count
    $ArchiveEnabled = ($Results | Where-Object { $_.'Archive Status' -eq 'Enabled' }).Count
    
    Write-Host "Mailboxes ≥80% full: " -NoNewline -ForegroundColor Red
    Write-Host $HighUsage -ForegroundColor White
    
    Write-Host "Mailboxes 50-79% full: " -NoNewline -ForegroundColor Yellow  
    Write-Host $MediumUsage -ForegroundColor White
    
    Write-Host "Mailboxes <50% full: " -NoNewline -ForegroundColor Green
    Write-Host $LowUsage -ForegroundColor White
    
    Write-Host "Archives enabled: " -NoNewline -ForegroundColor Cyan
    Write-Host $ArchiveEnabled -ForegroundColor White
    
    Write-Host "`nReport completed successfully!" -ForegroundColor Green
	
	$TopMailboxes | Export-Csv -Path "MailboxStorageReport_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation
	Write-Host "Exported CSV File MailboxStorageReport_$(Get-Date -Format 'yyyyMMdd').csv" -ForegroundColor Green
}
catch {
    Write-Error "An error occurred during script execution: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
