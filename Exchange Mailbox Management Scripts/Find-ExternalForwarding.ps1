# ============================================================================
# Exchange Online External Forwarding Rule Scanner
# ============================================================================
# This script scans all mailboxes for inbox rules that forward email
# externally and generates a comprehensive report.
# ============================================================================

# Configuration
$organizationDomains = @("contoso.com", "contoso.co.uk")  # Replace with your organization's domains
$exportPath = "ExternalForwardingRules_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$throttleLimit = 10  # Number of concurrent operations

# ============================================================================
# Functions
# ============================================================================

function Parse-ForwardingAddress {
    param([string]$Address)
    
    if ([string]::IsNullOrWhiteSpace($Address)) {
        return ""
    }
    
    # Extract email from Exchange format
    if ($Address -match '\[([^\]]+)\]') {
        return $Matches[1]
    }
    
    return $Address
}

function Test-ExternalAddress {
    param([string]$EmailAddress, [array]$InternalDomains)
    
    if ([string]::IsNullOrWhiteSpace($EmailAddress)) {
        return $false
    }
    
    # Check if email domain matches any internal domain
    foreach ($domain in $InternalDomains) {
        if ($EmailAddress -like "*@$domain") {
            return $false
        }
    }
    
    return $true
}

# ============================================================================
# Main Script
# ============================================================================

Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host "Exchange Online External Forwarding Rule Scanner" -ForegroundColor Cyan
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host ""

# Retrieve all mailboxes
Write-Host "Retrieving all mailboxes..." -ForegroundColor Yellow
$mailboxes = Get-Mailbox -ResultSize Unlimited

$userMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" }
$sharedMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }

Write-Host "Found $($userMailboxes.Count) user mailboxes" -ForegroundColor Green
Write-Host "Found $($sharedMailboxes.Count) shared mailboxes" -ForegroundColor Green
Write-Host "Total: $($mailboxes.Count) mailboxes" -ForegroundColor Green
Write-Host ""

# Process mailboxes with throttling
Write-Host "Scanning for forwarding rules (this may take several minutes)..." -ForegroundColor Yellow
$startTime = Get-Date

$results = @()
$processedCount = 0
$totalCount = $mailboxes.Count

# Process in batches for better performance
$batchSize = $throttleLimit
for ($i = 0; $i -lt $mailboxes.Count; $i += $batchSize) {
    $batch = $mailboxes[$i..[Math]::Min($i + $batchSize - 1, $mailboxes.Count - 1)]
    
    $batchResults = $batch | ForEach-Object -ThrottleLimit $throttleLimit -Parallel {
        $mailbox = $_
        $domains = $using:organizationDomains
        
        try {
            $rules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction Stop
            
            $mailboxResults = @()
            
            foreach ($rule in $rules) {
                if ($rule.ForwardTo -or $rule.ForwardAsAttachmentTo -or $rule.RedirectTo) {
                    
                    $allForwardingAddresses = @()
                    
                    # Collect all forwarding addresses
                    if ($rule.ForwardTo) {
                        $allForwardingAddresses += $rule.ForwardTo | ForEach-Object { $_.ToString() }
                    }
                    if ($rule.ForwardAsAttachmentTo) {
                        $allForwardingAddresses += $rule.ForwardAsAttachmentTo | ForEach-Object { $_.ToString() }
                    }
                    if ($rule.RedirectTo) {
                        $allForwardingAddresses += $rule.RedirectTo | ForEach-Object { $_.ToString() }
                    }
                    
                    $allForwardingAddresses = $allForwardingAddresses | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    
                    # Check if any address is external
                    $hasExternalForwarding = $false
                    $parsedAddresses = @()
                    
                    foreach ($addr in $allForwardingAddresses) {
                        # Extract email from Exchange format
                        $email = if ($addr -match '\[([^\]]+)\]') {
                            $Matches[1]
                        } else {
                            $addr
                        }
                        
                        $parsedAddresses += $email
                        
                        # Check if external
                        $isExternal = $true
                        foreach ($domain in $domains) {
                            if ($email -like "*@$domain") {
                                $isExternal = $false
                                break
                            }
                        }
                        
                        if ($isExternal) {
                            $hasExternalForwarding = $true
                        }
                    }
                    
                    # Only include rules with external forwarding
                    if ($hasExternalForwarding) {
                        $mailboxResults += [PSCustomObject]@{
                            MailboxType = $mailbox.RecipientTypeDetails
                            Mailbox = $mailbox.DisplayName
                            UserPrincipalName = $mailbox.UserPrincipalName
                            RuleName = $rule.Name
                            RuleEnabled = $rule.Enabled
                            ForwardingTo = ($parsedAddresses -join "; ")
                            ExternalForwarding = "Yes"
                            RuleDescription = $rule.Description
                        }
                    }
                }
            }
            
            return $mailboxResults
        }
        catch {
            Write-Warning "Error processing mailbox: $($mailbox.UserPrincipalName) - $($_.Exception.Message)"
            return $null
        }
    }
    
    $results += $batchResults | Where-Object { $_ -ne $null }
    
    $processedCount += $batch.Count
    $percentComplete = [Math]::Round(($processedCount / $totalCount) * 100, 1)
    Write-Host "Progress: $processedCount / $totalCount mailboxes processed ($percentComplete%)" -ForegroundColor Gray
}

$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host ""
Write-Host "Scan completed in $([math]::Round($duration.TotalMinutes, 2)) minutes" -ForegroundColor Green
Write-Host ""

# Display results
if ($results.Count -gt 0) {
    $userRules = $results | Where-Object { $_.MailboxType -eq "UserMailbox" }
    $sharedRules = $results | Where-Object { $_.MailboxType -eq "SharedMailbox" }
    
    Write-Host "============================================================================" -ForegroundColor Yellow
    Write-Host "RESULTS SUMMARY" -ForegroundColor Yellow
    Write-Host "============================================================================" -ForegroundColor Yellow
    Write-Host "Total external forwarding rules found: $($results.Count)" -ForegroundColor Yellow
    Write-Host "  - User mailboxes: $($userRules.Count)" -ForegroundColor Yellow
    Write-Host "  - Shared mailboxes: $($sharedRules.Count)" -ForegroundColor Yellow
    Write-Host ""
    
    if ($userRules.Count -gt 0) {
        Write-Host "USER MAILBOX FORWARDING RULES:" -ForegroundColor Cyan
        $userRules | Format-Table -AutoSize
    }
    
    if ($sharedRules.Count -gt 0) {
        Write-Host "SHARED MAILBOX FORWARDING RULES:" -ForegroundColor Cyan
        $sharedRules | Format-Table -AutoSize
    }
    
    # Export to CSV
    $results | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "============================================================================" -ForegroundColor Green
    Write-Host "Results exported to: $exportPath" -ForegroundColor Green
    Write-Host "============================================================================" -ForegroundColor Green
}
else {
    Write-Host "============================================================================" -ForegroundColor Green
    Write-Host "No external forwarding rules found - your organization is clean!" -ForegroundColor Green
    Write-Host "============================================================================" -ForegroundColor Green
}

Write-Host ""