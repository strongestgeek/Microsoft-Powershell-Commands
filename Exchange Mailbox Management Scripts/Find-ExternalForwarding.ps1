# ============================================================================
# Exchange Online External Forwarding Rule Scanner (Single-threaded version)
# ============================================================================
# Scans all mailboxes for inbox rules that forward email externally
# and exports a report.
# ============================================================================

# Configuration
$organizationDomains = @("contoso.com", "contoso.co.uk")  # <-- Replace with your domains
$exportPath = "ExternalForwardingRules_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# ============================================================================
# Ensure Exchange Online Connection
# ============================================================================

if (-not (Get-Module ExchangeOnlineManagement)) {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}

if (-not (Get-PSSession | Where-Object { $_.ComputerName -like '*outlook.office365.com*' })) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -UserPrincipalName "your.admin@contoso.com"
}

# ============================================================================
# Main Script
# ============================================================================

Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host "Exchange Online External Forwarding Rule Scanner" -ForegroundColor Cyan
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Retrieving all mailboxes..." -ForegroundColor Yellow
$mailboxes = Get-Mailbox -ResultSize Unlimited

$userMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" }
$sharedMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }

Write-Host "Found $($userMailboxes.Count) user mailboxes" -ForegroundColor Green
Write-Host "Found $($sharedMailboxes.Count) shared mailboxes" -ForegroundColor Green
Write-Host "Total: $($mailboxes.Count) mailboxes" -ForegroundColor Green
Write-Host ""

Write-Host "Scanning for forwarding rules (this may take several minutes)..." -ForegroundColor Yellow
$startTime = Get-Date

$results = @()
$processedCount = 0
$totalCount = $mailboxes.Count

foreach ($mailbox in $mailboxes) {
    try {
        $rules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction Stop

        foreach ($rule in $rules) {
            if ($rule.ForwardTo -or $rule.ForwardAsAttachmentTo -or $rule.RedirectTo) {

                $allForwardingAddresses = @()
                if ($rule.ForwardTo) { $allForwardingAddresses += $rule.ForwardTo | ForEach-Object { $_.ToString() } }
                if ($rule.ForwardAsAttachmentTo) { $allForwardingAddresses += $rule.ForwardAsAttachmentTo | ForEach-Object { $_.ToString() } }
                if ($rule.RedirectTo) { $allForwardingAddresses += $rule.RedirectTo | ForEach-Object { $_.ToString() } }

                $allForwardingAddresses = $allForwardingAddresses | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

                $parsedAddresses = @()
                $hasExternalForwarding = $false

                foreach ($addr in $allForwardingAddresses) {
                    $email = if ($addr -match '\[([^\]]+)\]') { $Matches[1] } else { $addr }
                    $parsedAddresses += $email

                    $isExternal = $true
                    foreach ($domain in $organizationDomains) {
                        if ($email -like "*@$domain") {
                            $isExternal = $false
                            break
                        }
                    }

                    if ($isExternal) { $hasExternalForwarding = $true }
                }

                if ($hasExternalForwarding) {
                    $results += [PSCustomObject]@{
                        MailboxType       = $mailbox.RecipientTypeDetails
                        Mailbox           = $mailbox.DisplayName
                        UserPrincipalName = $mailbox.UserPrincipalName
                        RuleName          = $rule.Name
                        RuleEnabled       = $rule.Enabled
                        ForwardingTo      = ($parsedAddresses -join "; ")
                        ExternalForwarding = "Yes"
                        RuleDescription   = $rule.Description
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Error processing mailbox: $($mailbox.UserPrincipalName) - $($_.Exception.Message)"
    }

    $processedCount++
    $percentComplete = [Math]::Round(($processedCount / $totalCount) * 100, 1)
    Write-Host "Progress: $processedCount / $totalCount mailboxes processed ($percentComplete%)" -ForegroundColor Gray
}

$endTime = Get-Date
$duration = $endTime - $startTime
Write-Host ""
Write-Host "Scan completed in $([math]::Round($duration.TotalMinutes, 2)) minutes" -ForegroundColor Green
Write-Host ""

# ============================================================================
# Output Results
# ============================================================================

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