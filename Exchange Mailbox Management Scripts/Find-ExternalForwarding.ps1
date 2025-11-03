Write-Host "Retrieving all mailboxes..." -ForegroundColor Cyan
$mailboxes = Get-Mailbox -ResultSize Unlimited

$results = @()
$count = 0
$total = $mailboxes.Count

foreach ($mailbox in $mailboxes) {
    $count++
    Write-Progress -Activity "Scanning mailboxes for forwarding rules" -Status "Processing $($mailbox.DisplayName)" -PercentComplete (($count / $total) * 100)
    try {
        $rules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
            # Check each rule for forwarding actions
            foreach ($rule in $rules) {
                if ($rule.ForwardTo -or $rule.ForwardAsAttachmentTo -or $rule.RedirectTo) {
                    $forwardingAddresses = @()
                    if ($rule.ForwardTo) {
                        $forwardingAddresses += $rule.ForwardTo
                    }
                    if ($rule.ForwardAsAttachmentTo) {
                        $forwardingAddresses += $rule.ForwardAsAttachmentTo
                    }
                    if ($rule.RedirectTo) {
                        $forwardingAddresses += $rule.RedirectTo
                    }
                    $results += [PSCustomObject]@{
                        Mailbox = $mailbox.DisplayName
                        UserPrincipalName = $mailbox.UserPrincipalName
                        RuleName = $rule.Name
                        RuleEnabled = $rule.Enabled
                        ForwardingTo = ($forwardingAddresses -join "; ")
                        RuleDescription = $rule.Description
                    }
                }
            }
        }
        catch {
            Write-Warning "Error processing mailbox: $($mailbox.UserPrincipalName) - $($_.Exception.Message)"
        }
    }
Write-Progress -Activity "Scanning mailboxes for forwarding rules" -Completed

# Display results
if ($results.Count -gt 0) {
    Write-Host "`nFound $($results.Count) forwarding rule(s) across $($mailboxes.Count) mailboxes:" -ForegroundColor Yellow
    $results | Format-Table -AutoSize

    # Export to CSV
    $exportPath = "ForwardingRules_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $results | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "`nResults exported to: $exportPath" -ForegroundColor Green
}
else {
    Write-Host "`nNo forwarding rules found." -ForegroundColor Green
}