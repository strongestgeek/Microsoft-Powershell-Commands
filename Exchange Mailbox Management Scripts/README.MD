# Exchange Management Utility Scripts

## Overview

This repository contains a collection of PowerShell scripts designed for managing and auditing Exchange mailboxes and users within an Exchange environment (both on-premise and Exchange Online). These scripts streamline common administrative tasks, providing insights into mailbox sizes, user memberships in distribution groups, forwarding addresses, and more.

## Scripts

### 1. **Check-SharedMailboxAccess.ps1**
- **Description:** Checks if a specified user has "Full Access" permissions to any shared mailboxes.
- **Usage:** Set the `$UserPrincipalName` variable to the user's email address you want to check.

### 2. **Check-EmailAddressLocation.ps1**
- **Description:** Finds the location of a specified email address within the Exchange environment.
- **Usage:** Set the `$emailAddress` variable to the email address you want to search.

### 3. **Check-UserDistributionGroupMembership.ps1**
- **Description:** Checks which distribution groups a specified user is a member of.
- **Usage:** Replace the `$User` variable with the email address or username of the user you want to check.

### 4. **Export-MailboxSizes.ps1**
- **Description:** Retrieves mailbox sizes and associated statistics for all mailboxes in the Exchange environment and exports the results to a CSV file.
- **Usage:** Set the `$outputFile` variable to specify the path for the output CSV file.

### 5. **Check-ForwardingMailboxes.ps1**
- **Description:** Identifies mailboxes with forwarding addresses set to specified users.
- **Usage:** Set the `$forwardingAddress` variable to the email address you want to check.

### 6. **Check-MailboxAlias.ps1**
- **Description:** Checks if a specified email address is an alias for any mailbox.
- **Usage:** Set the `$UserPrincipalName` variable to the user's email address you want to check.

## Requirements

- PowerShell with Exchange Management Shell or appropriate modules for Exchange Online.
- Sufficient permissions to execute the scripts and access mailbox statistics.

## Getting Started

1. Clone this repository to your local machine:
   ```bash
   git clone https://github.com/strongestgeek/Microsoft-PowerShell-Commands.git
