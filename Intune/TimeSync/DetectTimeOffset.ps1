<#
.SYNOPSIS
This detection script checks if the system time and NTP server time are more than two minutes out of sync.

.DESCRIPTION
This PowerShell script checks the availability of the NTP server and the time offset between the system clock and the server. If the server is reachable and the offset is more than two minutes, it returns a non-compliant status.
#>

# Set your NTP server
$ntpServer = "time.windows.com" # You can change this to your preferred NTP server

# Function to check if the NTP server is reachable
function Test-NtpServerReachable {
    param (
        [string]$ntpServer
    )

    try {
        $null = Test-Connection -ComputerName $ntpServer -Count 2 -ErrorAction Stop
        return $true
    } catch {
        Write-Host "Failed to reach NTP server '$ntpServer'. Error: $_"
        return $false
    }
}

# Function to get the time offset between the system clock and an NTP server
function Get-TimeOffset {
    param (
        [string]$ntpServer
    )

    try {
        $response = Invoke-RestMethod -Uri "http://$ntpServer" -TimeoutSec 5 -ErrorAction Stop
        $ntpTime = [System.DateTime]::ParseExact($response, 'yyyy-MM-ddTHH:mm:ss.fffffffZ', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        Write-Host "Failed to retrieve NTP time from '$ntpServer'. Error: $_"
        return $null
    }

    $systemTime = Get-Date

    return ($ntpTime - $systemTime).TotalMinutes
}

# Check if the NTP server is reachable
if (Test-NtpServerReachable -ntpServer $ntpServer) {
    # Check if the time offset is more than two minutes
    $timeOffset = Get-TimeOffset -ntpServer $ntpServer
    if ($null -ne $timeOffset -and $timeOffset -gt 2) {
        # Return non-compliant status
        exit 1
    }
}

# Return compliant status
exit 0
