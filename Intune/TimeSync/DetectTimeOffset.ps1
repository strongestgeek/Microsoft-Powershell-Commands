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
        # NTP packet structure (48 bytes)
        $ntpData = New-Object byte[] 48
        $ntpData[0] = 0x1B  # LI = 0 (no warning), VN = 3 (IPv4 only), Mode = 3 (Client Mode)

        # Create UDP client
        $socket = New-Object System.Net.Sockets.Socket([System.Net.Sockets.AddressFamily]::InterNetwork, 
                                                         [System.Net.Sockets.SocketType]::Dgram, 
                                                         [System.Net.Sockets.ProtocolType]::Udp)
        $socket.ReceiveTimeout = 5000
        $socket.SendTimeout = 5000

        # Connect to NTP server on port 123
        $socket.Connect($ntpServer, 123)

        # Send NTP request
        $null = $socket.Send($ntpData)

        # Receive NTP response
        $null = $socket.Receive($ntpData)
        $socket.Close()

        # Extract transmit timestamp (bytes 40-47)
        $intPart = [BitConverter]::ToUInt32($ntpData[43..40], 0)
        $fracPart = [BitConverter]::ToUInt32($ntpData[47..44], 0)

        # Convert to DateTime (NTP epoch is 1900-01-01)
        $ntpEpoch = New-Object DateTime(1900, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
        $ntpTime = $ntpEpoch.AddSeconds($intPart + ($fracPart / [Math]::Pow(2, 32)))

        # Get system time in UTC
        $systemTime = [DateTime]::UtcNow

        # Calculate offset in minutes
        $offsetMinutes = ($ntpTime - $systemTime).TotalMinutes

        return $offsetMinutes

    } catch {
        Write-Host "Failed to retrieve NTP time from '$ntpServer'. Error: $_"
        return $null
    }
}

# Check if the NTP server is reachable
if (Test-NtpServerReachable -ntpServer $ntpServer) {
    # Check if the time offset is more than two minutes
    $timeOffset = Get-TimeOffset -ntpServer $ntpServer
    if ($null -ne $timeOffset) {
        # Check absolute value to detect both positive and negative drift
        if ([Math]::Abs($timeOffset) -gt 2) {
            exit 1
        }
    } else {
        exit 1
    }
}

# Return compliant status
exit 0RestMethod -Uri "http://$ntpServer" -TimeoutSec 5 -ErrorAction Stop
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
