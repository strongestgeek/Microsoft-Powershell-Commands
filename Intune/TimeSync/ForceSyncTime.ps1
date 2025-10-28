# Set your NTP server
$ntpServer = "time.windows.com"

try {
    # Stop Windows Time service
    Stop-Service w32time -ErrorAction SilentlyContinue

    # Configure the NTP server
    w32tm /config /manualpeerlist:$ntpServer /syncfromflags:manual /reliable:YES /update

    # Start Windows Time service
    Start-Service w32time

    # Force an immediate sync
    w32tm /resync /force

    # Wait a moment for sync to complete
    Start-Sleep -Seconds 3

    # Verify the sync was successful
    $timeSource = w32tm /query /source
    if ($timeSource -like "*$ntpServer*" -or $timeSource -notlike "*Local CMOS Clock*") {
        exit 0  # Success
    } else {
        exit 1  # Failed to sync
    }

} catch {
    exit 1  # Error occurred
}