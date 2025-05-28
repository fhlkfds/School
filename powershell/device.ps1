# Google Workspace Device Activity Report Script
# Uses GAM (Google Apps Manager) to retrieve device login activity
# Filters results for logins between May 22, 2025 3:05 PM and May 28, 2025

# Define parameters
param(
    [string]$OutputPath = ".\device_activity_report.csv",
    [string]$GAMPath = "gam"  # Adjust if GAM is not in PATH
)

# Set error handling
$ErrorActionPreference = "Stop"

# Initialize results array
$allResults = @()

# Define date range for filtering
$StartDate = Get-Date "2025-05-22 15:05:00"  # May 22, 2025 at 3:05 PM
$EndDate = Get-Date "2025-05-28 23:59:59"    # May 28, 2025 at end of day

Write-Host "Google Workspace Device Activity Report Generator" -ForegroundColor Cyan
Write-Host "Date Range: $($StartDate.ToString("yyyy-MM-dd HH:mm:ss")) to $($EndDate.ToString("yyyy-MM-dd HH:mm:ss"))" -ForegroundColor Yellow
Write-Host ""

# Function to check if GAM is available and properly authenticated
function Test-GAMAuthentication {
    Write-Host "Checking GAM authentication..." -ForegroundColor Green
    
    try {
        # Test GAM connection by getting OAuth info (this should always work if authenticated)
        $testResult = & $GAMPath oauth info 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            throw "GAM authentication failed or GAM not properly configured"
        }
        
        # Check if we get valid OAuth info
        $oauthInfo = $testResult -join "`n"
        if ($oauthInfo -match "Client OAuth2 File:" -and $oauthInfo -match "Google Workspace Admin:") {
            Write-Host "GAM authentication successful" -ForegroundColor Green
            
            # Extract admin email for verification
            if ($oauthInfo -match "Google Workspace Admin:\s*(.+)") {
                $adminEmail = $matches[1].Trim()
                Write-Host "Authenticated as: $adminEmail" -ForegroundColor Cyan
            }
            
            return $true
        }
        else {
            throw "GAM OAuth information not found or incomplete"
        }
    }
    catch {
        Write-Error "GAM Error: $($_.Exception.Message)"
        Write-Host "Please ensure:" -ForegroundColor Red
        Write-Host "1. GAM is installed and in your PATH" -ForegroundColor Red
        Write-Host "2. You are authenticated with 'gam oauth create'" -ForegroundColor Red
        Write-Host "3. You have appropriate admin permissions" -ForegroundColor Red
        return $false
    }
}

# Function to convert GAM date format to PowerShell DateTime
function ConvertFrom-GAMDate {
    param([string]$GAMDateString)
    
    try {
        # GAM typically returns dates in ISO format or RFC3339
        # Handle multiple possible formats
        if ($GAMDateString -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}') {
            return [DateTime]::Parse($GAMDateString)
        }
        elseif ($GAMDateString -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}') {
            return [DateTime]::ParseExact($GAMDateString, "yyyy-MM-dd HH:mm:ss", $null)
        }
        else {
            # Try generic parsing
            return [DateTime]::Parse($GAMDateString)
        }
    }
    catch {
        Write-Warning "Could not parse date: $GAMDateString"
        return $null
    }
}

# Function to retrieve device activity reports
function Get-DeviceActivityReport {
    Write-Host "Retrieving device activity reports..." -ForegroundColor Green
    
    # Create temporary file for GAM output
    $tempFile = [System.IO.Path]::GetTempFileName()
    
    try {
        # GAM command to get login reports for the specified date range
        $startDateGAM = $StartDate.ToString("yyyy-MM-dd")
        $endDateGAM = $EndDate.ToString("yyyy-MM-dd")
        
        Write-Host "Executing GAM command for login reports from $startDateGAM to $endDateGAM..." -ForegroundColor Yellow
        
        # Get login activity reports using the reports API
        & $GAMPath report logins $startDateGAM $endDateGAM > $tempFile
        
        # Check if the command was successful
        if ($LASTEXITCODE -eq 0 -and (Test-Path $tempFile) -and (Get-Item $tempFile).Length -gt 0) {
            $rawData = Get-Content $tempFile
            Write-Host "Successfully retrieved login reports" -ForegroundColor Green
            return $rawData
        }
        else {
            Write-Host "No login data found for the specified date range" -ForegroundColor Yellow
            return @()
        }
    }
    catch {
        Write-Warning "Failed to retrieve login reports: $($_.Exception.Message)"
        return @()
    }
    finally {
        # Clean up temp file
        if (Test-Path $tempFile) {
            Remove-Item $tempFile -Force
        }
    }
}

# Function to get mobile device details with last sync information
function Get-MobileDeviceActivity {
    Write-Host "Retrieving mobile device activity..." -ForegroundColor Green
    
    try {
        # Get all mobile devices with last sync data
        $tempFile = [System.IO.Path]::GetTempFileName()
        
        # Execute GAM command to get mobile devices
        & $GAMPath print mobile > $tempFile
        
        if (Test-Path $tempFile) {
            $deviceData = Import-Csv $tempFile
            Remove-Item $tempFile -Force
            return $deviceData
        }
        
        return $null
    }
    catch {
        Write-Error "Failed to retrieve mobile device data: $($_.Exception.Message)"
        return $null
    }
}

# Function to process and filter device activity data
function Process-DeviceActivity {
    param([array]$RawData)
    
    Write-Host "Processing device activity data..." -ForegroundColor Green
    
    $filteredResults = @()
    
    # Skip if no data
    if (-not $RawData -or $RawData.Count -eq 0) {
        return $filteredResults
    }
    
    # Find header line to understand the data structure
    $headerLine = $RawData | Where-Object { $_ -like "*date*" -or $_ -like "*time*" -or $_ -like "*user*" } | Select-Object -First 1
    
    foreach ($line in $RawData) {
        # Skip header lines, empty lines, and comments
        if ([string]::IsNullOrWhiteSpace($line) -or 
            $line -like "*Getting*" -or 
            $line -like "*date*" -or 
            $line -like "#*" -or
            $line.StartsWith("Getting")) {
            continue
        }
        
        try {
            # Parse CSV-like data from GAM login reports
            # GAM login reports typically have: date, time, user, event, ip_address, user_agent
            $fields = $line -split ","
            $fields = $fields | ForEach-Object { $_.Trim().Trim('"') }
            
            # Extract relevant information (adjust based on actual GAM output)
            if ($fields.Count -ge 3) {
                # Try to construct datetime from date and time fields
                $dateStr = $fields[0]
                $timeStr = if ($fields.Count -gt 1) { $fields[1] } else { "00:00:00" }
                $userEmail = if ($fields.Count -gt 2) { $fields[2] } else { "Unknown" }
                $eventType = if ($fields.Count -gt 3) { $fields[3] } else { "login" }
                $ipAddress = if ($fields.Count -gt 4) { $fields[4] } else { "Unknown" }
                $userAgent = if ($fields.Count -gt 5) { $fields[5] } else { "Unknown" }
                
                # Try to parse the date/time
                $loginDateTime = $null
                try {
                    # Handle different date formats that GAM might return
                    if ($dateStr -match '^\d{4}-\d{2}-\d{2}' -and $timeStr -match '^\d{2}:\d{2}:\d{2}') {
                        $loginDateTime = [DateTime]::ParseExact("$dateStr $timeStr", "yyyy-MM-dd HH:mm:ss", $null)
                    }
                    elseif ($dateStr -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}') {
                        $loginDateTime = [DateTime]::Parse($dateStr)
                    }
                    else {
                        # Try generic parsing
                        $loginDateTime = [DateTime]::Parse("$dateStr $timeStr")
                    }
                }
                catch {
                    Write-Warning "Could not parse date/time: $dateStr $timeStr"
                    continue
                }
                
                # Filter by date range
                if ($loginDateTime -and $loginDateTime -ge $StartDate -and $loginDateTime -le $EndDate) {
                    $deviceInfo = [PSCustomObject]@{
                        DateTime = $loginDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                        UserEmail = $userEmail
                        EventType = $eventType
                        IPAddress = $ipAddress
                        UserAgent = $userAgent
                        DeviceInfo = if ($userAgent -like "*Mobile*" -or $userAgent -like "*Android*" -or $userAgent -like "*iPhone*") { "Mobile Device" } else { "Desktop/Web" }
                        DeviceId = ""
                        DeviceModel = ""
                        DeviceOS = ""
                        DeviceStatus = ""
                        DeviceIMEI = ""
                        DeviceSerial = ""
                        RawData = $line
                        SortDateTime = $loginDateTime
                    }
                    
                    $filteredResults += $deviceInfo
                }
            }
        }
        catch {
            Write-Warning "Could not process line: $line"
            Write-Warning "Error: $($_.Exception.Message)"
        }
    }
    
    return $filteredResults
}

# Main execution block
try {
    # Check GAM authentication
    if (-not (Test-GAMAuthentication)) {
        exit 1
    }
    
    # Get login activity reports
    Write-Host "Retrieving login activity reports..." -ForegroundColor Green
    $loginData = Get-DeviceActivityReport
    
    if ($loginData -and $loginData.Count -gt 0) {
        $processedLogins = Process-DeviceActivity -RawData $loginData
        $allResults += $processedLogins
        Write-Host "Found $($processedLogins.Count) login events in date range" -ForegroundColor Cyan
    }
    
    # Get mobile device activity (most common for login tracking)
    Write-Host "Retrieving mobile device information..." -ForegroundColor Green
    $mobileDevices = Get-MobileDeviceActivity
    
    if ($mobileDevices) {
        Write-Host "Found $($mobileDevices.Count) mobile devices" -ForegroundColor Green
        
        foreach ($device in $mobileDevices) {
            # Check if device has recent activity within our date range
            if ($device.lastSync) {
                $lastSyncDate = ConvertFrom-GAMDate -GAMDateString $device.lastSync
                
                if ($lastSyncDate -and $lastSyncDate -ge $StartDate -and $lastSyncDate -le $EndDate) {
                    $deviceResult = [PSCustomObject]@{
                        DateTime = $lastSyncDate.ToString("yyyy-MM-dd HH:mm:ss")
                        UserEmail = $device.email
                        EventType = "Device Sync"
                        IPAddress = "N/A"
                        UserAgent = "$($device.type) - $($device.model)"
                        DeviceInfo = $device.type
                        DeviceId = $device.resourceId
                        DeviceModel = $device.model
                        DeviceOS = $device.os
                        DeviceStatus = $device.status
                        DeviceIMEI = $device.imei
                        DeviceSerial = $device.serialNumber
                        RawData = "Mobile Device Sync"
                        SortDateTime = $lastSyncDate
                    }
                    
                    $allResults += $deviceResult
                    Write-Host "Found device activity: $($device.email) - $($device.type) - Last Sync: $($lastSyncDate)" -ForegroundColor Cyan
                }
            }
        }
    }
    
    # Output results
    if ($allResults.Count -gt 0) {
        Write-Host ""
        Write-Host "Found $($allResults.Count) devices with activity in the specified time range" -ForegroundColor Green
        
        # Sort by date/time
        $allResults = $allResults | Sort-Object SortDateTime -Descending
        
        # Export to CSV
        $allResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
        
        # Display summary
        Write-Host ""
        Write-Host "Summary:" -ForegroundColor Yellow
        Write-Host "--------" -ForegroundColor Yellow
        
        $allResults | ForEach-Object {
            Write-Host "$($_.UserEmail) - $($_.DeviceInfo) - $($_.DateTime) - $($_.EventType)" -ForegroundColor White
        }
    }
    else {
        Write-Host "No devices found with activity in the specified time range" -ForegroundColor Red
        Write-Host "Date Range: $($StartDate.ToString("yyyy-MM-dd HH:mm:ss")) to $($EndDate.ToString("yyyy-MM-dd HH:mm:ss"))" -ForegroundColor Yellow
        
        # Create empty CSV with headers for consistency
        $emptyResult = [PSCustomObject]@{
            DateTime = ""
            UserEmail = ""
            EventType = ""
            IPAddress = ""
            UserAgent = ""
            DeviceInfo = ""
            DeviceId = ""
            DeviceModel = ""
            DeviceOS = ""
            DeviceStatus = ""
            DeviceIMEI = ""
            DeviceSerial = ""
            RawData = ""
        }
        
        @($emptyResult)[0..0] | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        (Get-Content $OutputPath | Select-Object -Skip 1) | Set-Content $OutputPath  # Remove the empty row
        
        Write-Host "Empty results file created: $OutputPath" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Script completed successfully!" -ForegroundColor Green
