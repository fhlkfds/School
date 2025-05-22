<#
.SYNOPSIS
    Automated laptop exchange processing script using CSV data and SnipeIT
.DESCRIPTION
    Reads CSV data, validates laptop assignments, and processes exchanges
.AUTHOR
    PowerShell Automation Engineer
.VERSION
    1.3
#>

# Configuration
$Config = @{
    # CSV Configuration
    CSVFilePath = "laptop_swaps.csv"
    
    # Email Configuration (Using App Password for Gmail)
    EmailRecipient = "ldecareaux@nomma.net"
    EmailSender = "your-gmail@gmail.com"  # CHANGE THIS
    EmailAppPassword = "your-16-char-app-password"  # CHANGE THIS - Generate from Google Account settings
    SMTPServer = "smtp.gmail.com"
    SMTPPort = 587
    
    # Logging
    LogFile = "laptop_exchange_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    
    # SnipeIT Status IDs - CONFIGURE THESE based on your instance
    MaintenanceStatusId = 4
    MaintenanceSupplierId = 1
    DamagedStatusId = 3
    DeployedStatusId = 2
    ReadyToDeployStatusId = 2
}

# Initialize logging
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to console with color coding
    switch ($Level) {
        "INFO" { Write-Host $logEntry -ForegroundColor White }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
    }
    
    # Append to log file
    Add-Content -Path $Config.LogFile -Value $logEntry
}

# Connect to SnipeIT
function Connect-SnipeIT {
    try {
        Write-Log "Connecting to SnipeIT using stored credentials" "INFO"
        
        # Use the stored credential file
        Connect-SnipeitPS -siteCred (Import-CliXml snipecred.xml)
        
        # Test connection by getting a status label
        $testConnection = Get-SnipeitStatus -limit 1
        if ($testConnection) {
            Write-Log "Successfully connected to SnipeIT" "SUCCESS"
            return $true
        }
        else {
            Write-Log "Failed to verify SnipeIT connection" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Failed to connect to SnipeIT: $($_.Exception.Message)" "ERROR"
        Write-Log "Make sure 'snipecred.xml' file exists in the current directory" "ERROR"
        return $false
    }
}

# Email function using modern approach
function Send-AlertEmail {
    param(
        [string]$Subject,
        [string]$Body,
        [string]$Priority = "Normal"
    )
    
    try {
        # Create credential object for Gmail authentication
        $securePassword = ConvertTo-SecureString $Config.EmailAppPassword -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($Config.EmailSender, $securePassword)
        
        # Use MailKit/MimeKit approach if available, otherwise use Send-MailMessage with credential
        if (Get-Module -ListAvailable -Name "Mailozaurr") {
            # If Mailozaurr module is available (recommended)
            Import-Module Mailozaurr
            
            $emailParams = @{
                From = $Config.EmailSender
                To = $Config.EmailRecipient
                Subject = $Subject
                Text = $Body
                Server = $Config.SMTPServer
                Port = $Config.SMTPPort
                Credential = $credential
                SecureSocketOptions = "StartTls"
                Priority = $Priority
            }
            
            Send-EmailMessage @emailParams
            Write-Log "Email sent successfully using Mailozaurr: $Subject" "SUCCESS"
            return $true
        }
        else {
            # Fallback to Send-MailMessage (will show obsolete warning)
            $emailParams = @{
                To = $Config.EmailRecipient
                From = $Config.EmailSender
                Subject = $Subject
                Body = $Body
                SmtpServer = $Config.SMTPServer
                Port = $Config.SMTPPort
                UseSsl = $true
                Credential = $credential
                Priority = $Priority
            }
            
            Send-MailMessage @emailParams -WarningAction SilentlyContinue
            Write-Log "Email sent successfully: $Subject" "SUCCESS"
            return $true
        }
    }
    catch {
        Write-Log "Failed to send email: $($_.Exception.Message)" "ERROR"
        Write-Log "Make sure you're using a Gmail App Password, not your regular password" "WARNING"
        Write-Log "Generate an App Password at: https://myaccount.google.com/apppasswords" "WARNING"
        return $false
    }
}

# Alternative email function using .NET classes
function Send-AlertEmailDotNet {
    param(
        [string]$Subject,
        [string]$Body,
        [string]$Priority = "Normal"
    )
    
    try {
        Add-Type -AssemblyName System.Net.Mail
        
        $smtp = New-Object System.Net.Mail.SmtpClient($Config.SMTPServer, $Config.SMTPPort)
        $smtp.EnableSsl = $true
        $smtp.Credentials = New-Object System.Net.NetworkCredential($Config.EmailSender, $Config.EmailAppPassword)
        
        $message = New-Object System.Net.Mail.MailMessage
        $message.From = $Config.EmailSender
        $message.To.Add($Config.EmailRecipient)
        $message.Subject = $Subject
        $message.Body = $Body
        
        switch ($Priority) {
            "High" { $message.Priority = [System.Net.Mail.MailPriority]::High }
            "Low" { $message.Priority = [System.Net.Mail.MailPriority]::Low }
            default { $message.Priority = [System.Net.Mail.MailPriority]::Normal }
        }
        
        $smtp.Send($message)
        
        Write-Log "Email sent successfully using .NET: $Subject" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to send email using .NET: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Get available status labels to help configure maintenance status ID
function Get-SnipeitStatusLabels {
    try {
        Write-Log "Retrieving available status labels from SnipeIT" "INFO"
        
        # Try different methods to get status labels
        $statusLabels = $null
        
        # Method 1: Get-SnipeitStatus (most common)
        if (Get-Command "Get-SnipeitStatus" -ErrorAction SilentlyContinue) {
            $statusLabels = Get-SnipeitStatus
        }
        # Method 2: Get-SnipeitStatusLabel
        elseif (Get-Command "Get-SnipeitStatusLabel" -ErrorAction SilentlyContinue) {
            $statusLabels = Get-SnipeitStatusLabel
        }
        
        if ($statusLabels) {
            Write-Log "Available Status Labels:" "INFO"
            foreach ($status in $statusLabels) {
                $statusInfo = "  ID: $($status.id) - Name: $($status.name)"
                if ($status.status_type) {
                    $statusInfo += " - Type: $($status.status_type)"
                }
                Write-Log $statusInfo "INFO"
            }
            return $statusLabels
        }
        else {
            Write-Log "Could not retrieve status labels - check your SnipeIT PS module version" "WARNING"
            return $null
        }
    }
    catch {
        Write-Log "Exception retrieving status labels: $($_.Exception.Message)" "WARNING"
        return $null
    }
}

# Get CSV data
function Get-CSVData {
    Write-Log "Reading CSV data from: $($Config.CSVFilePath)" "INFO"
    
    try {
        # Check if CSV file exists
        if (-not (Test-Path $Config.CSVFilePath)) {
            Write-Log "CSV file not found: $($Config.CSVFilePath)" "ERROR"
            return $null
        }
        
        # Read CSV with proper header handling
        $csvData = Import-Csv -Path $Config.CSVFilePath -Encoding UTF8
        
        if ($csvData.Count -eq 0) {
            Write-Log "CSV file is empty or has no data rows" "WARNING"
            return $null
        }
        
        # Validate required headers exist
        $requiredHeaders = @(
            "What is your student ID number? (All 7 digits)",
            "How did you break your laptop?",
            "What is the Asset Tag (The numbers on the red sticker) on the broken laptop?",
            "What is the Asset Tag (The numbers on the red sticker) on the new laptop?"
        )
        
        $csvHeaders = $csvData[0].PSObject.Properties.Name
        $missingHeaders = @()
        
        foreach ($header in $requiredHeaders) {
            if ($header -notin $csvHeaders) {
                $missingHeaders += $header
            }
        }
        
        if ($missingHeaders.Count -gt 0) {
            Write-Log "Missing required CSV headers:" "ERROR"
            foreach ($missing in $missingHeaders) {
                Write-Log "  - $missing" "ERROR"
            }
            Write-Log "Available headers:" "INFO"
            foreach ($available in $csvHeaders) {
                Write-Log "  - $available" "INFO"
            }
            return $null
        }
        
        Write-Log "Successfully loaded $($csvData.Count) records from CSV" "SUCCESS"
        return $csvData
    }
    catch {
        Write-Log "Failed to read CSV data: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

# Validate student and laptop assignment
function Test-StudentLaptopAssignment {
    param(
        [string]$StudentID,
        [string]$AssetTag
    )
    
    try {
        Write-Log "Validating student $StudentID with laptop $AssetTag" "INFO"
        
        # Get student by ID
        $student = Get-SnipeitUser -search $StudentID
        if (-not $student) {
            Write-Log "Student not found: $StudentID" "WARNING"
            return @{ Valid = $false; Error = "Student not found" }
        }
        
        # Get asset by tag
        $asset = Get-SnipeitAsset -asset_tag $AssetTag
        if (-not $asset) {
            Write-Log "Asset not found: $AssetTag" "WARNING"
            return @{ Valid = $false; Error = "Asset not found" }
        }
        
        # Check if asset is assigned to the student
        if ($asset.assigned_to.id -eq $student.id) {
            Write-Log "Validation successful: Asset $AssetTag is assigned to student $StudentID" "SUCCESS"
            return @{ Valid = $true; Student = $student; Asset = $asset }
        }
        else {
            $assignedTo = if ($asset.assigned_to) { $asset.assigned_to.name } else { "Unassigned" }
            Write-Log "Validation failed: Asset $AssetTag is assigned to '$assignedTo', not student $StudentID" "WARNING"
            return @{ Valid = $false; Error = "Asset assignment mismatch"; AssignedTo = $assignedTo }
        }
    }
    catch {
        Write-Log "Exception during validation: $($_.Exception.Message)" "ERROR"
        return @{ Valid = $false; Error = $_.Exception.Message }
    }
}

# Set asset to damaged status and create maintenance record
function Set-AssetMaintenance {
    param(
        [string]$AssetTag,
        [string]$Reason = "Laptop Exchange Process",
        [string]$AssetType = "broken"  # "broken" or "replacement"
    )
    
    try {
        Write-Log "Processing maintenance for asset $AssetTag (Type: $AssetType)" "INFO"
        
        $asset = Get-SnipeitAsset -asset_tag $AssetTag
        if (-not $asset) {
            Write-Log "Asset not found: $AssetTag" "ERROR"
            return $false
        }
        
        # Determine the appropriate status based on asset type
        $targetStatusId = if ($AssetType -eq "broken") { $Config.DamagedStatusId } else { $Config.ReadyToDeployStatusId }
        $statusName = if ($AssetType -eq "broken") { "Damaged" } else { "Ready for Assignment" }
        
        # Validate status ID is greater than 0
        if ($targetStatusId -le 0) {
            Write-Log "Invalid status ID ($targetStatusId) for $AssetType asset. Please configure proper status IDs in the Config section." "ERROR"
            Write-Log "Current Config values - DamagedStatusId: $($Config.DamagedStatusId), ReadyToDeployStatusId: $($Config.ReadyToDeployStatusId)" "ERROR"
            return $false
        }
        
        # Method 1: Create maintenance record for broken assets
        if ($AssetType -eq "broken") {
            try {
                if (Get-Command "New-SnipeitAssetMaintenance" -ErrorAction SilentlyContinue) {
                    $maintenanceParams = @{
                        asset_id = $asset.id
                        supplier_id = $Config.MaintenanceSupplierId
                        asset_maintenance_type = "repair"
                        title = "Laptop Exchange - Damaged Device"
                        notes = $Reason
                        cost = 0
                        start_date = (Get-Date).ToString("yyyy-MM-dd")
                    }
                    
                    $maintenanceResult = New-SnipeitAssetMaintenance @maintenanceParams
                    
                    if ($maintenanceResult) {
                        Write-Log "Maintenance record created successfully for damaged asset $AssetTag" "SUCCESS"
                    }
                    else {
                        Write-Log "Failed to create maintenance record for $AssetTag" "WARNING"
                    }
                }
                else {
                    Write-Log "New-SnipeitAssetMaintenance command not available" "WARNING"
                }
            }
            catch {
                Write-Log "Failed to create maintenance record: $($_.Exception.Message)" "WARNING"
            }
        }
        
        # Method 2: Update asset status
        try {
            Write-Log "Attempting to update asset $AssetTag to status ID $targetStatusId ($statusName)" "INFO"
            
            $updateParams = @{
                id = $asset.id
                notes = $Reason
                status_id = $targetStatusId
            }
            $updateResult = Set-SnipeitAsset @updateParams
            
            if ($updateResult) {
                Write-Log "Asset $AssetTag status updated to '$statusName' (ID: $targetStatusId) successfully" "SUCCESS"
                return $true
            }
            else {
                Write-Log "Failed to update asset status for $AssetTag" "ERROR"
            }
        }
        catch {
            Write-Log "Failed to update asset status: $($_.Exception.Message)" "WARNING"
        }
        
        # Method 3: Just update notes if status update fails
        try {
            $noteUpdateResult = Set-SnipeitAsset -id $asset.id -notes "$Reason - STATUS: $statusName"
            if ($noteUpdateResult) {
                Write-Log "Asset $AssetTag notes updated (status not changed)" "WARNING"
                return $true
            }
        }
        catch {
            Write-Log "Failed to update asset notes: $($_.Exception.Message)" "ERROR"
        }
        
        Write-Log "Failed to process asset $AssetTag using all methods" "ERROR"
        return $false
    }
    catch {
        Write-Log "Exception processing asset $AssetTag : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Update asset notes
function Update-AssetNotes {
    param(
        [string]$AssetTag,
        [string]$Notes
    )
    
    try {
        Write-Log "Updating notes for asset $AssetTag" "INFO"
        
        $asset = Get-SnipeitAsset -asset_tag $AssetTag
        if (-not $asset) {
            Write-Log "Asset not found for note update: $AssetTag" "ERROR"
            return $false
        }
        
        $updateResult = Set-SnipeitAsset -id $asset.id -notes $Notes
        
        if ($updateResult) {
            Write-Log "Notes updated successfully for asset $AssetTag" "SUCCESS"
            return $true
        }
        else {
            Write-Log "Failed to update notes for asset $AssetTag" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Exception updating notes for $AssetTag : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Process individual laptop exchange
function Process-LaptopExchange {
    param(
        [PSCustomObject]$ExchangeData
    )
    
    $studentID = $ExchangeData."What is your student ID number? (All 7 digits)"
    $breakReason = $ExchangeData."How did you break your laptop?"
    $brokenAssetTag = $ExchangeData."What is the Asset Tag (The numbers on the red sticker) on the broken laptop?"
    $newAssetTag = $ExchangeData."What is the Asset Tag (The numbers on the red sticker) on the new laptop?"
    
    # Validate required fields are not empty
    if ([string]::IsNullOrWhiteSpace($studentID) -or 
        [string]::IsNullOrWhiteSpace($brokenAssetTag) -or 
        [string]::IsNullOrWhiteSpace($newAssetTag)) {
        Write-Log "Skipping row with missing required data" "WARNING"
        return $false
    }
    
    Write-Log "Processing exchange for Student ID: $studentID" "INFO"
    Write-Log "Broken laptop: $brokenAssetTag, New laptop: $newAssetTag" "INFO"
    
    # Validate broken laptop assignment
    $validation = Test-StudentLaptopAssignment -StudentID $studentID -AssetTag $brokenAssetTag
    
    if (-not $validation.Valid) {
        Write-Log "Validation failed for student $studentID and broken laptop $brokenAssetTag" "ERROR"
        
        $alertSubject = "ALERT: Laptop Assignment Conflict"
        $alertBody = @"
Student ID: $studentID
Broken Laptop Asset Tag: $brokenAssetTag
Issue: $($validation.Error)
$(if ($validation.AssignedTo) { "Currently assigned to: $($validation.AssignedTo)" })

Please review this laptop exchange request manually.
"@
        Send-AlertEmail -Subject $alertSubject -Body $alertBody -Priority "High"
        return $false
    }
    
    $student = $validation.Student
    
    # Check if new laptop is already assigned
    try {
        $newAsset = Get-SnipeitAsset -asset_tag $newAssetTag
        if (-not $newAsset) {
            Write-Log "New asset not found: $newAssetTag" "ERROR"
            return $false
        }
        
        if ($newAsset.assigned_to -and $newAsset.assigned_to.id -ne $student.id) {
            Write-Log "New laptop $newAssetTag is already assigned to someone else" "WARNING"
            
            $alertSubject = "ALERT: New Laptop Already Assigned"
            $alertBody = @"
Student ID: $studentID ($($student.name))
New Laptop Asset Tag: $newAssetTag
Currently assigned to: $($newAsset.assigned_to.name)
Should be assigned to: $($student.name) (ID: $studentID)

Please resolve this assignment conflict before proceeding.
"@
            Send-AlertEmail -Subject $alertSubject -Body $alertBody -Priority "High"
            return $false
        }
    }
    catch {
        Write-Log "Exception checking new asset assignment: $($_.Exception.Message)" "ERROR"
        return $false
    }
    
    # Process the exchange
    try {
        # Set broken laptop to damaged status and create maintenance record
        if (-not (Set-AssetMaintenance -AssetTag $brokenAssetTag -Reason "Broken - $breakReason" -AssetType "broken")) {
            Write-Log "Failed to set broken laptop to damaged status" "ERROR"
            return $false
        }
        
        # Prepare new laptop (ensure it's ready for deployment)
        if (-not (Set-AssetMaintenance -AssetTag $newAssetTag -Reason "Prepared for assignment via laptop exchange" -AssetType "replacement")) {
            Write-Log "Failed to prepare new laptop for assignment" "ERROR"
            return $false
        }
        
        # Update notes on broken laptop
        $noteText = "BROKEN: $breakReason - Processed on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        if (-not (Update-AssetNotes -AssetTag $brokenAssetTag -Notes $noteText)) {
            Write-Log "Failed to update notes on broken laptop" "WARNING"
        }
        
        # Unassign broken laptop using Reset-SnipeitAssetOwner
        $brokenAsset = Get-SnipeitAsset -asset_tag $brokenAssetTag
        # Reset owner removes assignment and optionally sets status
        $unassignResult = Reset-SnipeitAssetOwner -id $brokenAsset.id -status_id $Config.DamagedStatusId -note "Unassigned due to damage: $breakReason"
        if (-not $unassignResult) {
            Write-Log "Failed to unassign broken laptop" "ERROR"
            return $false
        }
        Write-Log "Successfully unassigned broken laptop $brokenAssetTag" "SUCCESS"
        
        # Assign new laptop
        $assignResult = Set-SnipeitAssetOwner -id $newAsset.id -assigned_id $student.id
        if (-not $assignResult) {
            Write-Log "Failed to assign new laptop" "ERROR"
            return $false
        }
        
        Write-Log "Exchange completed successfully for student $studentID" "SUCCESS"
        
        # Send success email
        $successSubject = "SUCCESS: Laptop Exchange Completed"
        $successBody = @"
Laptop exchange completed successfully:

Student: $($student.name) (ID: $studentID)
Broken Laptop: $brokenAssetTag (Unassigned)
New Laptop: $newAssetTag (Assigned to $($student.name))
Break Reason: $breakReason
Processed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
        Send-AlertEmail -Subject $successSubject -Body $successBody
        
        return $true
    }
    catch {
        Write-Log "Exception during exchange process: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Main execution
function Start-LaptopExchangeProcess {
    Write-Log "=== Starting Laptop Exchange Process ===" "INFO"
    
    # Connect to SnipeIT first
    if (-not (Connect-SnipeIT)) {
        Write-Log "Failed to connect to SnipeIT. Please check your snipecred.xml file." "ERROR"
        Write-Log "Make sure the credential file exists and contains valid credentials." "ERROR"
        return
    }
    
    # Validate configuration
    Write-Log "--- Validating Configuration ---" "INFO"
    Write-Log "DamagedStatusId: $($Config.DamagedStatusId)" "INFO"
    Write-Log "ReadyToDeployStatusId: $($Config.ReadyToDeployStatusId)" "INFO"
    
    if ($Config.DamagedStatusId -le 0 -or $Config.ReadyToDeployStatusId -le 0) {
        Write-Log "ERROR: Invalid status ID configuration. Please set proper status IDs in the Config section." "ERROR"
        Write-Log "DamagedStatusId and ReadyToDeployStatusId must be greater than 0." "ERROR"
        return
    }
    
    # First, get available status labels to help with configuration
    Write-Log "--- Checking Available Status Labels ---" "INFO"
    $statusLabels = Get-SnipeitStatusLabels
    
    try {
        # Get CSV data
        $exchangeData = Get-CSVData
        if (-not $exchangeData) {
            Write-Log "Failed to load CSV data. Exiting." "ERROR"
            return
        }
        
        # Process each exchange request
        $successCount = 0
        $failCount = 0
        $rowNumber = 0
        
        foreach ($row in $exchangeData) {
            $rowNumber++
            Write-Log "--- Processing Row $rowNumber ---" "INFO"
            
            if (Process-LaptopExchange -ExchangeData $row) {
                $successCount++
            }
            else {
                $failCount++
            }
            
            # Add small delay between processing
            Start-Sleep -Seconds 2
        }
        
        Write-Log "=== Process Complete ===" "INFO"
        Write-Log "Successful exchanges: $successCount" "SUCCESS"
        Write-Log "Failed exchanges: $failCount" "INFO"
        
        # Send summary email
        $summarySubject = "Laptop Exchange Process Summary"
        $summaryBody = @"
Laptop exchange batch process completed:

Total requests processed: $($exchangeData.Count)
Successful exchanges: $successCount
Failed exchanges: $failCount

Log file: $($Config.LogFile)
Processed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
        Send-AlertEmail -Subject $summarySubject -Body $summaryBody
        
    }
    catch {
        Write-Log "Critical exception in main process: $($_.Exception.Message)" "ERROR"
        
        $errorSubject = "CRITICAL: Laptop Exchange Process Failed"
        $errorBody = @"
The laptop exchange process encountered a critical error:

Error: $($_.Exception.Message)
Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Log file: $($Config.LogFile)

Please review the logs and restart the process manually.
"@
        Send-AlertEmail -Subject $errorSubject -Body $errorBody -Priority "High"
    }
}

# Execute the main function
Start-LaptopExchangeProcess
