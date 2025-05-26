<#
.SYNOPSIS
    Automated laptop exchange processing script using Google Sheets data and SnipeIT
.DESCRIPTION
    Downloads CSV data from Google Sheets using rclone, validates laptop assignments, and processes exchanges
.AUTHOR
    PowerShell Automation Engineer
.VERSION
    1.5 - Added Google Sheets download capability with rclone
#>

# Configuration
$Config = @{
  # Google Sheets Configuration
  RcloneRemoteName = "gdrive"  # CHANGE THIS - Your rclone remote name for Google Drive
  GoogleSheetFileId = ""       # CHANGE THIS - The Google Sheets file ID from the URL
  GoogleSheetName = "Sheet1"   # CHANGE THIS - The sheet name/tab to download
  LocalCSVFileName = "laptop_swaps.csv"  # Local filename for downloaded CSV
  RcloneConfigPath = ""        # OPTIONAL - Path to rclone config file (leave empty for default)
    
  # CSV Configuration (kept for backward compatibility)
  CSVFilePath = "laptop_swaps.csv"
    
  # Email Configuration (Using App Password for Gmail)
  EmailRecipient = ""
  EmailSender = ""  # CHANGE THIS
  EmailAppPassword = ""  # CHANGE THIS - Generate from Google Account settings
  SMTPServer = "smtp.gmail.com"
  SMTPPort = 587
    
  # Logging
  LogFile = "laptop_exchange_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
  LogRetentionDays = 30  # Keep logs for 30 days
    
  # SnipeIT Status IDs - CONFIGURE THESE based on your instance
  MaintenanceStatusId = 4
  MaintenanceSupplierId = 1
  DamagedStatusId = 6
  DeployedStatusId = 5
  ReadyToDeployStatusId = 2
}

# Test rclone installation and configuration
function Test-RcloneSetup
{
  try
  {
    Write-Log "Testing rclone installation and configuration" "INFO"
        
    # Check if rclone is installed and accessible
    $rcloneVersion = & rclone version 2>$null
    if ($LASTEXITCODE -ne 0)
    {
      Write-Log "rclone is not installed or not in PATH" "ERROR"
      Write-Log "Please install rclone from https://rclone.org/downloads/" "ERROR"
      return $false
    }
        
    Write-Log "rclone found: $($rcloneVersion[0])" "SUCCESS"
        
    # Check if the specified remote exists
    $remotes = & rclone listremotes 2>$null
    if ($LASTEXITCODE -ne 0)
    {
      Write-Log "Failed to list rclone remotes" "ERROR"
      return $false
    }
        
    $remoteExists = $remotes -contains "$($Config.RcloneRemoteName):"
    if (-not $remoteExists)
    {
      Write-Log "rclone remote '$($Config.RcloneRemoteName)' not found" "ERROR"
      Write-Log "Available remotes: $($remotes -join ', ')" "ERROR"
      Write-Log "Please configure your Google Drive remote using: rclone config" "ERROR"
      return $false
    }
        
    Write-Log "rclone remote '$($Config.RcloneRemoteName)' found" "SUCCESS"
        
    # Test remote connectivity
    Write-Log "Testing remote connectivity..." "INFO"
    $testResult = & rclone lsd "$($Config.RcloneRemoteName):" --max-depth 1 2>$null
    if ($LASTEXITCODE -ne 0)
    {
      Write-Log "Failed to connect to remote '$($Config.RcloneRemoteName)'" "ERROR"
      Write-Log "Please check your rclone configuration and internet connection" "ERROR"
      return $false
    }
        
    Write-Log "Remote connectivity test successful" "SUCCESS"
    return $true
  } catch
  {
    Write-Log "Exception testing rclone setup: $($_.Exception.Message)" "ERROR"
    return $false
  }
}

# Download Google Sheets as CSV using rclone
function Get-GoogleSheetData
{
  try
  {
    Write-Log "Starting Google Sheets download process" "INFO"
        
    # Validate required configuration
    if ([string]::IsNullOrWhiteSpace($Config.GoogleSheetFileId))
    {
      Write-Log "GoogleSheetFileId is not configured" "ERROR"
      Write-Log "Please set the Google Sheets file ID in the Config section" "ERROR"
      Write-Log "You can find the file ID in the Google Sheets URL: https://docs.google.com/spreadsheets/d/FILE_ID_HERE/edit" "ERROR"
      return $false
    }
        
    # Test rclone setup first
    if (-not (Test-RcloneSetup))
    {
      Write-Log "rclone setup test failed" "ERROR"
      return $false
    }
        
    # Construct the Google Sheets export URL for CSV format
    # Google Sheets can be exported as CSV using a specific URL format
    $exportUrl = "https://docs.google.com/spreadsheets/d/$($Config.GoogleSheetFileId)/export?format=csv&gid=0"
        
    # If a specific sheet name is provided, we need to get the GID (sheet ID)
    if (-not [string]::IsNullOrWhiteSpace($Config.GoogleSheetName) -and $Config.GoogleSheetName -ne "Sheet1")
    {
      Write-Log "Note: Specific sheet name specified. You may need to manually find the GID for sheet '$($Config.GoogleSheetName)'" "WARNING"
      Write-Log "To find the GID, open the sheet and look at the URL: ...#gid=SHEET_ID" "WARNING"
    }
        
    # Method 1: Try to copy the file directly and then export
    Write-Log "Attempting to download Google Sheet using rclone..." "INFO"
        
    # Construct rclone command arguments
    $rcloneArgs = @()
        
    # Add config file path if specified
    if (-not [string]::IsNullOrWhiteSpace($Config.RcloneConfigPath))
    {
      $rcloneArgs += "--config"
      $rcloneArgs += $Config.RcloneConfigPath
    }
        
    # Try different approaches to get the Google Sheets data
        
    # Approach 1: Direct file copy (if the sheet is shared as CSV)
    try
    {
      Write-Log "Trying direct file copy method..." "INFO"
      $sourceFile = "$($Config.RcloneRemoteName):$($Config.GoogleSheetFileId)"
            
      $copyArgs = $rcloneArgs + @("copy", $sourceFile, ".", "--include", "*.csv", "-v")
      $copyResult = & rclone @copyArgs 2>&1
            
      if ($LASTEXITCODE -eq 0 -and (Test-Path $Config.LocalCSVFileName))
      {
        Write-Log "Successfully downloaded CSV using direct copy method" "SUCCESS"
        return $true
      }
    } catch
    {
      Write-Log "Direct copy method failed: $($_.Exception.Message)" "WARNING"
    }
        
    # Approach 2: Use rclone with Google Sheets export URL
    try
    {
      Write-Log "Trying Google Sheets export URL method..." "INFO"
            
      # Create a temporary rclone remote for the HTTP download
      $tempRemoteName = "temp_sheets_download"
            
      # Download using rclone's HTTP capabilities or direct download
      Write-Log "Attempting direct download using PowerShell..." "INFO"
            
      # Use PowerShell to download the CSV directly
      $webClient = New-Object System.Net.WebClient
      $webClient.DownloadFile($exportUrl, $Config.LocalCSVFileName)
            
      if (Test-Path $Config.LocalCSVFileName)
      {
        $fileSize = (Get-Item $Config.LocalCSVFileName).Length
        if ($fileSize -gt 0)
        {
          Write-Log "Successfully downloaded CSV using direct HTTP method ($fileSize bytes)" "SUCCESS"
          return $true
        } else
        {
          Write-Log "Downloaded file is empty" "WARNING"
        }
      }
    } catch
    {
      Write-Log "HTTP download method failed: $($_.Exception.Message)" "WARNING"
    }
        
    # Approach 3: Manual rclone copy with file ID
    try
    {
      Write-Log "Trying rclone copy with file ID..." "INFO"
            
      $copyArgs = $rcloneArgs + @("copy", "$($Config.RcloneRemoteName):$($Config.GoogleSheetFileId)", ".", "-v")
      $copyOutput = & rclone @copyArgs 2>&1
      Write-Log "rclone copy output: $copyOutput" "INFO"
            
      # Look for any CSV files that were downloaded
      $csvFiles = Get-ChildItem -Path "." -Filter "*.csv" | Sort-Object LastWriteTime -Descending
            
      if ($csvFiles.Count -gt 0)
      {
        $latestCsv = $csvFiles[0]
        if ($latestCsv.Name -ne $Config.LocalCSVFileName)
        {
          Copy-Item -Path $latestCsv.FullName -Destination $Config.LocalCSVFileName -Force
          Write-Log "Renamed downloaded file to $($Config.LocalCSVFileName)" "INFO"
        }
        Write-Log "Successfully downloaded CSV using rclone copy method" "SUCCESS"
        return $true
      }
    } catch
    {
      Write-Log "rclone copy method failed: $($_.Exception.Message)" "WARNING"
    }
        
    Write-Log "All download methods failed" "ERROR"
    Write-Log "Please check your Google Sheets configuration and ensure the file is accessible" "ERROR"
    return $false
  } catch
  {
    Write-Log "Exception during Google Sheets download: $($_.Exception.Message)" "ERROR"
    return $false
  }
}

# Backup existing CSV file
function Backup-ExistingCSV
{
  try
  {
    if (Test-Path $Config.CSVFilePath)
    {
      $backupName = "laptop_swaps_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
      Copy-Item -Path $Config.CSVFilePath -Destination $backupName
      Write-Log "Existing CSV backed up as: $backupName" "SUCCESS"
      return $true
    }
    return $true
  } catch
  {
    Write-Log "Failed to backup existing CSV: $($_.Exception.Message)" "WARNING"
    return $false
  }
}

# Clean up old log files
function Remove-OldLogFiles
{
  try
  {
    Write-Host "Cleaning up old log files..." -ForegroundColor Yellow
        
    $logPattern = "laptop_exchange_*.log"
    $cutoffDate = (Get-Date).AddDays(-$Config.LogRetentionDays)
        
    $oldLogFiles = Get-ChildItem -Path "." -Filter $logPattern | Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
    if ($oldLogFiles)
    {
      foreach ($logFile in $oldLogFiles)
      {
        try
        {
          Remove-Item -Path $logFile.FullName -Force
          Write-Host "Deleted old log file: $($logFile.Name)" -ForegroundColor Green
        } catch
        {
          Write-Host "Failed to delete log file $($logFile.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
      }
      Write-Host "Cleaned up $($oldLogFiles.Count) old log file(s)" -ForegroundColor Green
    } else
    {
      Write-Host "No old log files found to clean up" -ForegroundColor Green
    }
  } catch
  {
    Write-Host "Error during log cleanup: $($_.Exception.Message)" -ForegroundColor Red
  }
}

# Initialize logging
function Write-Log
{
  param(
    [string]$Message,
    [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
    [string]$Level = "INFO"
  )
    
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $logEntry = "[$timestamp] [$Level] $Message"
    
  # Write to console with color coding
  switch ($Level)
  {
    "INFO"
    { Write-Host $logEntry -ForegroundColor White 
    }
    "WARNING"
    { Write-Host $logEntry -ForegroundColor Yellow 
    }
    "ERROR"
    { Write-Host $logEntry -ForegroundColor Red 
    }
    "SUCCESS"
    { Write-Host $logEntry -ForegroundColor Green 
    }
  }
    
  # Append to log file
  Add-Content -Path $Config.LogFile -Value $logEntry
}

# Connect to SnipeIT
function Connect-SnipeIT
{
  try
  {
    Write-Log "Connecting to SnipeIT using stored credentials" "INFO"
        
    # Use the stored credential file
    Connect-SnipeitPS -siteCred (Import-CliXml snipecred.xml)
        
    # Test connection by getting a status label
    $testConnection = Get-SnipeitStatus -limit 1
    if ($testConnection)
    {
      Write-Log "Successfully connected to SnipeIT" "SUCCESS"
      return $true
    } else
    {
      Write-Log "Failed to verify SnipeIT connection" "ERROR"
      return $false
    }
  } catch
  {
    Write-Log "Failed to connect to SnipeIT: $($_.Exception.Message)" "ERROR"
    Write-Log "Make sure 'snipecred.xml' file exists in the current directory" "ERROR"
    return $false
  }
}

# Email function using modern approach
function Send-AlertEmail
{
  param(
    [string]$Subject,
    [string]$Body,
    [string]$Priority = "Normal"
  )
    
  try
  {
    # Create credential object for Gmail authentication
    $securePassword = ConvertTo-SecureString $Config.EmailAppPassword -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Config.EmailSender, $securePassword)
        
    # Use MailKit/MimeKit approach if available, otherwise use Send-MailMessage with credential
    if (Get-Module -ListAvailable -Name "Mailozaurr")
    {
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
    } else
    {
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
  } catch
  {
    Write-Log "Failed to send email: $($_.Exception.Message)" "ERROR"
    Write-Log "Make sure you're using a Gmail App Password, not your regular password" "WARNING"
    Write-Log "Generate an App Password at: https://myaccount.google.com/apppasswords" "WARNING"
    return $false
  }
}

# Enhanced function to send detailed conflict emails
function Send-ConflictEmail
{
  param(
    [string]$StudentID,
    [string]$StudentName = "Unknown",
    [string]$BrokenAssetTag,
    [string]$NewAssetTag,
    [string]$BreakReason,
    [string]$ConflictType,
    [string]$ConflictDetails,
    [string]$AssignedTo = ""
  )
    
  $processedTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    
  $alertSubject = "ALERT: Laptop Exchange Conflict - $ConflictType"
  $alertBody = @"
🚨 LAPTOP EXCHANGE CONFLICT DETECTED 🚨

STUDENT INFORMATION:
• Student ID: $StudentID
• Student Name: $StudentName

LAPTOP INFORMATION:
• Broken Laptop Asset Tag: $BrokenAssetTag
• New Laptop Asset Tag: $NewAssetTag
• Break Reason: $BreakReason

CONFLICT DETAILS:
• Conflict Type: $ConflictType
• Issue: $ConflictDetails
$(if ($AssignedTo) { "• Currently Assigned To: $AssignedTo" })

PROCESSING INFORMATION:
• Processed: $processedTime
• Log File: $($Config.LogFile)

⚠️ ACTION REQUIRED: Please review this laptop exchange request manually and resolve the conflict before proceeding.

This exchange has been skipped and will need to be processed manually once the conflict is resolved.
"@
    
  Send-AlertEmail -Subject $alertSubject -Body $alertBody -Priority "High"
  Write-Log "Conflict email sent for Student ID: $StudentID" "INFO"
}

# Get available status labels to help configure maintenance status ID
function Get-SnipeitStatusLabels
{
  try
  {
    Write-Log "Retrieving available status labels from SnipeIT" "INFO"
        
    # Try different methods to get status labels
    $statusLabels = $null
        
    # Method 1: Get-SnipeitStatus (most common)
    if (Get-Command "Get-SnipeitStatus" -ErrorAction SilentlyContinue)
    {
      $statusLabels = Get-SnipeitStatus
    }
    # Method 2: Get-SnipeitStatusLabel
    elseif (Get-Command "Get-SnipeitStatusLabel" -ErrorAction SilentlyContinue)
    {
      $statusLabels = Get-SnipeitStatusLabel
    }
        
    if ($statusLabels)
    {
      Write-Log "Available Status Labels:" "INFO"
      foreach ($status in $statusLabels)
      {
        $statusInfo = "  ID: $($status.id) - Name: $($status.name)"
        if ($status.status_type)
        {
          $statusInfo += " - Type: $($status.status_type)"
        }
        Write-Log $statusInfo "INFO"
      }
      return $statusLabels
    } else
    {
      Write-Log "Could not retrieve status labels - check your SnipeIT PS module version" "WARNING"
      return $null
    }
  } catch
  {
    Write-Log "Exception retrieving status labels: $($_.Exception.Message)" "WARNING"
    return $null
  }
}

# Get CSV data (now expects downloaded file)
function Get-CSVData
{
  Write-Log "Reading CSV data from: $($Config.CSVFilePath)" "INFO"
    
  try
  {
    # Check if CSV file exists
    if (-not (Test-Path $Config.CSVFilePath))
    {
      Write-Log "CSV file not found: $($Config.CSVFilePath)" "ERROR"
      return $null
    }
        
    # Read CSV with proper header handling
    $csvData = Import-Csv -Path $Config.CSVFilePath -Encoding UTF8
        
    if ($csvData.Count -eq 0)
    {
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
        
    foreach ($header in $requiredHeaders)
    {
      if ($header -notin $csvHeaders)
      {
        $missingHeaders += $header
      }
    }
        
    if ($missingHeaders.Count -gt 0)
    {
      Write-Log "Missing required CSV headers:" "ERROR"
      foreach ($missing in $missingHeaders)
      {
        Write-Log "  - $missing" "ERROR"
      }
      Write-Log "Available headers:" "INFO"
      foreach ($available in $csvHeaders)
      {
        Write-Log "  - $available" "INFO"
      }
      return $null
    }
        
    Write-Log "Successfully loaded $($csvData.Count) records from CSV" "SUCCESS"
    return $csvData
  } catch
  {
    Write-Log "Failed to read CSV data: $($_.Exception.Message)" "ERROR"
    return $null
  }
}

# Validate student and laptop assignment
function Test-StudentLaptopAssignment
{
  param(
    [string]$StudentID,
    [string]$AssetTag
  )
    
  try
  {
    Write-Log "Validating student $StudentID with laptop $AssetTag" "INFO"
        
    # Get student by ID
    $student = Get-SnipeitUser -search $StudentID
    if (-not $student)
    {
      Write-Log "Student not found: $StudentID" "WARNING"
      return @{ Valid = $false; Error = "Student not found"; Student = $null; Asset = $null }
    }
        
    # Get asset by tag
    $asset = Get-SnipeitAsset -asset_tag $AssetTag
    if (-not $asset)
    {
      Write-Log "Asset not found: $AssetTag" "WARNING"
      return @{ Valid = $false; Error = "Asset not found"; Student = $student; Asset = $null }
    }
        
    # Check if asset is assigned to the student
    if ($asset.assigned_to.id -eq $student.id)
    {
      Write-Log "Validation successful: Asset $AssetTag is assigned to student $StudentID" "SUCCESS"
      return @{ Valid = $true; Student = $student; Asset = $asset }
    } else
    {
      $assignedTo = if ($asset.assigned_to)
      { $asset.assigned_to.name 
      } else
      { "Unassigned" 
      }
      Write-Log "Validation failed: Asset $AssetTag is assigned to '$assignedTo', not student $StudentID" "WARNING"
      return @{ Valid = $false; Error = "Asset assignment mismatch"; AssignedTo = $assignedTo; Student = $student; Asset = $asset }
    }
  } catch
  {
    Write-Log "Exception during validation: $($_.Exception.Message)" "ERROR"
    return @{ Valid = $false; Error = $_.Exception.Message; Student = $null; Asset = $null }
  }
}

# Set asset to damaged status and create maintenance record
function Set-AssetMaintenance
{
  param(
    [string]$AssetTag,
    [string]$Reason = "Laptop Exchange Process",
    [string]$AssetType = "broken"  # "broken" or "replacement"
  )
    
  try
  {
    Write-Log "Processing maintenance for asset $AssetTag (Type: $AssetType)" "INFO"
        
    $asset = Get-SnipeitAsset -asset_tag $AssetTag
    if (-not $asset)
    {
      Write-Log "Asset not found: $AssetTag" "ERROR"
      return $false
    }
        
    # Determine the appropriate status based on asset type
    $targetStatusId = if ($AssetType -eq "broken")
    { $Config.DamagedStatusId 
    } else
    { $Config.ReadyToDeployStatusId 
    }
    $statusName = if ($AssetType -eq "broken")
    { "Damaged" 
    } else
    { "Ready for Assignment" 
    }
        
    # Validate status ID is greater than 0
    if ($targetStatusId -le 0)
    {
      Write-Log "Invalid status ID ($targetStatusId) for $AssetType asset. Please configure proper status IDs in the Config section." "ERROR"
      Write-Log "Current Config values - DamagedStatusId: $($Config.DamagedStatusId), ReadyToDeployStatusId: $($Config.ReadyToDeployStatusId)" "ERROR"
      return $false
    }
        
    # Method 1: Create maintenance record for broken assets
    if ($AssetType -eq "broken")
    {
      try
      {
        if (Get-Command "New-SnipeitAssetMaintenance" -ErrorAction SilentlyContinue)
        {
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
                    
          if ($maintenanceResult)
          {
            Write-Log "Maintenance record created successfully for damaged asset $AssetTag" "SUCCESS"
          } else
          {
            Write-Log "Failed to create maintenance record for $AssetTag" "WARNING"
          }
        } else
        {
          Write-Log "New-SnipeitAssetMaintenance command not available" "WARNING"
        }
      } catch
      {
        Write-Log "Failed to create maintenance record: $($_.Exception.Message)" "WARNING"
      }
    }
        
    # Method 2: Update asset status
    try
    {
      Write-Log "Attempting to update asset $AssetTag to status ID $targetStatusId ($statusName)" "INFO"
            
      $updateParams = @{
        id = $asset.id
        notes = $Reason
        status_id = $targetStatusId
      }
      $updateResult = Set-SnipeitAsset @updateParams
            
      if ($updateResult)
      {
        Write-Log "Asset $AssetTag status updated to '$statusName' (ID: $targetStatusId) successfully" "SUCCESS"
        return $true
      } else
      {
        Write-Log "Failed to update asset status for $AssetTag" "ERROR"
      }
    } catch
    {
      Write-Log "Failed to update asset status: $($_.Exception.Message)" "WARNING"
    }
        
    # Method 3: Just update notes if status update fails
    try
    {
      $noteUpdateResult = Set-SnipeitAsset -id $asset.id -notes "$Reason - STATUS: $statusName"
      if ($noteUpdateResult)
      {
        Write-Log "Asset $AssetTag notes updated (status not changed)" "WARNING"
        return $true
      }
    } catch
    {
      Write-Log "Failed to update asset notes: $($_.Exception.Message)" "ERROR"
    }
        
    Write-Log "Failed to process asset $AssetTag using all methods" "ERROR"
    return $false
  } catch
  {
    Write-Log "Exception processing asset $AssetTag : $($_.Exception.Message)" "ERROR"
    return $false
  }
}

# Update asset notes
function Update-AssetNotes
{
  param(
    [string]$AssetTag,
    [string]$Notes
  )
    
  try
  {
    Write-Log "Updating notes for asset $AssetTag" "INFO"
        
    $asset = Get-SnipeitAsset -asset_tag $AssetTag
    if (-not $asset)
    {
      Write-Log "Asset not found for note update: $AssetTag" "ERROR"
      return $false
    }
        
    $updateResult = Set-SnipeitAsset -id $asset.id -notes $Notes
        
    if ($updateResult)
    {
      Write-Log "Notes updated successfully for asset $AssetTag" "SUCCESS"
      return $true
    } else
    {
      Write-Log "Failed to update notes for asset $AssetTag" "ERROR"
      return $false
    }
  } catch
  {
    Write-Log "Exception updating notes for $AssetTag : $($_.Exception.Message)" "ERROR"
    return $false
  }
}

# Process individual laptop exchange
function Process-LaptopExchange
{
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
    [string]::IsNullOrWhiteSpace($newAssetTag))
  {
    Write-Log "Skipping row with missing required data" "WARNING"
    return $false
  }
    
  Write-Log "Processing exchange for Student ID: $studentID" "INFO"
  Write-Log "Broken laptop: $brokenAssetTag, New laptop: $newAssetTag" "INFO"
    
  # Validate broken laptop assignment
  $validation = Test-StudentLaptopAssignment -StudentID $studentID -AssetTag $brokenAssetTag
    
  if (-not $validation.Valid)
  {
    Write-Log "Validation failed for student $studentID and broken laptop $brokenAssetTag" "ERROR"
        
    # Send detailed conflict email
    $studentName = if ($validation.Student)
    { $validation.Student.name 
    } else
    { "Unknown" 
    }
    $assignedTo = if ($validation.AssignedTo)
    { $validation.AssignedTo 
    } else
    { "" 
    }
        
    Send-ConflictEmail -StudentID $studentID -StudentName $studentName -BrokenAssetTag $brokenAssetTag -NewAssetTag $newAssetTag -BreakReason $breakReason -ConflictType "Broken Laptop Assignment Mismatch" -ConflictDetails $validation.Error -AssignedTo $assignedTo
        
    return $false
  }
    
  $student = $validation.Student
    
  # Check if new laptop is already assigned
  try
  {
    $newAsset = Get-SnipeitAsset -asset_tag $newAssetTag
    if (-not $newAsset)
    {
      Write-Log "New asset not found: $newAssetTag" "ERROR"
            
      # Send conflict email for missing new asset
      Send-ConflictEmail -StudentID $studentID -StudentName $student.name -BrokenAssetTag $brokenAssetTag -NewAssetTag $newAssetTag -BreakReason $breakReason -ConflictType "New Laptop Not Found" -ConflictDetails "The new laptop asset tag does not exist in SnipeIT"
            
      return $false
    }
        
    if ($newAsset.assigned_to -and $newAsset.assigned_to.id -ne $student.id)
    {
      Write-Log "New laptop $newAssetTag is already assigned to someone else" "WARNING"
            
      # Send conflict email for already assigned new laptop
      Send-ConflictEmail -StudentID $studentID -StudentName $student.name -BrokenAssetTag $brokenAssetTag -NewAssetTag $newAssetTag -BreakReason $breakReason -ConflictType "New Laptop Already Assigned" -ConflictDetails "The new laptop is already assigned to another user" -AssignedTo $newAsset.assigned_to.name
            
      return $false
    }
  } catch
  {
    Write-Log "Exception checking new asset assignment: $($_.Exception.Message)" "ERROR"
        
    # Send conflict email for exception
    Send-ConflictEmail -StudentID $studentID -StudentName $student.name -BrokenAssetTag $brokenAssetTag -NewAssetTag $newAssetTag -BreakReason $breakReason -ConflictType "System Error" -ConflictDetails "Exception occurred while checking new laptop assignment: $($_.Exception.Message)"
        
    return $false
  }
    
  # Process the exchange
  try
  {
    # Set broken laptop to damaged status and create maintenance record
    if (-not (Set-AssetMaintenance -AssetTag $brokenAssetTag -Reason "Broken - $breakReason" -AssetType "broken"))
    {
      Write-Log "Failed to set broken laptop to damaged status" "ERROR"
      return $false
    }
        
    # Prepare new laptop (ensure it's ready for deployment)
    if (-not (Set-AssetMaintenance -AssetTag $newAssetTag -Reason "Prepared for assignment via laptop exchange" -AssetType "replacement"))
    {
      Write-Log "Failed to prepare new laptop for assignment" "ERROR"
      return $false
    }
        
    # Update notes on broken laptop
    $noteText = "BROKEN: $breakReason - Processed on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    if (-not (Update-AssetNotes -AssetTag $brokenAssetTag -Notes $noteText))
    {
      Write-Log "Failed to update notes on broken laptop" "WARNING"
    }
        
    # Unassign broken laptop using Reset-SnipeitAssetOwner
    $brokenAsset = Get-SnipeitAsset -asset_tag $brokenAssetTag
    # Reset owner removes assignment and optionally sets status
    $unassignResult = Reset-SnipeitAssetOwner -id $brokenAsset.id -status_id $Config.DamagedStatusId -note "Unassigned due to damage: $breakReason"
    if (-not $unassignResult)
    {
      Write-Log "Failed to unassign broken laptop" "ERROR"
      return $false
    }
    Write-Log "Successfully unassigned broken laptop $brokenAssetTag" "SUCCESS"
        
    # Assign new laptop
    $assignResult = Set-SnipeitAssetOwner -id $newAsset.id -assigned_id $student.id
    if (-not $assignResult)
    {
      Write-Log "Failed to assign new laptop" "ERROR"
      return $false
    }
        
    Write-Log "Exchange completed successfully for student $studentID" "SUCCESS"
        
    # Send success email
    $successSubject = "SUCCESS: Laptop Exchange Completed"
    $successBody = @"
✅ LAPTOP EXCHANGE COMPLETED SUCCESSFULLY ✅

STUDENT INFORMATION:
• Student: $($student.name) (ID: $studentID)

LAPTOP INFORMATION:
• Broken Laptop: $brokenAssetTag (Unassigned and set to Damaged status)
• New Laptop: $newAssetTag (Assigned to $($student.name))
• Break Reason: $breakReason

PROCESSING INFORMATION:
• Processed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
• Log File: $($Config.LogFile)

The laptop exchange has been completed successfully. The broken laptop has been unassigned and marked as damaged, and the new laptop has been assigned to the student.
"@
    Send-AlertEmail -Subject $successSubject -Body $successBody
        
    return $true
  } catch
  {
    Write-Log "Exception during exchange process: $($_.Exception.Message)" "ERROR"
        
    # Send conflict email for processing exception
    Send-ConflictEmail -StudentID $studentID -StudentName $student.name -BrokenAssetTag $brokenAssetTag -NewAssetTag $newAssetTag -BreakReason $breakReason -ConflictType "Processing Error" -ConflictDetails "Exception occurred during exchange processing: $($_.Exception.Message)"
        
    return $false
  }
}

# Main execution
function Start-LaptopExchangeProcess
{
  Write-Log "=== Starting Laptop Exchange Process with Google Sheets Integration ===" "INFO"
    
  # Clean up old log files first
  Remove-OldLogFiles
    
  # Step 1: Download Google Sheets data
  Write-Log "--- Step 1: Downloading Google Sheets Data ---" "INFO"
    
  # Backup existing CSV if it exists
  Backup-ExistingCSV
    
  # Download the Google Sheets data
  if (-not (Get-GoogleSheetData))
  {
    Write-Log "Failed to download Google Sheets data. Checking for existing CSV file..." "ERROR"
        
    # Check if we have a fallback CSV file
    if (-not (Test-Path $Config.CSVFilePath))
    {
      Write-Log "No CSV data available (neither downloaded nor existing file). Cannot proceed." "ERROR"
            
      # Send error email
      $errorSubject = "CRITICAL: Laptop Exchange Process Failed - No Data Available"
      $errorBody = @"
🚨 CRITICAL ERROR: NO DATA AVAILABLE FOR LAPTOP EXCHANGE PROCESS 🚨

DATA DOWNLOAD FAILURE:
• Failed to download Google Sheets data using rclone
• No existing CSV file found as fallback
• Cannot proceed with laptop exchange processing

CONFIGURATION TO CHECK:
• rclone remote name: '$($Config.RcloneRemoteName)'
• Google Sheets file ID: '$($Config.GoogleSheetFileId)'
• Local CSV file: '$($Config.CSVFilePath)'

TROUBLESHOOTING STEPS:
1. Verify rclone is installed and configured
2. Check rclone remote configuration: rclone config
3. Verify Google Sheets file ID is correct
4. Ensure Google Sheets file is accessible to your Google Drive account
5. Test rclone connectivity: rclone lsd $($Config.RcloneRemoteName):

TIME: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
LOG FILE: $($Config.LogFile)

⚠️ ACTION REQUIRED: Please resolve the data access issues and restart the process.
"@
      Send-AlertEmail -Subject $errorSubject -Body $errorBody -Priority "High"
      return
    } else
    {
      Write-Log "Using existing CSV file as fallback" "WARNING"
    }
  } else
  {
    Write-Log "Successfully downloaded Google Sheets data" "SUCCESS"
  }
    
  # Step 2: Connect to SnipeIT
  Write-Log "--- Step 2: Connecting to SnipeIT ---" "INFO"
  if (-not (Connect-SnipeIT))
  {
    Write-Log "Failed to connect to SnipeIT. Please check your snipecred.xml file." "ERROR"
    Write-Log "Make sure the credential file exists and contains valid credentials." "ERROR"
        
    # Send error email
    $errorSubject = "CRITICAL: Laptop Exchange Process Failed - SnipeIT Connection"
    $errorBody = @"
🚨 CRITICAL ERROR: SNIPEIT CONNECTION FAILED 🚨

CONNECTION FAILURE:
• Failed to connect to SnipeIT using stored credentials
• Cannot proceed with laptop exchange processing

TROUBLESHOOTING STEPS:
1. Verify 'snipecred.xml' file exists in the current directory
2. Check that the credentials in the file are valid and current
3. Test SnipeIT connectivity manually
4. Ensure SnipeIT server is accessible

TIME: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
LOG FILE: $($Config.LogFile)

⚠️ ACTION REQUIRED: Please resolve the SnipeIT connection issues and restart the process.
"@
    Send-AlertEmail -Subject $errorSubject -Body $errorBody -Priority "High"
    return
  }
    
  # Step 3: Validate configuration
  Write-Log "--- Step 3: Validating Configuration ---" "INFO"
  Write-Log "DamagedStatusId: $($Config.DamagedStatusId)" "INFO"
  Write-Log "ReadyToDeployStatusId: $($Config.ReadyToDeployStatusId)" "INFO"
  Write-Log "Google Sheets File ID: $($Config.GoogleSheetFileId)" "INFO"
  Write-Log "rclone Remote: $($Config.RcloneRemoteName)" "INFO"
    
  if ($Config.DamagedStatusId -le 0 -or $Config.ReadyToDeployStatusId -le 0)
  {
    Write-Log "ERROR: Invalid status ID configuration. Please set proper status IDs in the Config section." "ERROR"
    Write-Log "DamagedStatusId and ReadyToDeployStatusId must be greater than 0." "ERROR"
    return
  }
    
  # Step 4: Get available status labels to help with configuration
  Write-Log "--- Step 4: Checking Available Status Labels ---" "INFO"
  $statusLabels = Get-SnipeitStatusLabels
    
  try
  {
    # Step 5: Process CSV data
    Write-Log "--- Step 5: Processing Laptop Exchange Data ---" "INFO"
        
    # Get CSV data
    $exchangeData = Get-CSVData
    if (-not $exchangeData)
    {
      Write-Log "Failed to load CSV data. Exiting." "ERROR"
            
      # Send error email
      $errorSubject = "CRITICAL: Laptop Exchange Process Failed - CSV Data Error"
      $errorBody = @"
🚨 CRITICAL ERROR: CSV DATA PROCESSING FAILED 🚨

DATA PROCESSING FAILURE:
• Successfully downloaded/found CSV file
• Failed to parse or validate CSV data
• Cannot proceed with laptop exchange processing

POSSIBLE ISSUES:
• CSV file format is incorrect
• Missing required headers/columns
• Empty or corrupted CSV file
• Encoding issues

TIME: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
LOG FILE: $($Config.LogFile)

⚠️ ACTION REQUIRED: Please check the CSV file format and data, then restart the process.
"@
      Send-AlertEmail -Subject $errorSubject -Body $errorBody -Priority "High"
      return
    }
        
    # Process each exchange request
    $successCount = 0
    $failCount = 0
    $rowNumber = 0
        
    foreach ($row in $exchangeData)
    {
      $rowNumber++
      Write-Log "--- Processing Row $rowNumber ---" "INFO"
            
      if (Process-LaptopExchange -ExchangeData $row)
      {
        $successCount++
      } else
      {
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
📊 LAPTOP EXCHANGE BATCH PROCESS SUMMARY 📊

PROCESSING RESULTS:
• Total requests processed: $($exchangeData.Count)
• Successful exchanges: $successCount
• Failed exchanges: $failCount

DATA SOURCE INFORMATION:
• Google Sheets File ID: $($Config.GoogleSheetFileId)
• rclone Remote: $($Config.RcloneRemoteName)
• Downloaded to: $($Config.LocalCSVFileName)

PROCESSING INFORMATION:
• Log file: $($Config.LogFile)
• Processed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

$(if ($failCount -gt 0) { "⚠️  NOTE: $failCount exchange(s) failed and require manual review. Check conflict emails for details." })

All failed exchanges have been documented with detailed conflict emails containing:
• Student information (ID and name)
• Laptop details (broken and new asset tags)
• Break reason
• Specific conflict details
• Processing timestamp

Please review all conflict emails and resolve issues manually before reprocessing failed exchanges.

NEXT STEPS:
• Review any conflict emails for failed exchanges
• Update the Google Sheets with resolution status if needed
• Re-run the script if necessary after resolving conflicts
"@
    Send-AlertEmail -Subject $summarySubject -Body $summaryBody
        
  } catch
  {
    Write-Log "Critical exception in main process: $($_.Exception.Message)" "ERROR"
        
    $errorSubject = "CRITICAL: Laptop Exchange Process Failed"
    $errorBody = @"
🚨 CRITICAL ERROR: LAPTOP EXCHANGE PROCESS FAILED 🚨

SYSTEM ERROR DETAILS:
• Error: $($_.Exception.Message)
• Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
• Log file: $($Config.LogFile)

DATA SOURCE INFORMATION:
• Google Sheets File ID: $($Config.GoogleSheetFileId)
• rclone Remote: $($Config.RcloneRemoteName)

⚠️ ACTION REQUIRED: Please review the logs and restart the process manually.

The laptop exchange process encountered a critical system error and was unable to complete. Please check the log file for detailed error information and resolve any system issues before restarting the process.
"@
    Send-AlertEmail -Subject $errorSubject -Body $errorBody -Priority "High"
  }
}

# Execute the main function
Start-LaptopExchangeProcess

# Configuration Help and Setup Instructions
Write-Host ""
Write-Host "=== CONFIGURATION HELP ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "To set up this script with Google Sheets integration:" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. RCLONE SETUP:" -ForegroundColor Green
Write-Host "   - Install rclone from: https://rclone.org/downloads/" -ForegroundColor White
Write-Host "   - Configure Google Drive remote: rclone config" -ForegroundColor White
Write-Host "   - Follow the Google Drive setup instructions" -ForegroundColor White
Write-Host "   - Test your setup: rclone lsd gdrive:" -ForegroundColor White
Write-Host ""
Write-Host "2. GOOGLE SHEETS CONFIGURATION:" -ForegroundColor Green
Write-Host "   - Open your Google Sheets file" -ForegroundColor White
Write-Host "   - Copy the file ID from the URL:" -ForegroundColor White
Write-Host "     https://docs.google.com/spreadsheets/d/FILE_ID_HERE/edit" -ForegroundColor White
Write-Host "   - Update RcloneRemoteName and GoogleSheetFileId in the Config section" -ForegroundColor White
Write-Host ""
Write-Host "3. REQUIRED HEADERS IN GOOGLE SHEETS:" -ForegroundColor Green
Write-Host "   - What is your student ID number? (All 7 digits)" -ForegroundColor White
Write-Host "   - How did you break your laptop?" -ForegroundColor White
Write-Host "   - What is the Asset Tag (The numbers on the red sticker) on the broken laptop?" -ForegroundColor White
Write-Host "   - What is the Asset Tag (The numbers on the red sticker) on the new laptop?" -ForegroundColor White
Write-Host ""
Write-Host "4. EMAIL CONFIGURATION:" -ForegroundColor Green
Write-Host "   - Set up Gmail App Password at: https://myaccount.google.com/apppasswords" -ForegroundColor White
Write-Host "   - Update EmailSender and EmailAppPassword in the Config section" -ForegroundColor White
Write-Host ""
