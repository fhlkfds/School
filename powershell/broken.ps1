# Define Email Parameters
param(
  [Parameter(Mandatory=$false)]
  [string]$EmailTo,
    
  [Parameter(Mandatory=$false)]
  [string]$EmailFrom,
    
  [Parameter(Mandatory=$false)]
  [SecureString]$EmailPassword,
    
  [Parameter(Mandatory=$false)]
  [string]$SmtpServer = "smtp.gmail.com",
    
  [Parameter(Mandatory=$false)]
  [int]$SmtpPort = 587
)

# Default email settings - Change these to your values
$defaultEmailTo = ""  # Your recipient email
$defaultEmailFrom = ""  # Your Gmail address
$defaultEmailPasswordString = ""  # Your Gmail app password

# Use provided parameters or defaults
$EmailTo = if ($EmailTo)
{ $EmailTo 
} else
{ $defaultEmailTo 
}
$EmailFrom = if ($EmailFrom)
{ $EmailFrom 
} else
{ $defaultEmailFrom 
}

# Convert default password to SecureString if not provided
if (-not $EmailPassword)
{
  $EmailPassword = ConvertTo-SecureString $defaultEmailPasswordString -AsPlainText -Force
}

# Paths configuration
$CsvPath = "/home/liam/broken.csv"
$HistoryFilePath = "/home/liam/processed_exchanges.csv"

# First, use rclone to pull the latest data from Google Sheets
Write-Output "📥 Pulling latest data from Google Sheets..."
try
{
  # Assuming you have rclone configured with a remote named "gdrive" pointing to Google Drive
  # and the sheet is exported/published as CSV
    
  # rclone command to copy the Google Sheet (as CSV) to the local file
  # Replace these parameters with your actual rclone configuration and Google Sheet details
  $rcloneCommand = "rclone copy 'gdrive:path/to/your/sheet.csv' /home/liam/ --no-traverse"
    
  # Execute the rclone command
  Write-Output "Executing: $rcloneCommand"
  Invoke-Expression $rcloneCommand
    
  # Check if the file exists after download
  if (Test-Path -Path $CsvPath)
  {
    Write-Output "✅ Successfully downloaded CSV from Google Sheets to $CsvPath"
  } else
  {
    Write-Output "⚠️ Warning: CSV file not found at $CsvPath after rclone command"
    Write-Output "Continuing with existing file if available..."
  }
} catch
{
  Write-Output "❌ Error pulling data from Google Sheets: $_"
  Write-Output "Continuing with existing file if available..."
}

Write-Output "Using CSV path: $CsvPath"
Write-Output "Using history file: $HistoryFilePath"

# Create history file if it doesn't exist
if (-not (Test-Path -Path $HistoryFilePath))
{
  "StudentID,BrokenAssetTag,NewAssetTag,ProcessedDate" | Out-File -FilePath $HistoryFilePath
  Write-Output "Created new history file"
}

# Load previously processed records
$processedRecords = @{}
if (Test-Path -Path $HistoryFilePath)
{
  Import-Csv -Path $HistoryFilePath | ForEach-Object {
    # Create a unique key for each processed record
    $key = "$($_.StudentID)_$($_.BrokenAssetTag)_$($_.NewAssetTag)"
    $processedRecords[$key] = $_.ProcessedDate
  }
}
Write-Output "Loaded $($processedRecords.Count) previously processed records"

# Read the CSV (comma-delimited) and filter out empty records
$records = Import-Csv -Path $CsvPath -Delimiter "," | 
  Where-Object { ![string]::IsNullOrWhiteSpace($_.'What is your student ID number? (All 7 digits)') }

Write-Output "Found $($records.Count) valid records to process"

# Initialize array to collect issues for email reporting
$emailIssues = @()

# Process each record in the CSV
foreach ($record in $records)
{
  # Extract details from CSV
  $firstName = $record.'Cadet First Name:'
  $lastName = $record.'Cadet Last Name:'
  $studentID = $record.'What is your student ID number? (All 7 digits)'
  $brokenAssetTag = $record.'What is the Asset Tag (The numbers on the red sticker) on the broken laptop?'
  $newAssetTag = $record.'What is the Asset Tag (The numbers on the red sticker) on the new laptop?'
    
  # Create a unique key for this record
  $recordKey = "${studentID}_${brokenAssetTag}_${newAssetTag}"
    
  # Check if this record has already been processed
  if ($processedRecords.ContainsKey($recordKey))
  {
    $processedDate = $processedRecords[$recordKey]
    Write-Output "⏭️ Skipping already processed exchange for $firstName $lastName (ID: $studentID)"
    Write-Output "   Broken: $brokenAssetTag → New: $newAssetTag (Processed on: $processedDate)"
    continue
  }
    
  Write-Output "Processing exchange for: $firstName $lastName (ID: $studentID)"
  Write-Output "  Broken laptop: $brokenAssetTag → New laptop: $newAssetTag"

  try
  {
    # 1. Get the User by student ID
    Write-Output "🔍 Searching for user: $studentID"
    $user = Get-SnipeitUser -Search $studentID
        
    # Handle the case where we get multiple results
    if ($user -is [array])
    {
      $user = $user[0]  # Use the first result
    }
        
    if (!$user)
    {
      throw "No user found with student ID: $studentID"
    }
        
    $userID = $user.id
    $userName = $user.name
    Write-Output "🆔 Found User: $userName (ID: $userID)"
        
    # 2. Process the broken laptop
    Write-Output "🔍 Finding broken laptop: $brokenAssetTag"
    $brokenAsset = Get-SnipeitAsset -asset_tag $brokenAssetTag
        
    if (!$brokenAsset)
    {
      throw "Could not find broken laptop with tag: $brokenAssetTag"
    }
        
    $brokenAssetID = $brokenAsset.id
    Write-Output "🆔 Found Broken Laptop ID: $brokenAssetID"
        
    # 3. Process the new laptop
    Write-Output "🔍 Finding new laptop: $newAssetTag"
    $newAsset = Get-SnipeitAsset -asset_tag $newAssetTag
        
    if (!$newAsset)
    {
      throw "Could not find new laptop with tag: $newAssetTag"
    }
        
    if ($newAsset.assigned_to)
    {
      throw "New laptop $newAssetTag is already assigned to $($newAsset.assigned_to.name)"
    }
        
    if ($newAsset.status_label.status_meta -eq "undeployable")
    {
      throw "New laptop $newAssetTag is marked as undeployable"
    }
        
    $newAssetID = $newAsset.id
    Write-Output "🆔 Found New Laptop ID: $newAssetID"
        
    # 4. Do both operations
    # First archive the broken laptop
    Write-Output "📦 Archiving broken laptop: $brokenAssetTag"
    Reset-SnipeitAssetOwner -id $brokenAssetID -status_id 4  # Status 4 is typically "Archived"
    Write-Output "✔️ Archived broken laptop: $brokenAssetTag"
        
    # Then assign the new laptop
    Write-Output "💻 Assigning new laptop: $newAssetTag to $userName"
    Set-SnipeitAssetOwner -id $newAssetID -assigned_id $userID
    Set-SnipeitAsset -id $newAssetID -status_id 1  # Status 1 is typically "Deployed"
    Write-Output "✔️ Assigned new laptop: $newAssetTag to $userName"
        
    Write-Output "✅ Successfully completed exchange for $firstName $lastName"
        
    # Record this successful exchange in the history file
    $today = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$studentID,$brokenAssetTag,$newAssetTag,$today" | Out-File -FilePath $HistoryFilePath -Append
    Write-Output "✍️ Recorded exchange in history file"
  } catch
  {
    $errorMsg = "Error processing record for $firstName $lastName ($studentID): $_"
    Write-Output "❗ $errorMsg"
    $emailIssues += [PSCustomObject]@{
      FirstName = $firstName
      LastName = $lastName
      StudentID = $studentID
      Issue = "Processing error"
      Details = $errorMsg
    }
  }
}

# Send email with issues if there are any
if ($emailIssues.Count -gt 0)
{
  Write-Output "📧 Sending email report with $($emailIssues.Count) issues..."
    
  # Create email body
  $emailBody = "<h2>Laptop Exchange Issues Report</h2>"
  $emailBody += "<p>The following issues were encountered during the laptop exchange process:</p>"
  $emailBody += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"
  $emailBody += "<tr><th>Student Name</th><th>Student ID</th><th>Issue Type</th><th>Details</th></tr>"
    
  foreach ($issue in $emailIssues)
  {
    $emailBody += "<tr>"
    $emailBody += "<td>$($issue.FirstName) $($issue.LastName)</td>"
    $emailBody += "<td>$($issue.StudentID)</td>"
    $emailBody += "<td>$($issue.Issue)</td>"
    $emailBody += "<td>$($issue.Details)</td>"
    $emailBody += "</tr>"
  }
    
  $emailBody += "</table>"
    
  # Write report to file as backup in case email fails
  $reportPath = Join-Path (Get-Location) "LaptopExchangeReport-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
  $emailBody | Out-File -FilePath $reportPath
  Write-Output "Report saved to: $reportPath"
    
  # Send email using Google SMTP
  try
  {
    # Set TLS 1.2 which Gmail requires
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
    # Create credential object for Gmail
    $securePassword = ConvertTo-SecureString $defaultEmailPasswordString -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($EmailFrom, $securePassword)
        
    # Handle email sending
    Write-Output "Sending email to $EmailTo using $EmailFrom via $SmtpServer port $SmtpPort"
        
    # Use Send-MailMessage directly
    $emailParams = @{
      To = $EmailTo
      From = $EmailFrom
      Subject = "Laptop Exchange Issues Report - $(Get-Date -Format 'yyyy-MM-dd')"
      Body = $emailBody
      BodyAsHtml = $true
      SmtpServer = $SmtpServer
      Port = $SmtpPort
      Credential = $credential
      UseSsl = $true
      ErrorAction = "Stop"
    }
        
    # Suppress the warning about Send-MailMessage being obsolete
    $ProgressPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'
        
    Send-MailMessage @emailParams
        
    # Restore preferences
    $ProgressPreference = 'Continue'
    $WarningPreference = 'Continue'
        
    Write-Output "✔️ Email report sent successfully to $EmailTo"
  } catch
  {
    Write-Output "❌ Failed to send email report: $_"
    Write-Output "Please check your email credentials and server settings."
    Write-Output "The report has been saved to: $reportPath"
  }
}

Write-Output "✔️ Process completed"
