# Define CSV Path and Email Parameters
param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$EmailTo = "ldecareaux@nomma.net",
    
    [Parameter(Mandatory=$false)]
    [string]$EmailFrom = "scan2@nomma.net",
    
    [Parameter(Mandatory=$false)]
    [string]$EmailPasswordString = "uzds yjqo nwsx gimo",
    
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer = "smtp.gmail.com",
    
    [Parameter(Mandatory=$false)]
    [int]$SmtpPort = 587,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipEmail
)

# Read the CSV (comma-delimited)
$records = Import-Csv -Path $CsvPath -Delimiter ","

# Initialize array to collect issues for reporting
$issuesList = @()

# Process each record in the CSV
foreach ($record in $records)
{
    # Extract details from CSV
    $firstName = $record.'Cadet First Name:'
    $lastName = $record.'Cadet Last Name:'
    $studentID = $record.'What is your student ID number? (All 7 digits)'
    $brokenAssetTag = $record.'What is the Asset Tag (The numbers on the red sticker) on the broken laptop?'
    $newAssetTag = $record.'What is the Asset Tag (The numbers on the red sticker) on the new laptop?'

    try
    {
        # 1. Get the User Internal ID
        Write-Output "🔍 Searching for user by employee number: $studentID"
        $user = Get-SnipeitUser -Search $studentID | Where-Object { $_.employee_num -eq $studentID }
        if (!$user)
        {
            $errorMsg = "Could not find user with employee number: $studentID for $firstName $lastName"
            Write-Output "❗ $errorMsg"
            $issuesList += [PSCustomObject]@{
                FirstName = $firstName
                LastName = $lastName
                StudentID = $studentID
                Issue = "Invalid employee ID number"
                Details = $errorMsg
            }
            continue
        }
        $userID = $user.id
        $userName = $user.name
        Write-Output "🆔 Found User: $userName (ID: $userID)"

        # 2. Get the Broken Laptop Internal ID - Use exact asset tag search
        Write-Output "🔍 Searching for broken asset with exact tag: $brokenAssetTag"
        # First try direct asset tag match
        $brokenAsset = Get-SnipeitAsset -asset_tag $brokenAssetTag
    
        if (!$brokenAsset)
        {
            $errorMsg = "Could not find broken laptop with exact tag: $brokenAssetTag for $firstName $lastName"
            Write-Output "❗ $errorMsg"
            $issuesList += [PSCustomObject]@{
                FirstName = $firstName
                LastName = $lastName
                StudentID = $studentID
                Issue = "Invalid broken laptop asset tag"
                Details = $errorMsg
            }
            continue
        }
    
        # Check if broken laptop is already archived
        if ($brokenAsset.status_label.status_meta -eq "archived")
        {
            $errorMsg = "Broken laptop $brokenAssetTag is already archived for $firstName $lastName"
            Write-Output "⚠️ $errorMsg"
            $issuesList += [PSCustomObject]@{
                FirstName = $firstName
                LastName = $lastName
                StudentID = $studentID
                Issue = "Already archived laptop"
                Details = $errorMsg
            }
            continue
        }
        $brokenAssetID = $brokenAsset.id
        Write-Output "🆔 Found Broken Laptop: $brokenAssetTag (ID: $brokenAssetID)"
    
        # 3. Unassign and Archive the Broken Laptop
        Write-Output "📦 Archiving broken laptop: $brokenAssetTag (ID: $brokenAssetID)"
        Reset-SnipeitAssetOwner -id $brokenAssetID -status_id 4
        Write-Output "✔️ Archived and unassigned broken laptop: $brokenAssetTag"

        # 4. Get the New Laptop Internal ID - Use exact asset tag search
        Write-Output "🔍 Searching for new asset with exact tag: $newAssetTag"
        # First try direct asset tag match
        $newAsset = Get-SnipeitAsset -asset_tag $newAssetTag
    
        if (!$newAsset)
        {
            $errorMsg = "Could not find new laptop with exact tag: $newAssetTag for $firstName $lastName"
            Write-Output "❗ $errorMsg"
            $issuesList += [PSCustomObject]@{
                FirstName = $firstName
                LastName = $lastName
                StudentID = $studentID
                Issue = "Invalid new laptop asset tag"
                Details = $errorMsg
            }
            continue
        }
    
        # 5. Check if the new laptop is already deployed
        if ($newAsset.assigned_to)
        {
            $errorMsg = "New laptop $newAssetTag is already assigned to $($newAsset.assigned_to.name)"
            Write-Output "⚠️ $errorMsg"
            $issuesList += [PSCustomObject]@{
                FirstName = $firstName
                LastName = $lastName
                StudentID = $studentID
                Issue = "Already deployed laptop"
                Details = $errorMsg
            }
            continue
        }
    
        # 6. Check if the new laptop is deployable
        if ($newAsset.status_label.status_meta -eq "undeployable")
        {
            $errorMsg = "Laptop $newAssetTag is marked as undeployable. Status: $($newAsset.status_label.name)"
            Write-Output "⚠️ $errorMsg"
            $issuesList += [PSCustomObject]@{
                FirstName = $firstName
                LastName = $lastName
                StudentID = $studentID
                Issue = "Undeployable laptop"
                Details = $errorMsg
            }
            continue
        }
        $newAssetID = $newAsset.id
        Write-Output "🆔 Found New Laptop: $newAssetTag (ID: $newAssetID)"
    
        # 7. Assign the New Laptop to the User
        Write-Output "💻 Assigning new laptop: $newAssetTag (ID: $newAssetID) to user: $userID"
        Set-SnipeitAssetOwner -id $newAssetID -assigned_id $userID 
        Set-SnipeitAsset -id $newAssetID -status_id 1  
        Write-Output "✔️ Assigned new laptop: $newAssetTag to $userName (User ID: $userID)"
    
    } catch
    {
        $errorMsg = "Error processing record for $firstName $lastName ($studentID): $_"
        Write-Output "❗ $errorMsg"
        $issuesList += [PSCustomObject]@{
            FirstName = $firstName
            LastName = $lastName
            StudentID = $studentID
            Issue = "Processing error"
            Details = $errorMsg
        }
    }
}

# Create reports for any issues
if ($issuesList.Count -gt 0)
{
    Write-Output "📊 Creating report for $($issuesList.Count) issues..."
    
    # Get current timestamp for filenames
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    
    # Create HTML report file
    $htmlReportPath = Join-Path (Get-Location) "LaptopExchangeReport-$timestamp.html"
    
    # Create CSV report file (easier to import into spreadsheets)
    $csvReportPath = Join-Path (Get-Location) "LaptopExchangeReport-$timestamp.csv"
    
    # Create HTML report body
    $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <title>Laptop Exchange Issues Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #003366; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th { background-color: #003366; color: white; padding: 8px; text-align: left; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .summary { margin: 20px 0; padding: 10px; background-color: #e6f2ff; border-radius: 5px; }
    </style>
</head>
<body>
    <h1>Laptop Exchange Issues Report</h1>
    <div class="summary">
        <p><strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p><strong>Total Issues:</strong> $($issuesList.Count)</p>
    </div>
    <table>
        <tr>
            <th>Student Name</th>
            <th>Student ID</th>
            <th>Issue Type</th>
            <th>Details</th>
        </tr>
"@

    # Add each issue to the HTML report
    foreach ($issue in $issuesList)
    {
        $htmlBody += @"
        <tr>
            <td>$($issue.FirstName) $($issue.LastName)</td>
            <td>$($issue.StudentID)</td>
            <td>$($issue.Issue)</td>
            <td>$($issue.Details)</td>
        </tr>
"@
    }

    # Close the HTML
    $htmlBody += @"
    </table>
</body>
</html>
"@

    # Save the HTML report
    $htmlBody | Out-File -FilePath $htmlReportPath -Encoding UTF8
    Write-Output "✅ HTML Report saved to: $htmlReportPath"
    
    # Export to CSV (for easier data processing)
    $issuesList | Export-Csv -Path $csvReportPath -NoTypeInformation
    Write-Output "✅ CSV Report saved to: $csvReportPath"
    
    # Try to send email only if not skipped
    if (-not $SkipEmail)
    {
        try
        {
            Write-Output "📧 Attempting to send email report..."
            
            # Set TLS 1.2
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            
            # Create credential
            $securePassword = ConvertTo-SecureString $EmailPasswordString -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($EmailFrom, $securePassword)
            
            # Setup email parameters
            $emailParams = @{
                To = $EmailTo
                From = $EmailFrom
                Subject = "Laptop Exchange Issues Report - $(Get-Date -Format 'yyyy-MM-dd')"
                Body = $htmlBody
                BodyAsHtml = $true
                SmtpServer = $SmtpServer
                Port = $SmtpPort
                Credential = $credential
                UseSsl = $true
                ErrorAction = "Stop"
            }
            
            # Try to send email
            Send-MailMessage @emailParams
            Write-Output "✅ Email sent successfully to $EmailTo"
        }
        catch
        {
            Write-Output "❌ Email sending failed: $_"
            Write-Output "The reports have been saved locally - please share them manually."
        }
    }
    else
    {
        Write-Output "📧 Email sending skipped as requested."
        Write-Output "Please share the reports manually from the saved files."
    }
    
    # Open the HTML report in the default browser if on Windows
    if ($IsWindows)
    {
        Start-Process $htmlReportPath
    }
    else
    {
        Write-Output "To view the HTML report, open: $htmlReportPath"
    }
}
else
{
    Write-Output "✅ No issues found during the laptop exchange process."
}
