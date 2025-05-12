# Check if a CSV file path is passed as a parameter
param(
  [string]$CSV_FILE = $null
)

if (-not $CSV_FILE)
{
  Write-Host "Usage: .\pw_archive_users.ps1 -CSV_FILE <path_to_csv_file>"
  exit 1
}

# Expand the file path if it starts with '~'
if ($CSV_FILE.StartsWith("~"))
{
  $CSV_FILE = $CSV_FILE -replace "^~", (Get-Item -Path "~").FullName
}

# Check if the CSV file exists
if (-Not (Test-Path -Path $CSV_FILE))
{
  Write-Host "CSV file not found: $CSV_FILE"
  exit 1
}

# Read the CSV lines skipping the header
$csvLines = Get-Content -Path $CSV_FILE | Select-Object -Skip 1

foreach ($line in $csvLines)
{
  $columns = $line -split ','

  if ($columns.Count -lt 1)
  {
    Write-Host "Skipping line with insufficient columns"
    continue
  }

  $username = $columns[0].Trim()

  if ([string]::IsNullOrWhiteSpace($username))
  {
    Write-Host "Skipping empty username"
    continue
  }

  Write-Host "Processing user: $username"

  # Suspend the user
  gam update user $username suspended on

  # Delete all aliases
  gam user $username delete aliases

  # Move to 'Inactive' org unit
  gam update user $username org "/Inactive"

  # Email transformations
  $domain = $username.Split("@")[1]
  $localPart = $username.Split("@")[0]
  $newEmail = "inactive$localPart@$domain"

  # Update user's primary email
  gam update user $username email $newEmail
  Start-Sleep -Seconds 10

  # Delete alias from the new user
  Write-Host "Deleting alias $username from user $newEmail"
  gam user $newEmail delete alias $username
}

