# Check if a CSV file path is passed as a parameter
param(
  [string]$CSV_FILE = $null
)

if (-not $CSV_FILE)
{
  Write-Host "Usage: .\create_users.ps1 -CSV_FILE <path_to_csv_file>"
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

# Read the CSV file
$csvLines = Import-Csv -Path $CSV_FILE

foreach ($user in $csvLines)
{
  $firstName = $user.first_name.Trim()
  $lastName = $user.last_name.Trim()
  $grade = $user.grade.Trim()
    
  # Use a default password
  $password = "Password@1"

  if ([string]::IsNullOrWhiteSpace($firstName) -or 
    [string]::IsNullOrWhiteSpace($lastName) -or 
    [string]::IsNullOrWhiteSpace($grade))
  {
    Write-Host "Skipping user with incomplete information: $firstName $lastName"
    continue
  }

  # Extract the first letter of the first name
  $firstInitial = $firstName.Substring(0,1).ToLower()

  # Set the email format for nomma.net
  $email = "cadet$firstInitial$lastName@nomma.net".ToLower()

  $gradeth = "$grade" + "th"
  # Set the organizational unit based on the grade
  $orgUnit = "/Cadets/$gradeth Grade"

  Write-Host "Creating user: $email in OU: $orgUnit"

  # Create the user with GAM
  gam create user $email firstname "$firstName" lastname "$lastName" password "$password" org "$orgUnit"

  # Add logging or error handling as needed
}

