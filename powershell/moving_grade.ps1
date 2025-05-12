# Prompt for input
$csv_file = Read-Host "Enter the full path to the CSV file (must contain 'email' column)"
$new_grade = Read-Host "Enter the grade these users will be in (8-12)"

# Validate the grade input

$base_ou = "/Cadets/"
$destination_ou = "$base_ou$new_grade" + "th Grade/"

# CSV File Format Info
Write-Host "💡 Make sure your CSV file has the following format:" -ForegroundColor Yellow
Write-Host "email" -ForegroundColor Green
Write-Host "user1@example.com" -ForegroundColor Green
Write-Host "user2@example.com" -ForegroundColor Green
Write-Host "user3@example.com" -ForegroundColor Green

# Check if file exists
if (-Not (Test-Path $csv_file))
{
  Write-Host "❌ CSV file not found at $csv_file" -ForegroundColor Red
  exit 1
}

Write-Host "⏳ Starting to move users to $destination_ou..." -ForegroundColor Cyan

# Move users to the selected grade OU
& gam csv "$csv_file" gam update user ~email org "$destination_ou"

Write-Host "✅ All users moved successfully!" -ForegroundColor Green
