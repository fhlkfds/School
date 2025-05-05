
# Get the first argument (equivalent to $1 in bash)
$Email = $args[0]

# Check if the argument is missing
if (-not $Email)
{
  Write-Host "Usage: .\password_reset.ps1 user@example.com"
  exit 1
}



gam update user $email password "Password@1" changepassword on


