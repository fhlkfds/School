# Import the PSGSuite module
Import-Module PSGSuite

# Domain for the primary email
$domain = "test.lan"

# Path to the CSV file
$csvPath = "C:\path\to\users.csv"

# Read the CSV file
$users = Import-Csv $csvPath

foreach ($user in $users)
{
  try
  {
    # Generate the primary email in the "cadet<first initial><last name>" format
    $firstInitial = $user.FirstName.Substring(0,1).ToLower()
    $lastName = $user.LastName.ToLower()
    $primaryEmail = "cadet$firstInitial$lastName@$domain"

    # Generate a standard password
    $password = "Password@1"

    # Determine the Org Unit Path based on the grade
    switch ($user.Grade)
    {
      8
      { $orgUnitPath = "/Students/MiddleSchool/Grade8" 
      }
      9
      { $orgUnitPath = "/Students/HighSchool/Grade9" 
      }
      10
      { $orgUnitPath = "/Students/HighSchool/Grade10" 
      }
      11
      { $orgUnitPath = "/Students/HighSchool/Grade11" 
      }
      12
      { $orgUnitPath = "/Students/HighSchool/Grade12" 
      }
      default
      {
        Write-Output "Invalid grade $($user.Grade) for $($user.FirstName) $($user.LastName), skipping..."
        continue
      }
    }

    # Create the user
    New-GSUser -PrimaryEmail $primaryEmail `
      -GivenName $user.FirstName `
      -FamilyName $user.LastName `
      -Password $password `
      -OrgUnitPath $orgUnitPath `
      -ChangePasswordAtNextLogin $true

    Write-Output "Created user: $primaryEmail in $orgUnitPath"
        
    # Add the user to the appropriate grade group
    $groupEmail = "grade$($user.Grade)@example.com"
    Add-GSGroupMember -Group $groupEmail -Email $primaryEmail
    Write-Output "Added $primaryEmail to group $groupEmail"
  } catch
  {
    Write-Output "Failed to create user or add to group: $($user.FirstName) $($user.LastName) - $_"
  }
}

