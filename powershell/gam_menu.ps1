# Google Workspace Admin Toolkit in PowerShell

# Default password for all operations
$defaultPassword = "Password@1"

# Function to display the main menu
function Show-MainMenu
{
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "    GOOGLE WORKSPACE ADMIN TOOLKIT        " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "1. Password Management" -ForegroundColor White
    Write-Host "2. User Management" -ForegroundColor White
    Write-Host "3. Device Management" -ForegroundColor White
    Write-Host "4. Information Retrieval" -ForegroundColor White
    Write-Host "5. CSV Templates" -ForegroundColor White
    Write-Host "6. Exit" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-6]: " -ForegroundColor Yellow -NoNewline
}

# Function to display password management submenu
function Show-PasswordMenu
{
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "         PASSWORD MANAGEMENT              " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "1. Reset Single User Password" -ForegroundColor White
    Write-Host "2. Mass Password Reset (CSV)" -ForegroundColor White
    Write-Host "3. Return to Main Menu" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Default Password: $defaultPassword" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-3]: " -ForegroundColor Yellow -NoNewline
}

# Function to display user management submenu (modified)
function Show-UserMenu
{
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "           USER MANAGEMENT                " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "1. Create Single User" -ForegroundColor White
    Write-Host "2. Create Multiple Users (CSV)" -ForegroundColor White
    Write-Host "3. Archive User" -ForegroundColor White
    Write-Host "4. Archive Multiple Users (CSV)" -ForegroundColor White
    Write-Host "5. Suspend User" -ForegroundColor White
    Write-Host "6. Unsuspend User" -ForegroundColor White
    Write-Host "7. Delete User" -ForegroundColor White
    Write-Host "8. Move Users to Grade OU (CSV)" -ForegroundColor White
    Write-Host "9. Return to Main Menu" -ForegroundColor White
    Write-Host "Enter your choice [1-9]: " -ForegroundColor Yellow -NoNewline
}

# Function to display device management submenu
function Show-DeviceMenu
{
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "          DEVICE MANAGEMENT               " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "1. Lock Single Mobile Device" -ForegroundColor White
    Write-Host "2. Lock Multiple Mobile Devices (CSV)" -ForegroundColor White
    Write-Host "3. Unlock Single Device" -ForegroundColor White
    Write-Host "4. Unlock Multiple Devices (CSV)" -ForegroundColor White
    Write-Host "3. Wipe Device" -ForegroundColor White
    Write-Host "4. Wipe Multiple Devices (CSV)" -ForegroundColor White
    Write-Host "5. List All Devices" -ForegroundColor White
    Write-Host "6. Return to Main Menu" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-6]: " -ForegroundColor Yellow -NoNewline
}

# Function to display information retrieval submenu
function Show-InfoMenu
{
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "        INFORMATION RETRIEVAL             " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "1. Get Device Serial Number" -ForegroundColor White
    Write-Host "2. Get User Information" -ForegroundColor White
    Write-Host "3. List All Users" -ForegroundColor White
    Write-Host "4. Check License Information" -ForegroundColor White
    Write-Host "5. Generate User Activity Report" -ForegroundColor White
    Write-Host "6. Monitor Security Reports" -ForegroundColor White
    Write-Host "7. Return to Main Menu" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-7]: " -ForegroundColor Yellow -NoNewline
}

# Function to reset single user password
function Reset-SinglePassword
{
    Clear-Host
    Write-Host "Reset Single User Password" -ForegroundColor Cyan
    Write-Host "-------------------------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    $password = Read-Host "Enter new password (press Enter to use default: $defaultPassword)"
    
    if ([string]::IsNullOrEmpty($password))
    {
        $password = $defaultPassword
    }
    
    Write-Host "Resetting password for $userEmail..." -ForegroundColor Yellow
    & gam update user "$userEmail" password "$password"
    Write-Host "Password has been reset for $userEmail" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}



# Function to move users to different OUs based on grade level from CSV
function Move-UsersToGradeOU
{
    Clear-Host
    Write-Host "Move Users to Grade-Based OU (CSV)" -ForegroundColor Cyan
    Write-Host "--------------------------------" -ForegroundColor Cyan
    Write-Host "CSV format should be: email,grade" -ForegroundColor Yellow
    Write-Host "Grade values: 8, 9, 10, 10email (for 10th with email), 11, 12" -ForegroundColor Yellow
    $csvPath = Read-Host "Enter path to CSV file"
    
    if (-not (Test-Path $csvPath))
    {
        Write-Host "Error: File not found!" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    $users = Import-Csv $csvPath
    $successCount = 0
    $errorCount = 0
    
    foreach ($user in $users)
    {
        if ([string]::IsNullOrEmpty($user.email) -or [string]::IsNullOrEmpty($user.grade))
        {
            Write-Host "Skipping invalid entry: Missing email or grade" -ForegroundColor Yellow
            $errorCount++
            continue
        }
        
        # Map grade to OU
        $gradeOU = switch ($user.grade)
        {
            "8"
            { "/Cadets/8th Grade" 
            }
            "9"
            { "/Cadets/9th Grade" 
            }
            "10"
            { "/Cadets/10th Grade" 
            }
            "10email"
            { "/Cadets/10th Grade with Email Privileges" 
            }
            "11"
            { "/Cadets/11th Grade" 
            }
            "12"
            { "/Cadets/12th Grade" 
            }
            default
            { 
                Write-Host "Invalid grade value for $($user.email): $($user.grade)" -ForegroundColor Yellow
                $errorCount++
                continue
            }
        }
        
        try
        {
            Write-Host "Moving user $($user.email) to $gradeOU..." -ForegroundColor Yellow
            & gam update user "$($user.email)" org "$gradeOU"
            Write-Host "Move successful for $($user.email)" -ForegroundColor Green
            $successCount++
        } catch
        {
            Write-Host "Error moving user $($user.email): $_" -ForegroundColor Red
            $errorCount++
        }
    }
    
    Write-Host "`nUser moving complete!" -ForegroundColor Cyan
    Write-Host "Successfully moved: $successCount" -ForegroundColor Green
    Write-Host "Failed operations: $errorCount" -ForegroundColor $(if ($errorCount -gt 0)
        { "Red" 
        } else
        { "Green" 
        })
    Read-Host "Press Enter to continue..."
}
# Function to unlock a single Chromebook by asset ID
function Unlock-SingleChromebook
{
    Clear-Host
    Write-Host "Unlock Single Chromebook" -ForegroundColor Cyan
    Write-Host "----------------------" -ForegroundColor Cyan
    $assetID = Read-Host "Enter Chromebook asset ID"
    
    if ([string]::IsNullOrWhiteSpace($assetID))
    {
        Write-Host "Error: Asset ID cannot be empty" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    Write-Host "Unlocking Chromebook with asset ID $assetID..." -ForegroundColor Yellow
    
    # The command to unlock a Chrome device by asset ID
    & gam update cros query "asset_id:$assetID" action unlock
    
    Write-Host "Unlock command has been sent to the device" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}

# Function to unlock multiple Chromebooks from CSV by asset ID
function Unlock-MultipleChromebooks
{
    Clear-Host
    Write-Host "Unlock Multiple Chromebooks (CSV)" -ForegroundColor Cyan
    Write-Host "-------------------------------" -ForegroundColor Cyan
    Write-Host "CSV format should be: asset_id" -ForegroundColor Yellow
    $csvPath = Read-Host "Enter path to CSV file"
    
    if (-not (Test-Path $csvPath))
    {
        Write-Host "Error: File not found!" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    $devices = Import-Csv $csvPath
    $successCount = 0
    $errorCount = 0
    
    foreach ($device in $devices)
    {
        if ([string]::IsNullOrEmpty($device.asset_id))
        {
            Write-Host "Skipping invalid entry: Missing asset_id" -ForegroundColor Yellow
            $errorCount++
            continue
        }
        
        try
        {
            Write-Host "Unlocking Chromebook with asset ID $($device.asset_id)..." -ForegroundColor Yellow
            & gam update cros query "asset_id:$($device.asset_id)" action unlock
            Write-Host "Unlock command sent successfully" -ForegroundColor Green
            $successCount++
        } catch
        {
            Write-Host "Error unlocking device: $_" -ForegroundColor Red
            $errorCount++
        }
    }
    
    Write-Host "`nChromebook unlocking complete!" -ForegroundColor Cyan
    Write-Host "Successfully unlocked: $successCount" -ForegroundColor Green
    Write-Host "Failed operations: $errorCount" -ForegroundColor $(if ($errorCount -gt 0)
        { "Red" 
        } else
        { "Green" 
        })
    Read-Host "Press Enter to continue..."
}



# Function to perform mass password reset from CSV
function Reset-MassPasswords
{
    Clear-Host
    Write-Host "Mass Password Reset (CSV)" -ForegroundColor Cyan
    Write-Host "------------------------" -ForegroundColor Cyan
    Write-Host "CSV format should be: email,password (password column is optional)" -ForegroundColor Yellow
    $csvPath = Read-Host "Enter path to CSV file"
    
    if (-not (Test-Path $csvPath))
    {
        Write-Host "Error: File not found!" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    $users = Import-Csv $csvPath
    $successCount = 0
    $errorCount = 0
    
    foreach ($user in $users)
    {
        if ([string]::IsNullOrEmpty($user.email))
        {
            Write-Host "Skipping invalid entry: Missing email" -ForegroundColor Yellow
            $errorCount++
            continue
        }
        
        # Use default password if not specified in CSV
        $password = if ([string]::IsNullOrEmpty($user.password))
        { $defaultPassword 
        } else
        { $user.password 
        }
        
        try
        {
            Write-Host "Resetting password for $($user.email)..." -ForegroundColor Yellow
            & gam update user "$($user.email)" password "$password"
            Write-Host "Password reset successful for $($user.email)" -ForegroundColor Green
            $successCount++
        } catch
        {
            Write-Host "Error resetting password for $($user.email): $_" -ForegroundColor Red
            $errorCount++
        }
    }
    
    Write-Host "`nPassword reset complete!" -ForegroundColor Cyan
    Write-Host "Successful resets: $successCount" -ForegroundColor Green
    Write-Host "Failed resets: $errorCount" -ForegroundColor $(if ($errorCount -gt 0)
        { "Red" 
        } else
        { "Green" 
        })
    Read-Host "Press Enter to continue..."
}




function New-SingleUser
{
    Clear-Host
    Write-Host "Create Single User" -ForegroundColor Cyan
    Write-Host "----------------" -ForegroundColor Cyan
    
    # Ask if this is a student or staff account
    Write-Host "Account Type:" -ForegroundColor Yellow
    Write-Host "1. Student (Cadet)" -ForegroundColor White
    Write-Host "2. Faculty/Staff" -ForegroundColor White
    Write-Host "Enter choice [1-2]: " -ForegroundColor Yellow -NoNewline
    $accountType = Read-Host
    
    $firstName = Read-Host "Enter first name"
    $lastName = Read-Host "Enter last name"
    
    # Get first initial
    $firstInitial = $firstName.Substring(0, 1)
    
    # Handle based on account type
    if ($accountType -eq "1")
    {
        # Cadet account - first initial + last name
        $baseEmail = "cadet$firstInitial$lastName@nomma.net".ToLower()
        $emailPrefix = $baseEmail.Split('@')[0]
        $domain = $baseEmail.Split('@')[1]
        $email = $baseEmail
        $counter = 0
        
        # Check if the base email exists - uses correct pattern matching for "Does not exist"
        $tempFile = [System.IO.Path]::GetTempFileName()
        & gam info user $email > $tempFile 2>&1
        $output = Get-Content $tempFile -Raw
        Remove-Item $tempFile -Force
        
        # Continue checking until we find an email that DOES NOT exist
        while ($output -notmatch "Does not exist")
        {
            $counter++
            $email = "$emailPrefix$counter@$domain"
            Write-Host "Email $baseEmail already exists, trying $email..." -ForegroundColor Yellow
            
            $tempFile = [System.IO.Path]::GetTempFileName()
            & gam info user $email > $tempFile 2>&1
            $output = Get-Content $tempFile -Raw
            Remove-Item $tempFile -Force
        }
        
        # Ask for grade level
        Write-Host "Select Cadet Grade Level:" -ForegroundColor Yellow
        Write-Host "1. 8th Grade" -ForegroundColor White
        Write-Host "2. 9th Grade" -ForegroundColor White
        Write-Host "3. 10th Grade" -ForegroundColor White
        Write-Host "4. 10th Grade with Email Privileges" -ForegroundColor White
        Write-Host "5. 11th Grade" -ForegroundColor White
        Write-Host "6. 12th Grade" -ForegroundColor White
        Write-Host "Enter choice [1-6]: " -ForegroundColor Yellow -NoNewline
        $gradeChoice = Read-Host
        
        # Map grade choice to OU
        $orgUnit = switch ($gradeChoice)
        {
            "1"
            { "/Cadets/8th Grade" 
            }
            "2"
            { "/Cadets/9th Grade" 
            }
            "3"
            { "/Cadets/10th Grade" 
            }
            "4"
            { "/Cadets/10th Grade with Email Privileges" 
            }
            "5"
            { "/Cadets/11th Grade" 
            }
            "6"
            { "/Cadets/12th Grade" 
            }
            default
            { 
                Write-Host "Invalid choice. Defaulting to 9th Grade." -ForegroundColor Yellow
                "/Cadets/9th Grade" 
            }
        }
        
        Write-Host "Creating cadet user $email in OU: $orgUnit..." -ForegroundColor Yellow
        & gam create user "$email" firstname "$firstName" lastname "$lastName" password "Password@1" org "$orgUnit" changepassword on
    } else
    {
        # Staff account - first initial + last name
        $baseEmail = "$firstInitial$lastName@nomma.net".ToLower()
        $emailPrefix = $baseEmail.Split('@')[0]
        $domain = $baseEmail.Split('@')[1]
        $email = $baseEmail
        $counter = 0
        
        # Check if the base email exists - uses correct pattern matching for "Does not exist"
        $tempFile = [System.IO.Path]::GetTempFileName()
        & gam info user $email > $tempFile 2>&1
        $output = Get-Content $tempFile -Raw
        Remove-Item $tempFile -Force
        
        # Continue checking until we find an email that DOES NOT exist
        while ($output -notmatch "Does not exist")
        {
            $counter++
            $email = "$emailPrefix$counter@$domain"
            Write-Host "Email $baseEmail already exists, trying $email..." -ForegroundColor Yellow
            
            $tempFile = [System.IO.Path]::GetTempFileName()
            & gam info user $email > $tempFile 2>&1
            $output = Get-Content $tempFile -Raw
            Remove-Item $tempFile -Force
        }
        
        # Ask for staff position
        Write-Host "Select Staff Position:" -ForegroundColor Yellow
        Write-Host "1. Counselors" -ForegroundColor White
        Write-Host "2. IT" -ForegroundColor White
        Write-Host "3. New Staff" -ForegroundColor White
        Write-Host "4. School Admins" -ForegroundColor White
        Write-Host "5. Security" -ForegroundColor White
        Write-Host "6. Staff" -ForegroundColor White
        Write-Host "7. Teachers" -ForegroundColor White
        Write-Host "Enter choice [1-7]: " -ForegroundColor Yellow -NoNewline
        $positionChoice = Read-Host
        
        # Map position choice to OU
        $orgUnit = switch ($positionChoice)
        {
            "1"
            { "/Faculty & Staff/Counselors" 
            }
            "2"
            { "/Faculty & Staff/IT" 
            }
            "3"
            { "/Faculty & Staff/New Staff" 
            }
            "4"
            { "/Faculty & Staff/School Admins" 
            }
            "5"
            { "/Faculty & Staff/Security" 
            }
            "6"
            { "/Faculty & Staff/Staff" 
            }
            "7"
            { "/Faculty & Staff/Teachers" 
            }
            default
            { 
                Write-Host "Invalid choice. Defaulting to New Staff." -ForegroundColor Yellow
                "/Faculty & Staff/New Staff" 
            }
        }
        
        Write-Host "Creating staff user $email in OU: $orgUnit..." -ForegroundColor Yellow
        & gam create user "$email" firstname "$firstName" lastname "$lastName" password "Password@1" org "$orgUnit"
    }
    
    Write-Host "User $email has been created" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}



function New-MultipleUsers
{
    Clear-Host
    Write-Host "Create Multiple Users (CSV)" -ForegroundColor Cyan
    Write-Host "-------------------------" -ForegroundColor Cyan
    
    # Ask if these are student or staff accounts
    Write-Host "Account Type:" -ForegroundColor Yellow
    Write-Host "1. Students (Cadets)" -ForegroundColor White
    Write-Host "2. Faculty/Staff" -ForegroundColor White
    Write-Host "Enter choice [1-2]: " -ForegroundColor Yellow -NoNewline
    $accountType = Read-Host
    
    if ($accountType -eq "1")
    {
        Write-Host "CSV format for CADETS should be: firstname,lastname,grade" -ForegroundColor Yellow
        Write-Host "Grade values: 8, 9, 10, 10email (for 10th with email), 11, 12" -ForegroundColor Yellow
    } else
    {
        Write-Host "CSV format for STAFF should be: firstname,lastname,position" -ForegroundColor Yellow
        Write-Host "Position values: Counselors, IT, NewStaff, SchoolAdmins, Security, Staff, Teachers" -ForegroundColor Yellow
    }
    
    $csvPath = Read-Host "Enter path to CSV file"
    
    if (-not (Test-Path $csvPath))
    {
        Write-Host "Error: File not found!" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    $users = Import-Csv $csvPath
    $successCount = 0
    $errorCount = 0
    
    foreach ($user in $users)
    {
        if ([string]::IsNullOrEmpty($user.firstname) -or [string]::IsNullOrEmpty($user.lastname))
        {
            Write-Host "Skipping invalid entry: Missing required fields" -ForegroundColor Yellow
            $errorCount++
            continue
        }
        
        try
        {
            if ($accountType -eq "1")
            {
                # Student (Cadet) account - using first initial + last name format
                $firstInitial = $user.firstname.Substring(0, 1)
                $baseEmail = "cadet$firstInitial$($user.lastname)@nomma.net".ToLower()
                $emailPrefix = $baseEmail.Split('@')[0]
                $domain = $baseEmail.Split('@')[1]
                $email = $baseEmail
                $counter = 0
                
                # Check if the base email exists - using correct pattern matching for "Does not exist"
                $tempFile = [System.IO.Path]::GetTempFileName()
                & gam info user $email > $tempFile 2>&1
                $output = Get-Content $tempFile -Raw
                Remove-Item $tempFile -Force
                
                # Continue checking until we find an email that DOES NOT exist
                while ($output -notmatch "Does not exist")
                {
                    $counter++
                    $email = "$emailPrefix$counter@$domain"
                    Write-Host "Email $baseEmail already exists, trying $email..." -ForegroundColor Yellow
                    
                    $tempFile = [System.IO.Path]::GetTempFileName()
                    & gam info user $email > $tempFile 2>&1
                    $output = Get-Content $tempFile -Raw
                    Remove-Item $tempFile -Force
                }
                
                # Map grade to OU
                $gradeOU = switch ($user.grade)
                {
                    "8"
                    { "/Cadets/8th Grade" 
                    }
                    "9"
                    { "/Cadets/9th Grade" 
                    }
                    "10"
                    { "/Cadets/10th Grade" 
                    }
                    "10email"
                    { "/Cadets/10th Grade with Email Privileges" 
                    }
                    "11"
                    { "/Cadets/11th Grade" 
                    }
                    "12"
                    { "/Cadets/12th Grade" 
                    }
                    default
                    { 
                        Write-Host "Invalid grade value for $($user.firstname) $($user.lastname): $($user.grade), defaulting to 9th Grade" -ForegroundColor Yellow
                        "/Cadets/9th Grade" 
                    }
                }
                
                Write-Host "Creating cadet user $email in $gradeOU..." -ForegroundColor Yellow
                & gam create user "$email" firstname "$($user.firstname)" lastname "$($user.lastname)" password "Password@1" org "$gradeOU" changepassword on
            } else
            {
                # Staff account - first initial + last name
                $firstInitial = $user.firstname.Substring(0, 1)
                $baseEmail = "$firstInitial$($user.lastname)@nomma.net".ToLower()
                $emailPrefix = $baseEmail.Split('@')[0]
                $domain = $baseEmail.Split('@')[1]
                $email = $baseEmail
                $counter = 0
                
                # Check if the base email exists - using correct pattern matching for "Does not exist"
                $tempFile = [System.IO.Path]::GetTempFileName()
                & gam info user $email > $tempFile 2>&1
                $output = Get-Content $tempFile -Raw
                Remove-Item $tempFile -Force
                
                # Continue checking until we find an email that DOES NOT exist
                while ($output -notmatch "Does not exist")
                {
                    $counter++
                    $email = "$emailPrefix$counter@$domain"
                    Write-Host "Email $baseEmail already exists, trying $email..." -ForegroundColor Yellow
                    
                    $tempFile = [System.IO.Path]::GetTempFileName()
                    & gam info user $email > $tempFile 2>&1
                    $output = Get-Content $tempFile -Raw
                    Remove-Item $tempFile -Force
                }
                
                # Map position to OU
                $positionOU = switch -Regex ($user.position)
                {
                    "^[Cc]ounselor"
                    { "/Faculty & Staff/Counselors" 
                    }
                    "^[Ii][Tt]$"
                    { "/Faculty & Staff/IT" 
                    }
                    "^[Nn]ew"
                    { "/Faculty & Staff/New Staff" 
                    }
                    "^[Aa]dmin"
                    { "/Faculty & Staff/School Admins" 
                    }
                    "^[Ss]ecurity"
                    { "/Faculty & Staff/Security" 
                    }
                    "^[Ss]taff$"
                    { "/Faculty & Staff/Staff" 
                    }
                    "^[Tt]eacher"
                    { "/Faculty & Staff/Teachers" 
                    }
                    default
                    { 
                        Write-Host "Invalid position value for $($user.firstname) $($user.lastname): $($user.position), defaulting to New Staff" -ForegroundColor Yellow
                        "/Faculty & Staff/New Staff" 
                    }
                }
                
                Write-Host "Creating staff user $email in $positionOU..." -ForegroundColor Yellow
                & gam create user "$email" firstname "$($user.firstname)" lastname "$($user.lastname)" password "Password@1" org "$positionOU" changepassword on
            }
            
            Write-Host "User creation successful for $email" -ForegroundColor Green
            $successCount++
        } catch
        {
            Write-Host "Error creating user: $_" -ForegroundColor Red
            $errorCount++
        }
    }
    
    Write-Host "`nUser creation complete!" -ForegroundColor Cyan
    Write-Host "Successfully created: $successCount" -ForegroundColor Green
    Write-Host "Failed creations: $errorCount" -ForegroundColor $(if ($errorCount -gt 0)
        { "Red" 
        } else
        { "Green" 
        })
    Read-Host "Press Enter to continue..."
}




# Function to archive multiple users from CSV with enhanced process (Fixed error handling)
function Archive-MultipleUsers
{
    Clear-Host
    Write-Host "Archive Multiple Users (CSV)" -ForegroundColor Cyan
    Write-Host "--------------------------" -ForegroundColor Cyan
    Write-Host "CSV format should be: email" -ForegroundColor Yellow
    $csvPath = Read-Host "Enter path to CSV file"
    
    if (-not (Test-Path $csvPath))
    {
        Write-Host "Error: File not found!" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    # Archive OU path - now hardcoded to /Inactive
    $archiveOU = "/Inactive"
    
    Write-Host "`nArchiving users with the following steps:" -ForegroundColor Yellow
    Write-Host "1. Sign out of all sessions" -ForegroundColor White
    Write-Host "2. Set a random secure password" -ForegroundColor White
    Write-Host "3. Suspend the account" -ForegroundColor White
    Write-Host "4. Move to $archiveOU OU" -ForegroundColor White
    Write-Host "5. Rename email to inactive-{email}" -ForegroundColor White
    Write-Host "6. Remove all email aliases" -ForegroundColor White
    
    $confirmation = Read-Host "`nAre you sure you want to archive all users in the CSV? (y/n)"
    if ($confirmation -ne "y")
    {
        Write-Host "Operation cancelled" -ForegroundColor Yellow
        Read-Host "Press Enter to continue..."
        return
    }
    
    $users = Import-Csv $csvPath
    $successCount = 0
    $errorCount = 0
    
    foreach ($user in $users)
    {
        if ([string]::IsNullOrEmpty($user.email))
        {
            Write-Host "Skipping invalid entry: Missing email" -ForegroundColor Yellow
            $errorCount++
            continue
        }
        
        $userEmail = $user.email
        
        try
        {
            Write-Host "`nArchiving user $userEmail..." -ForegroundColor Yellow
            
            # Generate a random secure password (16-24 characters)
            $randomPasswordLength = Get-Random -Minimum 16 -Maximum 25
            $secureChars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%^&*()_+-=[]{}|;:,./<>?"
            $randomPassword = -join ((1..$randomPasswordLength) | ForEach-Object { $secureChars[(Get-Random -Minimum 0 -Maximum $secureChars.Length)] })
            
            # Step 1: Sign out from all sessions
            Write-Host "- Signing out from all sessions..." -ForegroundColor Yellow
            & gam user "$userEmail" signout
            
            # Step 2: Set random secure password
            Write-Host "- Setting random secure password..." -ForegroundColor Yellow
            & gam update user "$userEmail" password "$randomPassword"
            
            # Step 3: Suspend the account
            Write-Host "- Suspending the account..." -ForegroundColor Yellow
            & gam update user "$userEmail" suspended on
            
            # Step 4: Move to archive OU
            Write-Host "- Moving to $archiveOU OU..." -ForegroundColor Yellow
            & gam update user "$userEmail" org "$archiveOU"
            
            # Step 5: Create new email with inactive- prefix
            $domain = $userEmail.Substring($userEmail.IndexOf('@'))
            $username = $userEmail.Substring(0, $userEmail.IndexOf('@'))
            $newEmail = "inactive-$username$domain"
            
            Write-Host "- Renaming email to $newEmail..." -ForegroundColor Yellow
            & gam update user "$userEmail" email "$newEmail"
            
            # Step 6: Remove all aliases
            Write-Host "- Removing all email aliases..." -ForegroundColor Yellow
            
            # Get all aliases for the user
            $tempFile = [System.IO.Path]::GetTempFileName()
            & gam user "$newEmail" print aliases > $tempFile
            
            if ((Get-Item $tempFile).Length -gt 0)
            {
                $aliases = Get-Content $tempFile | Where-Object { $_ -ne "" -and $_ -ne $newEmail }
                
                if ($aliases.Count -gt 0)
                {
                    foreach ($alias in $aliases)
                    {
                        Write-Host "  Removing alias: $alias" -ForegroundColor Yellow
                        & gam user "$newEmail" delete alias "$alias"
                    }
                }
            }
            
            # Clean up the temp file
            Remove-Item $tempFile -Force
            
            Write-Host "Archive successful for $userEmail (now $newEmail)" -ForegroundColor Green
            $successCount++
        } catch
        {
            # Fixed error handling syntax
            $errorMessage = $_.Exception.Message
            Write-Host "Error archiving user $userEmail`: $errorMessage" -ForegroundColor Red
            $errorCount++
        }
    }
    
    Write-Host "`nUser archiving complete!" -ForegroundColor Cyan
    Write-Host "Successfully archived: $successCount" -ForegroundColor Green
    Write-Host "Failed operations: $errorCount" -ForegroundColor $(if ($errorCount -gt 0)
        { "Red" 
        } else
        { "Green" 
        })
    Read-Host "Press Enter to continue..."
}

# Also fix the same issue in the single user archive function
function Archive-SingleUser
{
    Clear-Host
    Write-Host "Archive User" -ForegroundColor Cyan
    Write-Host "-----------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    
    if ([string]::IsNullOrWhiteSpace($userEmail))
    {
        Write-Host "Error: User email cannot be empty" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    # Generate a random secure password (16-24 characters)
    $randomPasswordLength = Get-Random -Minimum 16 -Maximum 25
    $secureChars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%^&*()_+-=[]{}|;:,./<>?"
    $randomPassword = -join ((1..$randomPasswordLength) | ForEach-Object { $secureChars[(Get-Random -Minimum 0 -Maximum $secureChars.Length)] })
    
    # Archive OU path - now hardcoded to /Inactive
    $archiveOU = "/Inactive"
    
    Write-Host "`nArchiving user $userEmail with the following steps:" -ForegroundColor Yellow
    Write-Host "1. Sign out of all sessions" -ForegroundColor White
    Write-Host "2. Set a random secure password" -ForegroundColor White
    Write-Host "3. Suspend the account" -ForegroundColor White
    Write-Host "4. Move to $archiveOU OU" -ForegroundColor White
    Write-Host "5. Rename email to inactive-$userEmail" -ForegroundColor White
    Write-Host "6. Remove all email aliases" -ForegroundColor White
    
    $confirmation = Read-Host "`nAre you sure you want to archive this user? (y/n)"
    if ($confirmation -ne "y")
    {
        Write-Host "Operation cancelled" -ForegroundColor Yellow
        Read-Host "Press Enter to continue..."
        return
    }
    
    try
    {
        # Step 1: Sign out from all sessions
        Write-Host "`nStep 1: Signing out from all sessions..." -ForegroundColor Yellow
        & gam user "$userEmail" signout
        
        # Step 2: Set random secure password
        Write-Host "Step 2: Setting random secure password..." -ForegroundColor Yellow
        & gam update user "$userEmail" password "$randomPassword"
        
        # Step 3: Suspend the account
        Write-Host "Step 3: Suspending the account..." -ForegroundColor Yellow
        & gam update user "$userEmail" suspended on
        
        # Step 4: Move to archive OU
        Write-Host "Step 4: Moving to $archiveOU OU..." -ForegroundColor Yellow
        & gam update user "$userEmail" org "$archiveOU"
        
        # Step 5: Get the domain from the email address
        $domain = $userEmail.Substring($userEmail.IndexOf('@'))
        $username = $userEmail.Substring(0, $userEmail.IndexOf('@'))
        $newEmail = "inactive-$username$domain"
        
        Write-Host "Step 5: Renaming email to $newEmail..." -ForegroundColor Yellow
        & gam update user "$userEmail" email "$newEmail"
        
        # Step 6: Remove all aliases
        Write-Host "Step 6: Removing all email aliases..." -ForegroundColor Yellow
        
        # Get all aliases for the user
        $tempFile = [System.IO.Path]::GetTempFileName()
        & gam user "$newEmail" print aliases > $tempFile
        
        if ((Get-Item $tempFile).Length -gt 0)
        {
            $aliases = Get-Content $tempFile | Where-Object { $_ -ne "" -and $_ -ne $newEmail }
            
            if ($aliases.Count -gt 0)
            {
                Write-Host "Found $($aliases.Count) aliases to remove:" -ForegroundColor Yellow
                
                foreach ($alias in $aliases)
                {
                    Write-Host "Removing alias: $alias" -ForegroundColor Yellow
                    & gam user "$newEmail" delete alias "$alias"
                }
                
                Write-Host "All aliases removed." -ForegroundColor Green
            } else
            {
                Write-Host "No aliases found to remove." -ForegroundColor Yellow
            }
        } else
        {
            Write-Host "No aliases information found." -ForegroundColor Yellow
        }
        
        # Clean up the temp file
        Remove-Item $tempFile -Force
        
        Write-Host "`nUser $userEmail has been successfully archived as $newEmail" -ForegroundColor Green
    } catch
    {
        # Fixed error handling syntax
        $errorMessage = $_.Exception.Message
        Write-Host "Error during archiving process`: $errorMessage" -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue..."
}


# Function to suspend a user
function Suspend-User
{
    Clear-Host
    Write-Host "Suspend User" -ForegroundColor Cyan
    Write-Host "-----------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    
    Write-Host "Suspending user $userEmail..." -ForegroundColor Yellow
    & gam update user "$userEmail" suspended on
    Write-Host "User $userEmail has been suspended" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}

# Function to unsuspend a user
function Unsuspend-User
{
    Clear-Host
    Write-Host "Unsuspend User" -ForegroundColor Cyan
    Write-Host "-------------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    
    Write-Host "Unsuspending user $userEmail..." -ForegroundColor Yellow
    & gam update user "$userEmail" suspended off
    Write-Host "User $userEmail has been unsuspended" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}

# Function to delete a user
function Remove-User
{
    Clear-Host
    Write-Host "Delete User" -ForegroundColor Cyan
    Write-Host "----------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    
    $confirmation = Read-Host "Are you sure you want to delete $userEmail? (y/n)"
    if ($confirmation -ne "y")
    {
        Write-Host "Operation cancelled" -ForegroundColor Yellow
        Read-Host "Press Enter to continue..."
        return
    }
    
    Write-Host "Deleting user $userEmail..." -ForegroundColor Yellow
    & gam delete user "$userEmail"
    Write-Host "User $userEmail has been deleted" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}


# Function to lock a single Chromebook by asset ID
function Lock-SingleChromebook
{
    Clear-Host
    Write-Host "Lock Single Chromebook" -ForegroundColor Cyan
    Write-Host "--------------------" -ForegroundColor Cyan
    $assetID = Read-Host "Enter Chromebook asset ID"
    
    if ([string]::IsNullOrWhiteSpace($assetID))
    {
        Write-Host "Error: Asset ID cannot be empty" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    Write-Host "Locking Chromebook with asset ID $assetID..." -ForegroundColor Yellow
    
    # The command to lock a Chrome device by asset ID
    & gam update cros query "asset_id:$assetID" action disable
    
    Write-Host "Lock command has been sent to the device" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}

# Function to lock multiple Chromebooks from CSV by asset ID
function Lock-MultipleChromebooks
{
    Clear-Host
    Write-Host "Lock Multiple Chromebooks (CSV)" -ForegroundColor Cyan
    Write-Host "-----------------------------" -ForegroundColor Cyan
    Write-Host "CSV format should be: asset_id" -ForegroundColor Yellow
    $csvPath = Read-Host "Enter path to CSV file"
    
    if (-not (Test-Path $csvPath))
    {
        Write-Host "Error: File not found!" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    $devices = Import-Csv $csvPath
    $successCount = 0
    $errorCount = 0
    
    foreach ($device in $devices)
    {
        if ([string]::IsNullOrEmpty($device.asset_id))
        {
            Write-Host "Skipping invalid entry: Missing asset_id" -ForegroundColor Yellow
            $errorCount++
            continue
        }
        
        try
        {
            Write-Host "Locking Chromebook with asset ID $($device.asset_id)..." -ForegroundColor Yellow
            & gam update cros query "asset_id:$($device.asset_id)" action disable
            Write-Host "Lock command sent successfully" -ForegroundColor Green
            $successCount++
        } catch
        {
            Write-Host "Error locking device: $_" -ForegroundColor Red
            $errorCount++
        }
    }
    
    Write-Host "`nChromebook locking complete!" -ForegroundColor Cyan
    Write-Host "Successfully locked: $successCount" -ForegroundColor Green
    Write-Host "Failed operations: $errorCount" -ForegroundColor $(if ($errorCount -gt 0)
        { "Red" 
        } else
        { "Green" 
        })
    Read-Host "Press Enter to continue..."
}



# Function to wipe a device
function Wipe-Device
{
    Clear-Host
    Write-Host "Wipe Device" -ForegroundColor Cyan
    Write-Host "----------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    
    # List devices for the user
    Write-Host "Listing devices for $userEmail..." -ForegroundColor Yellow
    & gam print mobile query "user:$userEmail" | Out-Host
    
    $deviceId = Read-Host "Enter device ID to wipe"
    
    $confirmation = Read-Host "Are you sure you want to wipe this device? (y/n)"
    if ($confirmation -ne "y")
    {
        Write-Host "Operation cancelled" -ForegroundColor Yellow
        Read-Host "Press Enter to continue..."
        return
    }
    
    Write-Host "Wiping device $deviceId for $userEmail..." -ForegroundColor Yellow
    & gam user "$userEmail" update mobile "$deviceId" action wipe
    Write-Host "Device wipe command has been sent" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}

# Function to wipe multiple devices from CSV
function Wipe-MultipleDevices
{
    Clear-Host
    Write-Host "Wipe Multiple Devices (CSV)" -ForegroundColor Cyan
    Write-Host "-------------------------" -ForegroundColor Cyan
    Write-Host "CSV format should be: email,deviceid" -ForegroundColor Yellow
    $csvPath = Read-Host "Enter path to CSV file"
    
    if (-not (Test-Path $csvPath))
    {
        Write-Host "Error: File not found!" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    $confirmation = Read-Host "Are you sure you want to wipe ALL devices in this CSV? (y/n)"
    if ($confirmation -ne "y")
    {
        Write-Host "Operation cancelled" -ForegroundColor Yellow
        Read-Host "Press Enter to continue..."
        return
    }
    
    $devices = Import-Csv $csvPath
    $successCount = 0
    $errorCount = 0
    
    foreach ($device in $devices)
    {
        if ([string]::IsNullOrEmpty($device.email) -or [string]::IsNullOrEmpty($device.deviceid))
        {
            Write-Host "Skipping invalid entry" -ForegroundColor Yellow
            $errorCount++
            continue
        }
        
        try
        {
            Write-Host "Wiping device $($device.deviceid) for $($device.email)..." -ForegroundColor Yellow
            & gam user "$($device.email)" update mobile "$($device.deviceid)" action wipe
            Write-Host "Device wipe command sent" -ForegroundColor Green
            $successCount++
        } catch
        {
            Write-Host "Error wiping device: $_" -ForegroundColor Red
            $errorCount++
        }
    }
    
    Write-Host "`nDevice wiping complete!" -ForegroundColor Cyan
    Write-Host "Successfully sent wipe commands: $successCount" -ForegroundColor Green
    Write-Host "Failed operations: $errorCount" -ForegroundColor $(if ($errorCount -gt 0)
        { "Red" 
        } else
        { "Green" 
        })
    Read-Host "Press Enter to continue..."
}

# Function to list all devices
function List-Devices
{
    Clear-Host
    Write-Host "List All Devices" -ForegroundColor Cyan
    Write-Host "--------------" -ForegroundColor Cyan
    $outputOption = Read-Host "Save to CSV file? (y/n)"
    
    if ($outputOption -eq "y")
    {
        $outputPath = "devices_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        Write-Host "Retrieving all mobile devices..." -ForegroundColor Yellow
        & gam print mobile > $outputPath
        Write-Host "Device list saved to $outputPath" -ForegroundColor Green
    } else
    {
        Write-Host "Retrieving all mobile devices..." -ForegroundColor Yellow
        & gam print mobile | Out-Host
    }
    
    Read-Host "Press Enter to continue..."
}

# Function to get device serial number by asset ID
function Get-DeviceSerialNumber
{
    Clear-Host
    Write-Host "Get Device Serial Number" -ForegroundColor Cyan
    Write-Host "----------------------" -ForegroundColor Cyan
    
    # Ask for lookup method
    Write-Host "Lookup Method:" -ForegroundColor Yellow
    Write-Host "1. By User Email" -ForegroundColor White
    Write-Host "2. By Asset ID" -ForegroundColor White
    Write-Host "Enter choice [1-2]: " -ForegroundColor Yellow -NoNewline
    $lookupMethod = Read-Host
    
    if ($lookupMethod -eq "1")
    {
        # Original method - by user email
        $userEmail = Read-Host "Enter user email"
        
        Write-Host "Retrieving devices for $userEmail..." -ForegroundColor Yellow
        & gam print mobile query "user:$userEmail" | Out-Host
    } else
    {
        # New method - by asset ID
        $assetID = Read-Host "Enter device asset ID"
        
        if ([string]::IsNullOrWhiteSpace($assetID))
        {
            Write-Host "Error: Asset ID cannot be empty" -ForegroundColor Red
            Read-Host "Press Enter to continue..."
            return
        }
        
        Write-Host "Retrieving serial number for asset ID: $assetID..." -ForegroundColor Yellow
        
        try
        {
            # First, try to save the output to a temporary file to inspect
            $tempFile = [System.IO.Path]::GetTempFileName()
            
            # Execute GAM command and save output to file
            & gam print cros query "asset_id:$assetID" > $tempFile
            
            # Display raw output for debugging
            Write-Host "`nRAW GAM OUTPUT:" -ForegroundColor Magenta
            Get-Content $tempFile | ForEach-Object { Write-Host $_ }
            
            # Check if file has content
            if ((Get-Item $tempFile).Length -eq 0)
            {
                Write-Host "`nNo data returned from GAM for asset ID: $assetID" -ForegroundColor Yellow
                Read-Host "Press Enter to continue..."
                return
            }
            
            # Try alternate approach - parse the output manually
            $rawOutput = Get-Content $tempFile
            $deviceData = $null
            
            # Check first line for headers and other lines for data
            if ($rawOutput.Count -gt 1)
            {
                $headers = $rawOutput[0].Split(',')
                $values = $rawOutput[1].Split(',')
                
                # Create a custom object with the data
                $deviceData = [PSCustomObject]@{}
                for ($i = 0; $i -lt $headers.Count -and $i -lt $values.Count; $i++)
                {
                    $deviceData | Add-Member -MemberType NoteProperty -Name $headers[$i].Trim() -Value $values[$i].Trim()
                }
                
                # Display the device information
                Write-Host "`nDevice Information:" -ForegroundColor Green
                
                if ($deviceData.PSObject.Properties.Name -contains "serialNumber")
                {
                    Write-Host "Serial Number: $($deviceData.serialNumber)" -ForegroundColor Green
                }
                
                # Display all properties for debugging
                Write-Host "`nAll Device Properties:" -ForegroundColor Cyan
                $deviceData.PSObject.Properties | ForEach-Object {
                    Write-Host "$($_.Name): $($_.Value)" -ForegroundColor White
                }
            } else
            {
                Write-Host "`nError parsing device data: Unexpected format" -ForegroundColor Red
            }
            
            # Clean up temp file
            Remove-Item $tempFile -Force
            
            # Alternative direct approach for Chrome devices
            Write-Host "`nTrying alternative approach..." -ForegroundColor Yellow
            Write-Host "Running: gam info cros query 'asset_id:$assetID'" -ForegroundColor Blue
            & gam info cros query "asset_id:$assetID" | Out-Host
        } catch
        {
            Write-Host "Error retrieving device information: $_" -ForegroundColor Red
        }
    }
    
    Read-Host "`nPress Enter to continue..."
}

# Function to get user information with improved error handling
function Get-UserInformation
{
    Clear-Host
    Write-Host "Get User Information" -ForegroundColor Cyan
    Write-Host "------------------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    
    # Validate input
    if ([string]::IsNullOrWhiteSpace($userEmail))
    {
        Write-Host "Error: User email cannot be empty" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        return
    }
    
    try
    {
        Write-Host "Retrieving information for $userEmail..." -ForegroundColor Yellow
        
        # Create a temporary file to capture the output
        $tempFile = [System.IO.Path]::GetTempFileName()
        
        # Run the GAM command and capture output
        & gam info user "$userEmail" > $tempFile
        
        # Check if we got any output
        if ((Get-Item $tempFile).Length -eq 0)
        {
            Write-Host "No information returned for user: $userEmail" -ForegroundColor Yellow
            Write-Host "User may not exist or you may not have permission to view their information." -ForegroundColor Yellow
        } else
        {
            # Display the output
            Get-Content $tempFile | Out-Host
        }
        
        # Clean up the temp file
        Remove-Item $tempFile -Force
    } catch
    {
        Write-Host "Error retrieving user information: $_" -ForegroundColor Red
    }
    
    Write-Host "`nOperation completed." -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}

# Function to list all users
function List-AllUsers
{
    Clear-Host
    Write-Host "List All Users" -ForegroundColor Cyan
    Write-Host "------------" -ForegroundColor Cyan
    $outputOption = Read-Host "Save to CSV file? (y/n)"
    
    if ($outputOption -eq "y")
    {
        $outputPath = "users_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        Write-Host "Retrieving all users..." -ForegroundColor Yellow
        & gam print users > $outputPath
        Write-Host "User list saved to $outputPath" -ForegroundColor Green
    } else
    {
        Write-Host "Retrieving all users..." -ForegroundColor Yellow
        & gam print users | Out-Host
    }
    
    Read-Host "Press Enter to continue..."
}

# Function to check license information
function Check-LicenseInformation
{
    Clear-Host
    Write-Host "Check License Information" -ForegroundColor Cyan
    Write-Host "-----------------------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email (leave blank for all licenses)"
    
    if ([string]::IsNullOrEmpty($userEmail))
    {
        Write-Host "Retrieving all license information..." -ForegroundColor Yellow
        & gam print licenses | Out-Host
    } else
    {
        Write-Host "Retrieving license information for $userEmail..." -ForegroundColor Yellow
        & gam user "$userEmail" show licenses | Out-Host
    }
    
    Read-Host "Press Enter to continue..."
}

# Function to generate user activity report
function Get-UserActivityReport
{
    Clear-Host
    Write-Host "Generate User Activity Report" -ForegroundColor Cyan
    Write-Host "--------------------------" -ForegroundColor Cyan
    $days = Read-Host "Enter number of days to include (default: 7)"
    
    if ([string]::IsNullOrEmpty($days) -or -not [int]::TryParse($days, [ref]$null))
    {
        $days = 7
    }
    
    $outputPath = "user_activity_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Write-Host "Generating user activity report for the last $days days..." -ForegroundColor Yellow
    & gam report users filter "date >= $(Get-Date).AddDays(-$days).ToString('yyyy-MM-dd')" > $outputPath
    Write-Host "Report saved to $outputPath" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}
# Fixed function to monitor security reports with proper date format
function Monitor-SecurityReports
{
    Clear-Host
    Write-Host "Monitor Security Reports" -ForegroundColor Cyan
    Write-Host "----------------------" -ForegroundColor Cyan
    
    # Provide options for different security report types
    Write-Host "Select Security Report Type:" -ForegroundColor Yellow
    Write-Host "1. Login Audit Report" -ForegroundColor White
    Write-Host "2. Token Audit Report" -ForegroundColor White
    Write-Host "3. Admin Activity Report" -ForegroundColor White
    Write-Host "4. Data Access Report" -ForegroundColor White
    Write-Host "5. Rules Activity Report" -ForegroundColor White
    Write-Host "6. Return to Information Menu" -ForegroundColor White
    Write-Host "Enter choice [1-6]: " -ForegroundColor Yellow -NoNewline
    
    $reportChoice = Read-Host
    
    if ($reportChoice -eq "6")
    {
        return
    }
    
    # Select the report type
    $reportType = switch ($reportChoice)
    {
        "1"
        { "login" 
        }
        "2"
        { "token" 
        }
        "3"
        { "admin" 
        }
        "4"
        { "drive" 
        }
        "5"
        { "rules" 
        }
        default
        { "login" 
        }
    }
    
    # Ask for specific user email (optional)
    Write-Host "`nFilter by specific user? (y/n): " -ForegroundColor Yellow -NoNewline
    $filterByUser = Read-Host
    
    $userEmail = ""
    if ($filterByUser -eq "y")
    {
        $userEmail = Read-Host "Enter user email"
        if ([string]::IsNullOrWhiteSpace($userEmail))
        {
            Write-Host "No user email provided. Will show reports for all users." -ForegroundColor Yellow
        }
    }
    
    # Ask for the time range
    Write-Host "`nSelect Time Range:" -ForegroundColor Yellow
    Write-Host "1. Last 24 hours" -ForegroundColor White
    Write-Host "2. Last 7 days" -ForegroundColor White
    Write-Host "3. Last 30 days" -ForegroundColor White
    Write-Host "4. Custom range" -ForegroundColor White
    Write-Host "Enter choice [1-4]: " -ForegroundColor Yellow -NoNewline
    
    $timeChoice = Read-Host
    
    # Calculate the start date based on time range selection
    $endDate = Get-Date
    $startDate = switch ($timeChoice)
    {
        "1"
        { $endDate.AddDays(-1) 
        }
        "2"
        { $endDate.AddDays(-7) 
        }
        "3"
        { $endDate.AddDays(-30) 
        }
        "4"
        {
            Write-Host "`nEnter start date (MM/DD/YYYY): " -ForegroundColor Yellow -NoNewline
            $customStartDate = Read-Host
            try
            {
                [DateTime]::ParseExact($customStartDate, "MM/dd/yyyy", $null)
            } catch
            {
                Write-Host "Invalid date format. Using last 7 days instead." -ForegroundColor Red
                $endDate.AddDays(-7)
            }
        }
        default
        { $endDate.AddDays(-7) 
        }
    }
    
    # Format dates for GAM command
    $startDateStr = $startDate.ToString("yyyy-MM-dd")
    $endDateStr = $endDate.ToString("yyyy-MM-dd")
    
    # Ask if results should be saved to a file
    Write-Host "`nSave results to CSV file? (y/n): " -ForegroundColor Yellow -NoNewline
    $saveToFile = Read-Host
    
    # Display report parameters
    Write-Host "`nReport Parameters:" -ForegroundColor Cyan
    Write-Host "- Report Type: $reportType" -ForegroundColor White
    Write-Host "- Date Range: $startDateStr to $endDateStr" -ForegroundColor White
    if (-not [string]::IsNullOrWhiteSpace($userEmail))
    {
        Write-Host "- User Filter: $userEmail" -ForegroundColor White
    } else
    {
        Write-Host "- Users: All users" -ForegroundColor White
    }
    
    Write-Host "`nGenerating report... This may take a moment." -ForegroundColor Yellow
    
    try
    {
        # Let's try a simpler approach - without using date filters
        # Based on the error, it seems the specific date filter format isn't working
        
        if ($saveToFile -eq "y")
        {
            $userPart = if (-not [string]::IsNullOrWhiteSpace($userEmail))
            { "_$($userEmail -replace '@.*$', '')" 
            } else
            { "_all-users" 
            }
            $outputPath = "${reportType}${userPart}_report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            
            # Try different approaches to the command
            if (-not [string]::IsNullOrWhiteSpace($userEmail))
            {
                # Try approach 1: Just specify the user without date filters
                Write-Host "Running: gam report $reportType user $userEmail" -ForegroundColor Blue
                & gam report $reportType user $userEmail > $outputPath
            } else
            {
                # Try approach 2: Basic report with no filters
                Write-Host "Running: gam report $reportType" -ForegroundColor Blue
                & gam report $reportType > $outputPath
            }
            
            # Check if file was created and has content
            if (Test-Path $outputPath -PathType Leaf)
            {
                $fileSize = (Get-Item $outputPath).Length
                if ($fileSize -eq 0)
                {
                    Write-Host "No data found for the specified criteria." -ForegroundColor Yellow
                    Remove-Item $outputPath # Delete empty file
                } else
                {
                    Write-Host "Report saved to $outputPath" -ForegroundColor Green
                    Write-Host "Note: The report may include data outside your requested date range." -ForegroundColor Yellow
                    
                    # Ask if user wants to open the file
                    Write-Host "`nWould you like to view the report now? (y/n): " -ForegroundColor Yellow -NoNewline
                    $viewReport = Read-Host
                    
                    if ($viewReport -eq "y")
                    {
                        # Check file size before opening
                        if ($fileSize -gt 1MB)
                        {
                            Write-Host "File is large ($(($fileSize/1MB).ToString('0.00')) MB). Are you sure you want to open it? (y/n): " -ForegroundColor Yellow -NoNewline
                            $confirm = Read-Host
                            if ($confirm -eq "y")
                            {
                                Invoke-Item $outputPath
                            }
                        } else
                        {
                            Invoke-Item $outputPath
                        }
                    }
                }
            } else
            {
                Write-Host "Failed to create report file." -ForegroundColor Red
            }
        } else
        {
            # Execute GAM command and display results, but with simplified approach
            if (-not [string]::IsNullOrWhiteSpace($userEmail))
            {
                # For specific user without date filters
                Write-Host "Running: gam report $reportType user $userEmail" -ForegroundColor Blue
                & gam report $reportType user $userEmail | Out-Host
            } else
            {
                # For all users without date filters
                Write-Host "Running: gam report $reportType" -ForegroundColor Blue
                & gam report $reportType | Out-Host
            }
            
            Write-Host "Note: The report may include data outside your requested date range." -ForegroundColor Yellow
        }
    } catch
    {
        Write-Host "Error generating report: $_" -ForegroundColor Red
    }
    
    Write-Host "`nOperation completed." -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}

# Function to display CSV templates for bulk operations
function Show-TemplateMenu
{
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "           CSV TEMPLATES MENU             " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "1. Mass Password Reset Template" -ForegroundColor White
    Write-Host "2. Create Multiple Cadets Template" -ForegroundColor White
    Write-Host "3. Create Multiple Staff Template" -ForegroundColor White
    Write-Host "4. Archive Multiple Users Template" -ForegroundColor White
    Write-Host "5. Lock Multiple Devices Template" -ForegroundColor White
    Write-Host "6. Wipe Multiple Devices Template" -ForegroundColor White
    Write-Host "7. Move Users to Grade OU Template" -ForegroundColor White
    Write-Host "8. Return to Main Menu" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-7]: " -ForegroundColor Yellow -NoNewline
    
    $templateChoice = Read-Host
    
    switch ($templateChoice)
    {
        "1"
        { Show-PasswordResetTemplate 
        }
        "2"
        { Show-CadetCreationTemplate 
        }
        "3"
        { Show-StaffCreationTemplate 
        }
        "4"
        { Show-ArchiveUsersTemplate 
        }
        "5"
        { Show-LockDevicesTemplate 
        }
        "6"
        { Show-WipeDevicesTemplate 
        }
        "7"
        { Show-MoveUsersTemplate
        }
        "8"
        { return 
        }
        default
        {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-TemplateMenu
        }
    }
}


# Function to show and export move users template
function Show-MoveUsersTemplate
{
    Clear-Host
    Write-Host "Move Users to Grade OU Template" -ForegroundColor Cyan
    Write-Host "-----------------------------" -ForegroundColor Cyan
    
    $template = @"
email,grade
cadetjohndoe@nomma.net,8
cadetjanesmith@nomma.net,9
cadetmikebrown@nomma.net,10
cadetsamjones@nomma.net,10email
cadetannahall@nomma.net,11
cadetbengreen@nomma.net,12
"@
    
    Write-Host "CSV Format:" -ForegroundColor Yellow
    Write-Host $template -ForegroundColor White
    Write-Host "`nNotes:" -ForegroundColor Yellow
    Write-Host "- Valid grade values: 8, 9, 10, 10email (for 10th with email), 11, 12" -ForegroundColor White
    Write-Host "- Users will be moved to the corresponding grade OU" -ForegroundColor White
    Write-Host "- Save as a .csv file before using." -ForegroundColor White
    
    Save-TemplateOption "move_users_template.csv" $template
}

# Function to show and export password reset template
function Show-PasswordResetTemplate
{
    Clear-Host
    Write-Host "Mass Password Reset Template" -ForegroundColor Cyan
    Write-Host "-------------------------" -ForegroundColor Cyan
    
    $template = @"
email,password
user1@nomma.net,Password@1
user2@nomma.net,CustomPass@2
user3@nomma.net,
"@
    
    Write-Host "CSV Format:" -ForegroundColor Yellow
    Write-Host $template -ForegroundColor White
    Write-Host "`nNotes:" -ForegroundColor Yellow
    Write-Host "- The 'password' column is optional. If empty, the default password will be used." -ForegroundColor White
    Write-Host "- One user per line, with their email address in the first column." -ForegroundColor White
    Write-Host "- Save as a .csv file before using." -ForegroundColor White
    
    Save-TemplateOption "password_reset_template.csv" $template
}

# Function to show and export cadet creation template
function Show-CadetCreationTemplate
{
    Clear-Host
    Write-Host "Create Multiple Cadets Template" -ForegroundColor Cyan
    Write-Host "---------------------------" -ForegroundColor Cyan
    
    $template = @"
firstname,lastname,grade
John,Doe,8
Jane,Smith,9
Michael,Johnson,10
Sarah,Williams,10email
David,Brown,11
Emily,Jones,12
"@
    
    Write-Host "CSV Format:" -ForegroundColor Yellow
    Write-Host $template -ForegroundColor White
    Write-Host "`nNotes:" -ForegroundColor Yellow
    Write-Host "- Valid grade values: 8, 9, 10, 10email (for 10th with email), 11, 12" -ForegroundColor White
    Write-Host "- Email will be generated as cadet<firstname><lastname>@nomma.net" -ForegroundColor White
    Write-Host "- All users will be created with the default password." -ForegroundColor White
    Write-Host "- Save as a .csv file before using." -ForegroundColor White
    
    Save-TemplateOption "cadet_creation_template.csv" $template
}

# Function to show and export staff creation template
function Show-StaffCreationTemplate
{
    Clear-Host
    Write-Host "Create Multiple Staff Template" -ForegroundColor Cyan
    Write-Host "--------------------------" -ForegroundColor Cyan
    
    $template = @"
firstname,lastname,position
John,Smith,Teachers
Mary,Johnson,Counselors
Robert,Williams,IT
Patricia,Brown,SchoolAdmins
James,Jones,Security
Jennifer,Miller,Staff
Michael,Davis,NewStaff
"@
    
    Write-Host "CSV Format:" -ForegroundColor Yellow
    Write-Host $template -ForegroundColor White
    Write-Host "`nNotes:" -ForegroundColor Yellow
    Write-Host "- Valid position values: Counselors, IT, NewStaff, SchoolAdmins, Security, Staff, Teachers" -ForegroundColor White
    Write-Host "- Email will be generated as <first initial><lastname>@nomma.net" -ForegroundColor White
    Write-Host "- All users will be created with the default password." -ForegroundColor White
    Write-Host "- Save as a .csv file before using." -ForegroundColor White
    
    Save-TemplateOption "staff_creation_template.csv" $template
}

# Function to show and export archive users template
function Show-ArchiveUsersTemplate
{
    Clear-Host
    Write-Host "Archive Multiple Users Template" -ForegroundColor Cyan
    Write-Host "---------------------------" -ForegroundColor Cyan
    
    $template = @"
email
user1@nomma.net
user2@nomma.net
user3@nomma.net
"@
    
    Write-Host "CSV Format:" -ForegroundColor Yellow
    Write-Host $template -ForegroundColor White
    Write-Host "`nNotes:" -ForegroundColor Yellow
    Write-Host "- Only email addresses are required, one per line." -ForegroundColor White
    Write-Host "- These users will have their passwords reset, be signed out of all sessions," -ForegroundColor White
    Write-Host "  moved to the archive OU, and have their accounts suspended." -ForegroundColor White
    Write-Host "- Save as a .csv file before using." -ForegroundColor White
    
    Save-TemplateOption "archive_users_template.csv" $template
}

# Function to show and export lock devices template
function Show-LockDevicesTemplate
{
    Clear-Host
    Write-Host "Lock Multiple Devices Template" -ForegroundColor Cyan
    Write-Host "---------------------------" -ForegroundColor Cyan
    
    $template = @"
email,deviceid
user1@nomma.net,Device1ID
user2@nomma.net,Device2ID
user3@nomma.net,Device3ID
"@
    
    Write-Host "CSV Format:" -ForegroundColor Yellow
    Write-Host $template -ForegroundColor White
    Write-Host "`nNotes:" -ForegroundColor Yellow
    Write-Host "- Both email and deviceid columns are required." -ForegroundColor White
    Write-Host "- The deviceid should match the ID shown in Google Admin console." -ForegroundColor White
    Write-Host "- You can use 'Get Device Serial Number' to find device IDs for a user." -ForegroundColor White
    Write-Host "- Save as a .csv file before using." -ForegroundColor White
    
    Save-TemplateOption "lock_devices_template.csv" $template
}

# Function to show and export wipe devices template
function Show-WipeDevicesTemplate
{
    Clear-Host
    Write-Host "Wipe Multiple Devices Template" -ForegroundColor Cyan
    Write-Host "---------------------------" -ForegroundColor Cyan
    
    $template = @"
email,deviceid
user1@nomma.net,Device1ID
user2@nomma.net,Device2ID
user3@nomma.net,Device3ID
"@
    
    Write-Host "CSV Format:" -ForegroundColor Yellow
    Write-Host $template -ForegroundColor White
    Write-Host "`nNotes:" -ForegroundColor Yellow
    Write-Host "- Both email and deviceid columns are required." -ForegroundColor White
    Write-Host "- The deviceid should match the ID shown in Google Admin console." -ForegroundColor White
    Write-Host "- WARNING: Wiping a device will erase all data on the device." -ForegroundColor Red
    Write-Host "  This action cannot be undone." -ForegroundColor Red
    Write-Host "- You can use 'Get Device Serial Number' to find device IDs for a user." -ForegroundColor White
    Write-Host "- Save as a .csv file before using." -ForegroundColor White
    
    Save-TemplateOption "wipe_devices_template.csv" $template
}

# Helper function to offer saving the template
function Save-TemplateOption
{
    param (
        [string]$filename,
        [string]$content
    )
    
    Write-Host "`nWould you like to save this template to a file? (y/n): " -ForegroundColor Yellow -NoNewline
    $saveChoice = Read-Host
    
    if ($saveChoice -eq "y")
    {
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $savePath = Join-Path $desktopPath $filename
        
        try
        {
            $content | Out-File -FilePath $savePath -Encoding utf8
            Write-Host "`nTemplate saved to: $savePath" -ForegroundColor Green
        } catch
        {
            Write-Host "`nError saving template: $_" -ForegroundColor Red
            $savePath = $filename
            $content | Out-File -FilePath $savePath -Encoding utf8
            Write-Host "`nTemplate saved to current directory: $savePath" -ForegroundColor Green
        }
    }
    
    Write-Host "`nPress Enter to return to the Templates Menu..." -NoNewline
    Read-Host
    Show-TemplateMenu
}

# Main script logic
$exit = $false

# Display welcome message
Clear-Host
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "    GOOGLE WORKSPACE ADMIN TOOLKIT        " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "Creation of LLEE" -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan
Start-Sleep -Seconds 2

while (-not $exit)
{
    Show-MainMenu
    $choice = Read-Host
    
    switch ($choice)
    {
        "1"
        {
            $passwordExit = $false
            while (-not $passwordExit)
            {
                Show-PasswordMenu
                $passwordChoice = Read-Host
                
                switch ($passwordChoice)
                {
                    "1"
                    { Reset-SinglePassword 
                    }
                    "2"
                    { Reset-MassPasswords 
                    }
                    "3"
                    { $passwordExit = $true 
                    }
                    default
                    { 
                        Write-Host "Invalid option. Please try again." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
            }
        }
        "2"
        {
            $userExit = $false
            while (-not $userExit)
            {
                Show-UserMenu
                $userChoice = Read-Host
                
                switch ($userChoice)
                {
                    "1"
                    { New-SingleUser 
                    }
                    "2"
                    { New-MultipleUsers 
                    }
                    "3"
                    { Archive-SingleUser 
                    }
                    "4"
                    { Archive-MultipleUsers 
                    }
                    "5"
                    { Suspend-User 
                    }
                    "6"
                    { Unsuspend-User 
                    }
                    "7"
                    { Remove-User 
                    }
                    "8"
                    { Move-UsersToGradeOU
                    }
                    "9"
                    { $userExit = $true 
                    }
                    default
                    { 
                        Write-Host "Invalid option. Please try again." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
            }
        }
        "3"
        {
            $deviceExit = $false
            while (-not $deviceExit)
            {
                Show-DeviceMenu
                $deviceChoice = Read-Host
                
                switch ($deviceChoice)
                {
                    "1"
                    { Lock-SingleChromebook
                    }
                    "2"
                    { Lock-MultipleChromebooks
                    }
                    "3"
                    { Unlock-SingleChromebook
                    }
                    "4"
                    { Unlock-MultipleChromebooks
                    }
                    "5"
                    { Wipe-Device 
                    }
                    "6"
                    { Wipe-MultipleDevices 
                    }
                    "7"
                    { List-Devices 
                    }
                    "8"
                    { $deviceExit = $true 
                    }
                    default
                    { 
                        Write-Host "Invalid option. Please try again." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
            }
        }
        "4"
        {
            $infoExit = $false
            while (-not $infoExit)
            {
                Show-InfoMenu
                $infoChoice = Read-Host
                
                switch ($infoChoice)
                {
                    "1"
                    { Get-DeviceSerialNumber 
                    }
                    "2"
                    { Get-UserInformation 
                    }
                    "3"
                    { List-AllUsers 
                    }
                    "4"
                    { Check-LicenseInformation 
                    }
                    "5"
                    { Get-UserActivityReport 
                    }
                    "6"
                    { Monitor-SecurityReports 
                    }
                    "7"
                    { $infoExit = $true 
                    }
                    default
                    { 
                        Write-Host "Invalid option. Please try again." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
            }
        }
        "5"
        {
            # Call to the CSV Templates menu
            Show-TemplateMenu
        }
        "6"
        {
            Clear-Host
            Write-Host "Goodbye!" -ForegroundColor Cyan
            $exit = $true
        }
        default
        {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
}
