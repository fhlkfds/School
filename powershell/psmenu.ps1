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
    Write-Host "5. Exit" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-5]: " -ForegroundColor Yellow -NoNewline
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

# Function to display user management submenu
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
    Write-Host "8. Return to Main Menu" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Default Password: $defaultPassword" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-8]: " -ForegroundColor Yellow -NoNewline
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
    Write-Host "6. Return to Main Menu" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-6]: " -ForegroundColor Yellow -NoNewline
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


# Function to create a single user
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
    
    # Handle based on account type
    if ($accountType -eq "1")
    {
        # Cadet account
        $email = "cadet$firstName$lastName@nomma.net".ToLower()
        
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
        & gam create user "$email" firstname "$firstName" lastname "$lastName" password "Password@1" org "$orgUnit"
    } else
    {
        # Staff account - first initial + last name
        $firstInitial = $firstName.Substring(0, 1)
        $email = "$firstInitial$lastName@nomma.net".ToLower()
        
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

# Function to create multiple users from CSV
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
                # Student (Cadet) account
                $email = "cadet$($user.firstname)$($user.lastname)@nomma.net".ToLower()
                
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
                & gam create user "$email" firstname "$($user.firstname)" lastname "$($user.lastname)" password "Password@1" org "$gradeOU"
            } else
            {
                # Staff account - first initial + last name
                $firstInitial = $user.firstname.Substring(0, 1)
                $email = "$firstInitial$($user.lastname)@nomma.net".ToLower()
                
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
                & gam create user "$email" firstname "$($user.firstname)" lastname "$($user.lastname)" password "Password@1" org "$positionOU"
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

# Function to archive a single user
function Archive-SingleUser
{
    Clear-Host
    Write-Host "Archive User" -ForegroundColor Cyan
    Write-Host "-----------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    $archiveOU = Read-Host "Enter archive OU (default: /Archived Users)"
    
    if ([string]::IsNullOrEmpty($archiveOU))
    {
        $archiveOU = "/Archived Users"
    }
    
    Write-Host "Archiving user $userEmail..." -ForegroundColor Yellow
    
    # Reset password to the default secure password
    & gam update user "$userEmail" password "$defaultPassword"
    
    # Sign out from all sessions
    & gam user "$userEmail" signout
    
    # Move to archive OU
    & gam update user "$userEmail" org "$archiveOU"
    
    # Suspend the account
    & gam update user "$userEmail" suspended on
    
    Write-Host "User $userEmail has been archived" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}

# Function to archive multiple users from CSV
function Archive-MultipleUsers
{
    Clear-Host
    Write-Host "Archive Multiple Users (CSV)" -ForegroundColor Cyan
    Write-Host "--------------------------" -ForegroundColor Cyan
    Write-Host "CSV format should be: email" -ForegroundColor Yellow
    $csvPath = Read-Host "Enter path to CSV file"
    $archiveOU = Read-Host "Enter archive OU (default: /Archived Users)"
    
    if ([string]::IsNullOrEmpty($archiveOU))
    {
        $archiveOU = "/Archived Users"
    }
    
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
            Write-Host "Skipping invalid entry" -ForegroundColor Yellow
            $errorCount++
            continue
        }
        
        try
        {
            Write-Host "Archiving user $($user.email)..." -ForegroundColor Yellow
            
            # Reset password to the default secure password
            & gam update user "$($user.email)" password "$defaultPassword"
            
            # Sign out from all sessions
            & gam user "$($user.email)" signout
            
            # Move to archive OU
            & gam update user "$($user.email)" org "$archiveOU"
            
            # Suspend the account
            & gam update user "$($user.email)" suspended on
            
            Write-Host "Archive successful for $($user.email)" -ForegroundColor Green
            $successCount++
        } catch
        {
            Write-Host "Error archiving user $($user.email): $_" -ForegroundColor Red
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

# Function to lock a single mobile device
function Lock-SingleDevice
{
    Clear-Host
    Write-Host "Lock Single Mobile Device" -ForegroundColor Cyan
    Write-Host "----------------------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    
    # List devices for the user
    Write-Host "Listing devices for $userEmail..." -ForegroundColor Yellow
    & gam print mobile query "user:$userEmail" | Out-Host
    
    $deviceId = Read-Host "Enter device ID to lock"
    
    Write-Host "Locking device $deviceId for $userEmail..." -ForegroundColor Yellow
    & gam user "$userEmail" update mobile "$deviceId" action accountlock
    Write-Host "Device has been locked" -ForegroundColor Green
    Read-Host "Press Enter to continue..."
}

# Function to lock multiple mobile devices from CSV
function Lock-MultipleDevices
{
    Clear-Host
    Write-Host "Lock Multiple Mobile Devices (CSV)" -ForegroundColor Cyan
    Write-Host "-------------------------------" -ForegroundColor Cyan
    Write-Host "CSV format should be: email,deviceid" -ForegroundColor Yellow
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
        if ([string]::IsNullOrEmpty($device.email) -or [string]::IsNullOrEmpty($device.deviceid))
        {
            Write-Host "Skipping invalid entry" -ForegroundColor Yellow
            $errorCount++
            continue
        }
        
        try
        {
            Write-Host "Locking device $($device.deviceid) for $($device.email)..." -ForegroundColor Yellow
            & gam user "$($device.email)" update mobile "$($device.deviceid)" action accountlock
            Write-Host "Device lock successful" -ForegroundColor Green
            $successCount++
        } catch
        {
            Write-Host "Error locking device: $_" -ForegroundColor Red
            $errorCount++
        }
    }
    
    Write-Host "`nDevice locking complete!" -ForegroundColor Cyan
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

# Function to get device serial number
function Get-DeviceSerialNumber
{
    Clear-Host
    Write-Host "Get Device Serial Number" -ForegroundColor Cyan
    Write-Host "----------------------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    
    Write-Host "Retrieving devices for $userEmail..." -ForegroundColor Yellow
    & gam print mobile query "user:$userEmail" | Out-Host
    Read-Host "Press Enter to continue..."
}

# Function to get user information
function Get-UserInformation
{
    Clear-Host
    Write-Host "Get User Information" -ForegroundColor Cyan
    Write-Host "------------------" -ForegroundColor Cyan
    $userEmail = Read-Host "Enter user email"
    
    Write-Host "Retrieving information for $userEmail..." -ForegroundColor Yellow
    & gam info user "$userEmail" | Out-Host
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

# Main script logic
$exit = $false

# Display welcome message
Clear-Host
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "    GOOGLE WORKSPACE ADMIN TOOLKIT        " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "Default password set to: $defaultPassword" -ForegroundColor Yellow
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
                    { Lock-SingleDevice 
                    }
                    "2"
                    { Lock-MultipleDevices 
                    }
                    "3"
                    { Wipe-Device 
                    }
                    "4"
                    { Wipe-MultipleDevices 
                    }
                    "5"
                    { List-Devices 
                    }
                    "6"
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
