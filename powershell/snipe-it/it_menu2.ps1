<#
.SYNOPSIS
    Snipe-IT PowerShell Menu System
.DESCRIPTION
    This script provides a menu-driven interface for common Snipe-IT operations using the snipeitps module.
    Categories include: Asset Management, User Support, Identity Management, Routine Maintenance, Monitoring, and Reporting.
.NOTES
    Requires the snipeitps PowerShell module to be installed.
#>

# Check if snipeitps module is installed, install if not
if (-not (Get-Module -ListAvailable -Name "snipeitps"))
{
    Write-Host "SnipeIT PS module not found. Installing now..." -ForegroundColor Yellow
    Install-Module -Name snipeitps -Force -Scope CurrentUser
}

# Import the module
Import-Module snipeitps

# Configuration - Replace these with your actual values
$apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIzIiwianRpIjoiM2MzMGM0NzBmNmRkMGJiODQ0NGI1MDc2NjkwZWY2MzA0MTUzZjQzODhiMDA1MTEzNmM4NzE4NDc2NjQwOTZkYTk2YTQ5MDY4NWRlMzhkY2UiLCJpYXQiOjE3NDc0MDUyODUuNjUzOTY2LCJuYmYiOjE3NDc0MDUyODUuNjUzOTczLCJleHAiOjIyMjA3OTA4ODUuNjMzMTQ5LCJzdWIiOiIxIiwic2NvcGVzIjpbXX0.Bx9nImH2fmqQeWp1KjUi1_hbSpgOF-YJ__BLOthuhct-9OyaJzTgq7wjZ8DvYtEFvVIAr-_wXI37Ihe1PqOPT5SPuYqHl1vES51OQEFHNVHcPSPjQ5gJFraKY4f8Yqs26V5jiEYKo-z7wGfHRpEKAg3MzC8GgfIZUCbh-Xg5OmvdCjtYLQrsFB1G4M2alkGQyBzotI2QV__76JlA1dQIUdAX_6ZNadjxEVG0-GF1CPOO4IrYPZN-YZ6zztCEO8lR0vxSGj-Dtu1WCPqJM4iuE1Jy5TUeyLTCMOtk2Nw_G-LD_w_W6hhEhsxMca8HPwvDnN7V8YHYx1V5uTE5nacHw_gTTpK70kLV-XECljW3rSwfoV0SepHTml3GEECZk4EyNr4vK5DSk5DZwfrjUjzOBGqphyqH1q6mNU296-H5L7OrKfwEIO2HUscuyS6842JDBVZoFH-L2WEYc_PuX2Nndbc0vj0MW8kZUjypVJOn0_biBs2-xEPzgN7mroYGMf5xyaeWgomwEfaA-GX2fYfp7ovWLUhe4KXkkW16kGGgqkKqA62lC8lDYrUbCJuATJGMgBDNGeiSroldB7XlCmmskOwq2AcCUNyKbKMQZWJS89BgYjFmyM__Djv18-3oa0JW6w1norstpOzL8VExCSvM3jE-p2C0r9VMzEYKbaUfrcQ"
$snipeURL = "https://inv.nomma.lan"

# Configure connection to Snipe-IT
Set-SnipeitInfo -URL $snipeURL -APIKey $apiKey

# Clear screen and display menu
function Show-MainMenu
{
    Clear-Host
    Write-Host "===== Snipe-IT Management Console =====" -ForegroundColor Cyan
    Write-Host "1. Asset Management" -ForegroundColor Green
    Write-Host "2. User Support & Identity Management" -ForegroundColor Green
    Write-Host "3. Routine Maintenance & Monitoring" -ForegroundColor Green
    Write-Host "4. Reporting" -ForegroundColor Green
    Write-Host "5. CSV Templates" -ForegroundColor Green
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host "=====================================" -ForegroundColor Cyan
}

# Asset Management submenu
function Show-AssetManagementMenu
{
    Clear-Host
    Write-Host "===== Asset Management =====" -ForegroundColor Cyan
    Write-Host "1. Add New Asset" -ForegroundColor Green
    Write-Host "2. Check-out Asset" -ForegroundColor Green
    Write-Host "3. Check-in Asset" -ForegroundColor Green
    Write-Host "4. Update Asset Information" -ForegroundColor Green
    Write-Host "5. Bulk Import Assets" -ForegroundColor Green
    Write-Host "B. Back to Main Menu" -ForegroundColor Yellow
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host "===========================" -ForegroundColor Cyan
}

# User Support & Identity Management submenu
function Show-UserManagementMenu
{
    Clear-Host
    Write-Host "===== User Support & Identity Management =====" -ForegroundColor Cyan
    Write-Host "1. Add New User" -ForegroundColor Green
    Write-Host "2. Update User Information" -ForegroundColor Green
    Write-Host "3. View User Assets" -ForegroundColor Green
    Write-Host "4. Disable User" -ForegroundColor Green
    Write-Host "5. Bulk Import Users" -ForegroundColor Green
    Write-Host "B. Back to Main Menu" -ForegroundColor Yellow
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Cyan
}

# Routine Maintenance & Monitoring submenu
function Show-MaintenanceMenu
{
    Clear-Host
    Write-Host "===== Routine Maintenance & Monitoring =====" -ForegroundColor Cyan
    Write-Host "1. Schedule Maintenance" -ForegroundColor Green
    Write-Host "2. View Overdue Maintenance" -ForegroundColor Green
    Write-Host "3. View License Expirations" -ForegroundColor Green
    Write-Host "4. Check for Missing Assets" -ForegroundColor Green
    Write-Host "5. System Health Check" -ForegroundColor Green
    Write-Host "B. Back to Main Menu" -ForegroundColor Yellow
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Cyan
}

# Reporting submenu
function Show-ReportingMenu
{
    Clear-Host
    Write-Host "===== Reporting =====" -ForegroundColor Cyan
    Write-Host "1. Asset Audit Report" -ForegroundColor Green
    Write-Host "2. License Compliance Report" -ForegroundColor Green
    Write-Host "3. Depreciation Report" -ForegroundColor Green
    Write-Host "4. Activity Log Report" -ForegroundColor Green
    Write-Host "5. Custom Report Builder" -ForegroundColor Green
    Write-Host "B. Back to Main Menu" -ForegroundColor Yellow
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host "===================" -ForegroundColor Cyan
}

# CSV Templates submenu
function Show-TemplatesMenu
{
    Clear-Host
    Write-Host "===== CSV Templates =====" -ForegroundColor Cyan
    Write-Host "1. Generate Asset Import Template" -ForegroundColor Green
    Write-Host "2. Generate User Import Template" -ForegroundColor Green
    Write-Host "3. Generate License Import Template" -ForegroundColor Green
    Write-Host "4. Generate Maintenance Import Template" -ForegroundColor Green
    Write-Host "5. View Template Documentation" -ForegroundColor Green
    Write-Host "B. Back to Main Menu" -ForegroundColor Yellow
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host "======================" -ForegroundColor Cyan
}



# Function to check out laptop to a student
function Checkout-Laptop
{
    Clear-Host
    Write-Host "===== Check Out Laptop =====" -ForegroundColor Cyan
    
    # Variable to control the workflow
    $step = 1
    $userInternalId = $null
    $selectedUser = $null
    $laptop = $null
    $charger = $null
    $laptopCheckoutSuccess = $false
    $chargerCheckoutSuccess = $false
    
    # Enable verbose logging to help diagnose issues
    $VerbosePreference = "Continue"
    
    while ($true)
    {
        switch ($step)
        {
            1
            { # Get student identifier
                Clear-Host
                Write-Host "===== Check Out Laptop - Step 1: Find Student =====" -ForegroundColor Cyan
                $studentIdentifier = Read-Host "Enter student ID number or email (or 'q' to quit)"
                
                if ($studentIdentifier -eq 'q')
                {
                    return
                }
                
                # Validate student exists in Snipe-IT using the specific search command
                try
                {
                    Write-Verbose "Searching for student with identifier: $studentIdentifier"
                    # Use the specific search command to find the user
                    $user = Get-SnipeitUser -search $studentIdentifier
                    
                    if ($null -eq $user -or $user.Count -eq 0)
                    {
                        Write-Host "Student not found in the system. Please check the ID or email and try again." -ForegroundColor Red
                        Start-Sleep -Seconds 3
                        continue
                    }
                    
                    # If multiple users found, let the user select the correct one
                    if ($user.Count -gt 1)
                    {
                        Write-Host "Multiple users found with identifier '$studentIdentifier':" -ForegroundColor Yellow
                        for ($i = 0; $i -lt $user.Count; $i++)
                        {
                            Write-Host "[$i] Name: $($user[$i].name), Email: $($user[$i].email), ID: $($user[$i].id)" -ForegroundColor Cyan
                        }
                        
                        $userIndex = Read-Host "Enter the number of the correct user (or 'b' to go back)"
                        
                        if ($userIndex -eq 'b')
                        {
                            continue
                        }
                        
                        if (-not [int]::TryParse($userIndex, [ref]$null) -or $userIndex -lt 0 -or $userIndex -ge $user.Count)
                        {
                            Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                            Start-Sleep -Seconds 2
                            continue
                        }
                        
                        $selectedUser = $user[$userIndex]
                    } else
                    {
                        $selectedUser = $user
                    }
                    
                    # Display found user info without confirmation
                    Write-Host "Student found:" -ForegroundColor Green
                    Write-Host "Name: $($selectedUser.name)" -ForegroundColor Yellow
                    Write-Host "Email: $($selectedUser.email)" -ForegroundColor Yellow
                    Write-Host "Internal ID: $($selectedUser.id)" -ForegroundColor Yellow
                    
                    # Store the user's internal ID for later use
                    $userInternalId = $selectedUser.id
                    Write-Verbose "User internal ID: $userInternalId"
                    
                    # Wait 3 seconds before proceeding to the next step
                    Start-Sleep -Seconds 3
                    $step = 2
                    
                } catch
                {
                    Write-Host "An error occurred while searching for the student: $_" -ForegroundColor Red
                    Write-Verbose "Error details: $($_.Exception.Message)"
                    Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
                    Start-Sleep -Seconds 3
                    continue
                }
            }
            
            2
            { # Get laptop asset tag
                Clear-Host
                Write-Host "===== Check Out Laptop - Step 2: Find Laptop =====" -ForegroundColor Cyan
                Write-Host "Student: $($selectedUser.name)" -ForegroundColor Green
                
                $laptopAssetTag = Read-Host "Enter laptop asset tag (or 'b' to go back, 'q' to quit)"
                
                if ($laptopAssetTag -eq 'q')
                {
                    return
                }
                
                if ($laptopAssetTag -eq 'b')
                {
                    $step = 1
                    continue
                }
                
                # Validate laptop exists
                try
                {
                    Write-Verbose "Searching for laptop with asset tag: $laptopAssetTag"
                    $laptop = Get-SnipeitAsset -asset_tag $laptopAssetTag
                    
                    if ($null -eq $laptop)
                    {
                        Write-Host "Laptop with asset tag $laptopAssetTag not found." -ForegroundColor Red
                        Start-Sleep -Seconds 3
                        continue
                    }
                    
                    # Display laptop info
                    Write-Host "Laptop found:" -ForegroundColor Green
                    Write-Host "Name: $($laptop.name)" -ForegroundColor Yellow
                    Write-Host "Asset ID: $($laptop.id)" -ForegroundColor Yellow
                    Write-Host "Model: $($laptop.model.name)" -ForegroundColor Yellow
                    Write-Host "Status: $($laptop.status_label.name)" -ForegroundColor Yellow
                    
                    # Check if laptop is already checked out
                    if ($laptop.assigned_to -ne $null)
                    {
                        Write-Host "WARNING: This laptop is ALREADY checked out to: $($laptop.assigned_to.name)" -ForegroundColor Red
                        $alreadyAssignedAction = Read-Host "Options: [a] Assign to another laptop, [f] Force check-out anyway, [b] Go back, [q] Quit"
                        
                        switch ($alreadyAssignedAction.ToLower())
                        {
                            "a"
                            { continue 
                            } # Stay on this step to enter a different laptop
                            "f"
                            { 
                                Write-Host "Proceeding with force checkout..." -ForegroundColor Yellow
                                # Continue with this laptop
                            }
                            "b"
                            { 
                                $step = 1
                                continue 
                            }
                            "q"
                            { return 
                            }
                            default
                            { 
                                Write-Host "Invalid option. Please try again." -ForegroundColor Red
                                Start-Sleep -Seconds 2
                                continue
                            }
                        }
                    }
                    
                    # Wait 3 seconds before proceeding to the next step
                    Start-Sleep -Seconds 3
                    $step = 3
                    
                } catch
                {
                    Write-Host "An error occurred while searching for the laptop: $_" -ForegroundColor Red
                    Write-Verbose "Error details: $($_.Exception.Message)"
                    Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
                    Start-Sleep -Seconds 3
                    continue
                }
            }
            
            3
            { # Get charger asset tag
                Clear-Host
                Write-Host "===== Check Out Laptop - Step 3: Find Charger =====" -ForegroundColor Cyan
                Write-Host "Student: $($selectedUser.name)" -ForegroundColor Green
                Write-Host "Laptop: $($laptop.name) (Asset Tag: $($laptop.asset_tag))" -ForegroundColor Green
                
                $chargerAssetTag = Read-Host "Enter charger asset tag (or 'b' to go back, 'q' to quit, 's' to skip charger)"
                
                if ($chargerAssetTag -eq 'q')
                {
                    return
                }
                
                if ($chargerAssetTag -eq 'b')
                {
                    $step = 2
                    continue
                }
                
                if ($chargerAssetTag -eq 's')
                {
                    $charger = $null
                    $step = 4
                    continue
                }
                
                # Validate charger exists
                try
                {
                    Write-Verbose "Searching for charger with asset tag: $chargerAssetTag"
                    $charger = Get-SnipeitAsset -asset_tag $chargerAssetTag
                    
                    if ($null -eq $charger)
                    {
                        Write-Host "Charger with asset tag $chargerAssetTag not found." -ForegroundColor Red
                        Start-Sleep -Seconds 3
                        continue
                    }
                    
                    # Display charger info
                    Write-Host "Charger found:" -ForegroundColor Green
                    Write-Host "Name: $($charger.name)" -ForegroundColor Yellow
                    Write-Host "Asset ID: $($charger.id)" -ForegroundColor Yellow
                    Write-Host "Model: $($charger.model.name)" -ForegroundColor Yellow
                    Write-Host "Status: $($charger.status_label.name)" -ForegroundColor Yellow
                    
                    # Check if charger is already checked out
                    if ($charger.assigned_to -ne $null)
                    {
                        Write-Host "WARNING: This charger is ALREADY checked out to: $($charger.assigned_to.name)" -ForegroundColor Red
                        $alreadyAssignedAction = Read-Host "Options: [a] Assign another charger, [f] Force check-out anyway, [s] Skip charger, [b] Go back, [q] Quit"
                        
                        switch ($alreadyAssignedAction.ToLower())
                        {
                            "a"
                            { continue 
                            } # Stay on this step to enter a different charger
                            "f"
                            { 
                                Write-Host "Proceeding with force checkout..." -ForegroundColor Yellow
                                # Continue with this charger
                            }
                            "s"
                            { 
                                $charger = $null
                                $step = 4
                                continue
                            }
                            "b"
                            { 
                                $step = 2
                                continue 
                            }
                            "q"
                            { return 
                            }
                            default
                            { 
                                Write-Host "Invalid option. Please try again." -ForegroundColor Red
                                Start-Sleep -Seconds 2
                                continue
                            }
                        }
                    }
                    
                    # Wait 3 seconds before proceeding to the next step
                    Start-Sleep -Seconds 3
                    $step = 4
                    
                } catch
                {
                    Write-Host "An error occurred while searching for the charger: $_" -ForegroundColor Red
                    Write-Verbose "Error details: $($_.Exception.Message)"
                    Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
                    Start-Sleep -Seconds 3
                    continue
                }
            }
            
            4
            { # Get checkout details and complete checkout
                Clear-Host
                Write-Host "===== Check Out Laptop - Step 4: Complete Checkout =====" -ForegroundColor Cyan
                Write-Host "Student: $($selectedUser.name)" -ForegroundColor Green
                Write-Host "Laptop: $($laptop.name) (Asset Tag: $($laptop.asset_tag), ID: $($laptop.id))" -ForegroundColor Green
                if ($charger)
                {
                    Write-Host "Charger: $($charger.name) (Asset Tag: $($charger.asset_tag), ID: $($charger.id))" -ForegroundColor Green
                } else
                {
                    Write-Host "Charger: None selected" -ForegroundColor Yellow
                }
                
                $checkoutReason = Read-Host "Enter checkout reason (e.g., 'School Year 2025') (or 'b' to go back, 'q' to quit)"
                
                if ($checkoutReason -eq 'q')
                {
                    return
                }
                
                if ($checkoutReason -eq 'b')
                {
                    $step = 3
                    continue
                }
                
                $notes = Read-Host "Additional notes (press Enter if none) (or 'b' to go back, 'q' to quit)"
                
                if ($notes -eq 'q')
                {
                    return
                }
                
                if ($notes -eq 'b')
                {
                    continue
                }
                
                # Confirm checkout
                Write-Host "`nReady to check out the following:" -ForegroundColor Cyan
                Write-Host "Student: $($selectedUser.name) (Internal ID: $userInternalId)" -ForegroundColor Yellow
                Write-Host "Laptop: $($laptop.name) (Asset Tag: $($laptop.asset_tag), ID: $($laptop.id))" -ForegroundColor Yellow
                if ($charger)
                {
                    Write-Host "Charger: $($charger.name) (Asset Tag: $($charger.asset_tag), ID: $($charger.id))" -ForegroundColor Yellow
                } else
                {
                    Write-Host "Charger: None selected" -ForegroundColor Yellow
                }
                Write-Host "Reason: $checkoutReason" -ForegroundColor Yellow
                if ($notes)
                { Write-Host "Notes: $notes" -ForegroundColor Yellow 
                }
                
                $confirm = Read-Host "Proceed with checkout? (Y/N/b for back)"
                
                if ($confirm -eq 'b')
                {
                    continue
                }
                
                if ($confirm -ne "Y" -and $confirm -ne "y")
                {
                    return
                }
                
                # Perform the checkout operations
                try
                {
                    # Check out laptop to student
                    Write-Verbose "Attempting to check out laptop (ID: $($laptop.id)) to user (ID: $userInternalId)"
                    
                    # Test if parameters are valid
                    if ([string]::IsNullOrEmpty($laptop.id))
                    {
                        Write-Host "Error: Laptop ID is null or empty" -ForegroundColor Red
                        Start-Sleep -Seconds 3
                        return
                    }
                    
                    if ([string]::IsNullOrEmpty($userInternalId))
                    {
                        Write-Host "Error: User Internal ID is null or empty" -ForegroundColor Red
                        Start-Sleep -Seconds 3
                        return
                    }
                    
                    # Attempt to check out laptop
                    try
                    {
                        Write-Host "Checking out laptop..." -ForegroundColor Cyan
                        $checkoutLaptop = Set-SnipeitAssetOwner -id $laptop.id -assigned_id $userInternalId -checkout_to_type user -note "$checkoutReason - $notes"
                        
                        # Mark success
                        $laptopCheckoutSuccess = $true
                        Write-Host "Laptop successfully checked out to student." -ForegroundColor Green
                    } catch
                    {
                        if ($_.ToString() -like "*That asset is not available for checkout!*")
                        {
                            Write-Host "ERROR: The laptop is not available for checkout." -ForegroundColor Red
                            
                            $continueOption = Read-Host "Would you like to: [r] Try another laptop, [s] Skip to charger, [q] Quit?"
                            switch ($continueOption.ToLower())
                            {
                                "r"
                                {
                                    $step = 2  # Go back to laptop selection
                                    continue
                                }
                                "s"
                                {
                                    # Continue to charger checkout
                                }
                                default
                                {
                                    return
                                }
                            }
                        } else
                        {
                            Write-Host "Error checking out laptop: $_" -ForegroundColor Red
                            $continueOption = Read-Host "Skip laptop checkout and continue with charger? (Y/N)"
                            if ($continueOption -ne "Y" -and $continueOption -ne "y")
                            {
                                return
                            }
                        }
                    }
                    
                    # Check out charger to student if one was selected
                    if ($charger)
                    {
                        try
                        {
                            Write-Host "Checking out charger..." -ForegroundColor Cyan
                            $checkoutCharger = Set-SnipeitAssetOwner -id $charger.id -assigned_id $userInternalId -checkout_to_type user -note "$checkoutReason - $notes"
                            
                            # Mark success
                            $chargerCheckoutSuccess = $true
                            Write-Host "Charger successfully checked out to student." -ForegroundColor Green
                        } catch
                        {
                            if ($_.ToString() -like "*That asset is not available for checkout!*")
                            {
                                Write-Host "WARNING: The charger is not available for checkout." -ForegroundColor Red
                                Write-Host "This could be because:" -ForegroundColor Yellow
                                Write-Host "1. The charger status is not set to a deployable status" -ForegroundColor Yellow
                                Write-Host "2. The charger is already checked out to someone else" -ForegroundColor Yellow
                                
                                $retryOption = Read-Host "Would you like to: [r] Try another charger, [c] Continue without charger, [q] Quit?"
                                switch ($retryOption.ToLower())
                                {
                                    "r"
                                    {
                                        $step = 3  # Go back to charger selection
                                        continue
                                    }
                                    "c"
                                    {
                                        # Continue without charger
                                        Write-Host "Continuing without assigning a charger." -ForegroundColor Yellow
                                    }
                                    default
                                    {
                                        # If laptop was checked out but charger failed, we still continue to summary
                                        if ($laptopCheckoutSuccess)
                                        {
                                            Write-Host "Laptop was successfully checked out, but charger assignment failed." -ForegroundColor Yellow
                                            break
                                        }
                                        return
                                    }
                                }
                            } else
                            {
                                Write-Host "Error checking out charger: $_" -ForegroundColor Red
                                
                                # If laptop was checked out but charger failed, offer to try another charger
                                if ($laptopCheckoutSuccess)
                                {
                                    $retryOption = Read-Host "Would you like to: [r] Try another charger, [c] Continue without charger?"
                                    switch ($retryOption.ToLower())
                                    {
                                        "r"
                                        {
                                            $step = 3  # Go back to charger selection
                                            continue
                                        }
                                        default
                                        {
                                            # Continue without charger
                                            Write-Host "Continuing without assigning a charger." -ForegroundColor Yellow
                                        }
                                    }
                                } else
                                {
                                    # If neither was successful, offer to quit
                                    $continueOption = Read-Host "Both laptop and charger checkout failed. Quit? (Y/N)"
                                    if ($continueOption -eq "Y" -or $continueOption -eq "y")
                                    {
                                        return
                                    }
                                }
                            }
                        }
                    }
                    
                    # Final confirmation - show a summary based on what succeeded
                    Write-Host "`nCheckout Summary" -ForegroundColor Cyan
                    Write-Host "Student: $($selectedUser.name) (Internal ID: $userInternalId)" -ForegroundColor Yellow
                    
                    if ($laptopCheckoutSuccess)
                    {
                        Write-Host "Laptop: $($laptop.name) (Asset Tag: $($laptop.asset_tag)) - CHECKED OUT SUCCESSFULLY" -ForegroundColor Green
                    } else
                    {
                        Write-Host "Laptop: Not checked out" -ForegroundColor Red
                    }
                    
                    if ($charger)
                    {
                        if ($chargerCheckoutSuccess)
                        {
                            Write-Host "Charger: $($charger.name) (Asset Tag: $($charger.asset_tag)) - CHECKED OUT SUCCESSFULLY" -ForegroundColor Green
                        } else
                        {
                            Write-Host "Charger: $($charger.name) (Asset Tag: $($charger.asset_tag)) - CHECKOUT FAILED" -ForegroundColor Red
                            Write-Host "Note: To assign this charger later, you will need to use the Snipe-IT interface or run this script again." -ForegroundColor Yellow
                        }
                    } else
                    {
                        Write-Host "Charger: None selected" -ForegroundColor Yellow
                    }
                    
                    # Overall status
                    if ($laptopCheckoutSuccess)
                    {
                        if (!$charger || $chargerCheckoutSuccess)
                        {
                            Write-Host "`nAll requested items successfully checked out." -ForegroundColor Green
                        } else
                        {
                            Write-Host "`nPartially successful: Laptop checked out, but charger failed." -ForegroundColor Yellow
                        }
                    } else
                    {
                        Write-Host "`nCheckout failed: No items were successfully assigned." -ForegroundColor Red
                    }
                    
                    Write-Host "`nPress any key to continue..."
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    return
                    
                } catch
                {
                    Write-Host "An unhandled error occurred during checkout: $_" -ForegroundColor Red
                    Write-Verbose "Error details: $($_.Exception.Message)"
                    Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
                    Start-Sleep -Seconds 5
                }
            }
        }
    }
}

# Function to check in laptop from a student


# Function to check in laptop from a student
function Checkin-Laptop
{
    Clear-Host
    Write-Host "===== Check In Laptop =====" -ForegroundColor Cyan
    
    # Variable to control the workflow
    $step = 1
    $asset = $null
    $checkinSuccess = $false
    
    # Enable verbose logging to help diagnose issues
    $VerbosePreference = "Continue"
    
    while ($true)
    {
        switch ($step)
        {
            1
            { # Get asset tag
                Clear-Host
                Write-Host "===== Check In Laptop - Step 1: Find Asset =====" -ForegroundColor Cyan
                $assetTag = Read-Host "Enter asset tag (or 'q' to quit)"
                
                if ($assetTag -eq 'q')
                {
                    return
                }
                
                # Validate asset exists in Snipe-IT
                try
                {
                    Write-Verbose "Searching for asset with tag: $assetTag"
                    $asset = Get-SnipeitAsset -asset_tag $assetTag
                    
                    if ($null -eq $asset)
                    {
                        Write-Host "Asset not found in the system. Please check the asset tag and try again." -ForegroundColor Red
                        Start-Sleep -Seconds 3
                        continue
                    }
                    
                    # Display found asset info
                    Write-Host "Asset found:" -ForegroundColor Green
                    Write-Host "Name: $($asset.name)" -ForegroundColor Yellow
                    Write-Host "Model: $($asset.model.name)" -ForegroundColor Yellow
                    Write-Host "Status: $($asset.status_label.name)" -ForegroundColor Yellow
                    
                    # Check if asset is checked out
                    if ($asset.assigned_to -eq $null)
                    {
                        Write-Host "WARNING: This asset is not checked out to anyone." -ForegroundColor Red
                        $continueOption = Read-Host "Would you like to: [a] Try another asset, [q] Quit?"
                        
                        switch ($continueOption.ToLower())
                        {
                            "a"
                            { continue 
                            } # Stay on this step to enter a different asset
                            default
                            { return 
                            }
                        }
                    }
                    
                    # Display checkout info
                    Write-Host "Currently checked out to:" -ForegroundColor Green
                    Write-Host "Name: $($asset.assigned_to.name)" -ForegroundColor Yellow
                    if ($asset.assigned_to.type -eq "user")
                    {
                        Write-Host "Type: User" -ForegroundColor Yellow
                    } else
                    {
                        Write-Host "Type: $($asset.assigned_to.type)" -ForegroundColor Yellow
                    }
                    
                    # Wait 3 seconds before proceeding to the next step
                    Start-Sleep -Seconds 3
                    $step = 2
                    
                } catch
                {
                    Write-Host "An error occurred while searching for the asset: $_" -ForegroundColor Red
                    Write-Verbose "Error details: $($_.Exception.Message)"
                    Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
                    Start-Sleep -Seconds 3
                    continue
                }
            }
            
            2
            { # Get check-in details and complete check-in
                Clear-Host
                Write-Host "===== Check In Laptop - Step 2: Complete Check-in =====" -ForegroundColor Cyan
                Write-Host "Asset: $($asset.name) (Asset Tag: $($asset.asset_tag))" -ForegroundColor Green
                Write-Host "Currently checked out to: $($asset.assigned_to.name)" -ForegroundColor Green
                
                $notes = Read-Host "Additional notes (press Enter if none) (or 'b' to go back, 'q' to quit)"
                
                if ($notes -eq 'q')
                {
                    return
                }
                
                if ($notes -eq 'b')
                {
                    $step = 1
                    continue
                }
                
                # Check for condition change
                $conditionOptions = @(
                    "No Change", 
                    "New", 
                    "Good", 
                    "Fair", 
                    "Poor", 
                    "Broken"
                )
                
                Write-Host "`nSelect the current condition of the asset:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $conditionOptions.Count; $i++)
                {
                    Write-Host "[$i] $($conditionOptions[$i])" -ForegroundColor Yellow
                }
                
                $conditionIndex = Read-Host "Enter option number (or 'b' to go back, 'q' to quit)"
                
                if ($conditionIndex -eq 'q')
                {
                    return
                }
                
                if ($conditionIndex -eq 'b')
                {
                    continue
                }
                
                # Validate condition selection
                if (-not [int]::TryParse($conditionIndex, [ref]$null) -or 
                    $conditionIndex -lt 0 -or 
                    $conditionIndex -ge $conditionOptions.Count)
                {
                    Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
                
                $selectedCondition = $conditionOptions[$conditionIndex]
                
                # Confirm check-in
                Write-Host "`nReady to check in the following:" -ForegroundColor Cyan
                Write-Host "Asset: $($asset.name) (Asset Tag: $($asset.asset_tag), ID: $($asset.id))" -ForegroundColor Yellow
                Write-Host "Currently checked out to: $($asset.assigned_to.name)" -ForegroundColor Yellow
                Write-Host "New condition: $selectedCondition" -ForegroundColor Yellow
                if ($notes)
                { 
                    Write-Host "Notes: $notes" -ForegroundColor Yellow 
                }
                
                $confirm = Read-Host "Proceed with check-in? (Y/N/b for back)"
                
                if ($confirm -eq 'b')
                {
                    continue
                }
                
                if ($confirm -ne "Y" -and $confirm -ne "y")
                {
                    return
                }
                
                # Perform the check-in operation
                try
                {
                    Write-Verbose "Attempting to check in asset (ID: $($asset.id))"
                    
                    # Test if parameters are valid
                    if ([string]::IsNullOrEmpty($asset.id))
                    {
                        Write-Host "Error: Asset ID is null or empty" -ForegroundColor Red
                        Start-Sleep -Seconds 3
                        return
                    }
                    
                    # Prepare check-in parameters
                    $checkinParams = @{
                        id = $asset.id
                        note = $notes
                    }
                    
                    # Add status_id parameter if condition is changing
                    if ($selectedCondition -ne "No Change")
                    {
                        # Map condition to status ID (these would need to be configured for your environment)
                        $statusMap = @{
                            "New" = 1
                            "Good" = 2
                            "Fair" = 3
                            "Poor" = 4
                            "Broken" = 5
                        }
                        
                        if ($statusMap.ContainsKey($selectedCondition))
                        {
                            $checkinParams.Add("status_id", $statusMap[$selectedCondition])
                        }
                    }
                    
                    # Attempt to check in asset
                    Write-Host "Checking in asset..." -ForegroundColor Cyan
                    
                    # Use Reset-SnipeitAssetOwner for check-in with splatting to include all parameters
                    try
                    {
                        $checkin = Reset-SnipeitAssetOwner @checkinParams
                        
                        # Mark success
                        $checkinSuccess = $true
                        Write-Host "Asset successfully checked in." -ForegroundColor Green
                    } catch
                    {
                        Write-Host "Error during check-in: $_" -ForegroundColor Red
                        throw $_  # Re-throw to be caught by the outer catch block
                    }
                    
                    # Final confirmation - show a summary based on what succeeded
                    Write-Host "`nCheck-in Summary" -ForegroundColor Cyan
                    Write-Host "Asset: $($asset.name) (Asset Tag: $($asset.asset_tag))" -ForegroundColor Yellow
                    
                    if ($checkinSuccess)
                    {
                        Write-Host "Status: CHECKED IN SUCCESSFULLY" -ForegroundColor Green
                        Write-Host "New condition: $selectedCondition" -ForegroundColor Green
                        if ($notes)
                        {
                            Write-Host "Notes: $notes" -ForegroundColor Green
                        }
                    } else
                    {
                        Write-Host "Status: CHECK-IN FAILED" -ForegroundColor Red
                    }
                    
                    Write-Host "`nPress any key to continue..."
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    return
                    
                } catch
                {
                    Write-Host "An error occurred during check-in: $_" -ForegroundColor Red
                    Write-Verbose "Error details: $($_.Exception.Message)"
                    Write-Verbose "Stack trace: $($_.ScriptStackTrace)"
                    
                    if ($_.ToString() -like "*is not checked out*" -or $_.ToString() -like "*already checked in*")
                    {
                        Write-Host "This asset appears to be already checked in." -ForegroundColor Yellow
                    } elseif ($_.ToString() -like "*not found*")
                    {
                        Write-Host "Asset not found. It may have been deleted or the ID is incorrect." -ForegroundColor Yellow
                    } elseif ($_.ToString() -like "*access denied*" -or $_.ToString() -like "*permission*")
                    {
                        Write-Host "You don't have permission to check in this asset." -ForegroundColor Yellow
                    } else
                    {
                        Write-Host "Unknown error occurred. Please check that the snipeitps module is properly installed and configured." -ForegroundColor Yellow
                        Write-Host "You might need to run: Install-Module -Name snipeitps -Force -Scope CurrentUser" -ForegroundColor Yellow
                    }
                    
                    Write-Host "`nPress any key to continue..."
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    return
                }
            }
        }
    }
}

# Function to disable a user in Snipe-IT
function Disable-SnipeITUser
{
    Clear-Host
    Write-Host "===== Disable User =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Find the user
        $searchTerm = Read-Host "Enter username, email, or employee number to search for the user (or 'q' to quit)"
        if ($searchTerm -eq 'q')
        {
            return
        }
        
        Write-Host "Searching for user..." -ForegroundColor Yellow
        $users = Get-SnipeitUser -search $searchTerm
        
        if (!$users -or $users.Count -eq 0)
        {
            Write-Host "No users found with the search term: $searchTerm" -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }
        
        # Step 2: Select the user if multiple found
        $selectedUser = $null
        
        if ($users.Count -gt 1)
        {
            Write-Host "`nMultiple users found:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $users.Count; $i++)
            {
                $activeStatus = if ($users[$i].deleted_at)
                { "[INACTIVE]" 
                } else
                { "[ACTIVE]" 
                }
                Write-Host "[$i] $($users[$i].name) ($($users[$i].username)) - $($users[$i].email) $activeStatus" -ForegroundColor Cyan
            }
            
            $userIndex = Read-Host "Enter the number of the user to disable (or 'q' to quit)"
            if ($userIndex -eq 'q')
            {
                return
            }
            
            if ([int]::TryParse($userIndex, [ref]$null) -and [int]$userIndex -ge 0 -and [int]$userIndex -lt $users.Count)
            {
                $selectedUser = $users[[int]$userIndex]
            } else
            {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
                return
            }
        } else
        {
            $selectedUser = $users
        }
        
        # Step 3: Check if user is already inactive
        if ($selectedUser.deleted_at)
        {
            Write-Host "`nThis user is already inactive." -ForegroundColor Yellow
            
            # Option to reactivate
            $reactivate = Read-Host "Would you like to reactivate this user? (Y/N)"
            if ($reactivate -eq "Y" -or $reactivate -eq "y")
            {
                # Call the Snipe-IT API to restore the user
                try
                {
                    $restoredUser = Invoke-RestMethod -Method 'POST' -Uri "$snipeURL/api/v1/users/$($selectedUser.id)/restore" -Headers @{
                        "Authorization" = "Bearer $apiKey"
                        "Accept" = "application/json"
                        "Content-Type" = "application/json"
                    }
                    
                    if ($restoredUser.status -eq "success")
                    {
                        Write-Host "User successfully reactivated!" -ForegroundColor Green
                    } else
                    {
                        Write-Host "Failed to reactivate user: $($restoredUser.messages)" -ForegroundColor Red
                    }
                } catch
                {
                    Write-Host "An error occurred while trying to reactivate the user: $_" -ForegroundColor Red
                }
                
                Start-Sleep -Seconds 2
                return
            } else
            {
                return
            }
        }
        
        # Step 4: Confirm user disable
        Clear-Host
        Write-Host "===== Disable User Confirmation =====" -ForegroundColor Cyan
        Write-Host "You are about to disable the following user:" -ForegroundColor Yellow
        Write-Host "User ID: $($selectedUser.id)" -ForegroundColor Yellow
        Write-Host "Name: $($selectedUser.name)" -ForegroundColor Yellow
        Write-Host "Username: $($selectedUser.username)" -ForegroundColor Yellow
        Write-Host "Email: $($selectedUser.email)" -ForegroundColor Yellow
        
        # Check if user has assigned assets
        Write-Host "`nChecking for assigned assets..." -ForegroundColor Green
        $assets = Get-SnipeitAsset -assigned_to $selectedUser.id
        
        if ($assets -and $assets.Count -gt 0)
        {
            Write-Host "`nWARNING: This user has $($assets.Count) assets assigned to them." -ForegroundColor Red
            Write-Host "Assets should be checked in before disabling the user." -ForegroundColor Yellow
            
            # Option to view assets
            $viewAssets = Read-Host "Would you like to view these assets? (Y/N)"
            if ($viewAssets -eq "Y" -or $viewAssets -eq "y")
            {
                Write-Host "`nAssets assigned to $($selectedUser.name):" -ForegroundColor Green
                Write-Host "-----------------------------------------------------------" -ForegroundColor Gray
                Write-Host "| Asset Tag | Name | Model | Status |" -ForegroundColor Cyan
                Write-Host "-----------------------------------------------------------" -ForegroundColor Gray
                
                foreach ($asset in $assets)
                {
                    $assetTag = $asset.asset_tag
                    $name = if ($asset.name)
                    { $asset.name 
                    } else
                    { "N/A" 
                    }
                    $model = if ($asset.model)
                    { $asset.model.name 
                    } else
                    { "N/A" 
                    }
                    $status = if ($asset.status_label)
                    { $asset.status_label.name 
                    } else
                    { "N/A" 
                    }
                    
                    Write-Host "| $assetTag | $name | $model | $status |" -ForegroundColor White
                }
                Write-Host "-----------------------------------------------------------" -ForegroundColor Gray
            }
            
            # Ask whether to continue
            $continueDisable = Read-Host "`nDo you still want to disable this user? (Y/N)"
            if ($continueDisable -ne "Y" -and $continueDisable -ne "y")
            {
                Write-Host "User disable cancelled." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                return
            }
        }
        
        # Get a reason for disabling
        $notes = Read-Host "`nEnter a reason for disabling this user (optional)"
        
        # Final confirmation
        $confirmation = Read-Host "`nAre you ABSOLUTELY SURE you want to disable this user? This action can be reversed later. (Y/N)"
        if ($confirmation -ne "Y" -and $confirmation -ne "y")
        {
            Write-Host "User disable cancelled." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return
        }
        
        # Step 5: Disable the user by calling the API to delete (which soft-deletes in Snipe-IT)
        try
        {
            $disableResult = Invoke-RestMethod -Method 'DELETE' -Uri "$snipeURL/api/v1/users/$($selectedUser.id)" -Headers @{
                "Authorization" = "Bearer $apiKey"
                "Accept" = "application/json"
            }
            
            if ($disableResult.status -eq "success")
            {
                Write-Host "`nUser successfully disabled!" -ForegroundColor Green
                
                # Log the action with notes if provided
                if (![string]::IsNullOrWhiteSpace($notes))
                {
                    # Use the activity API to log a note if available
                    try
                    {
                        $activityPayload = @{
                            target_type = "user"
                            target_id = $selectedUser.id
                            action_type = "disabled"
                            notes = $notes
                        } | ConvertTo-Json
                        
                        $activityResult = Invoke-RestMethod -Method 'POST' -Uri "$snipeURL/api/v1/activities" -Headers @{
                            "Authorization" = "Bearer $apiKey"
                            "Accept" = "application/json"
                            "Content-Type" = "application/json"
                        } -Body $activityPayload
                    } catch
                    {
                        Write-Host "Note added but couldn't log activity." -ForegroundColor Yellow
                    }
                }
            } else
            {
                Write-Host "Failed to disable user: $($disableResult.messages)" -ForegroundColor Red
            }
        } catch
        {
            Write-Host "An error occurred while disabling the user: $_" -ForegroundColor Red
            Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
        }
    } catch
    {
        Write-Host "An error occurred: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to check for missing assets in Snipe-IT
function Check-MissingAssets
{
    Clear-Host
    Write-Host "===== Check for Missing Assets =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Define what "missing" means for your organization
        Write-Host "This function helps identify potentially missing assets based on various criteria." -ForegroundColor Yellow
        Write-Host "Choose which type of check to perform:" -ForegroundColor Green
        Write-Host "1. Assets not seen/audited in X days" -ForegroundColor Yellow
        Write-Host "2. Assets with specific status (e.g., 'Missing')" -ForegroundColor Yellow
        Write-Host "3. Checked-out assets without recent activity" -ForegroundColor Yellow
        Write-Host "4. Custom search by multiple criteria" -ForegroundColor Yellow
        Write-Host "5. Return to menu" -ForegroundColor Yellow
        
        $choice = Read-Host "Enter your choice (1-5)"
        
        switch ($choice)
        {
            "1"
            {
                # Option 1: Assets not audited in X days
                $dayThreshold = Read-Host "Enter number of days without an audit to consider as potentially missing"
                
                if (-not [int]::TryParse($dayThreshold, [ref]$null))
                {
                    Write-Host "Invalid number of days." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    return
                }
                
                $cutoffDate = (Get-Date).AddDays(-([int]$dayThreshold))
                $formattedDate = $cutoffDate.ToString("yyyy-MM-dd")
                
                Write-Host "`nSearching for assets not audited since $formattedDate..." -ForegroundColor Green
                
                # Get all assets
                $allAssets = Get-SnipeitAsset -all
                
                # Filter for assets with old or missing audit dates
                $missingAssets = $allAssets | Where-Object {
                    # Check if last_audit is missing or older than threshold
                    (-not $_.last_audit) -or (
                        ($_.last_audit -is [string] -and [DateTime]::Parse($_.last_audit) -lt $cutoffDate) -or
                        ($_.last_audit -is [DateTime] -and $_.last_audit -lt $cutoffDate)
                    )
                }
                
                Write-Host "`nFound $($missingAssets.Count) assets that haven't been audited in the last $dayThreshold days." -ForegroundColor Yellow
                
                if ($missingAssets.Count -gt 0)
                {
                    Display-AssetTable -assets $missingAssets -title "Assets Not Audited Since $formattedDate"
                    
                    # Option to export to CSV
                    $exportOption = Read-Host "Would you like to export this list to CSV? (Y/N)"
                    if ($exportOption -eq "Y" -or $exportOption -eq "y")
                    {
                        Export-AssetsToCSV -assets $missingAssets -fileNameSuffix "NotAuditedSince_$formattedDate"
                    }
                }
            }
            
            "2"
            {
                # Option 2: Assets with specific status
                Write-Host "`nRetrieving available statuses..." -ForegroundColor Green
                $statuses = Get-SnipeitStatus
                
                if ($statuses -and $statuses.Count -gt 0)
                {
                    Write-Host "Available Statuses:" -ForegroundColor Yellow
                    for ($i = 0; $i -lt $statuses.Count; $i++)
                    {
                        Write-Host "[$i] $($statuses[$i].name)" -ForegroundColor Cyan
                    }
                    
                    $statusIndex = Read-Host "Select status number to check for (or multiple separated by commas)"
                    
                    # Handle multiple status selections
                    $selectedStatusIds = @()
                    $statusIndices = $statusIndex -split ',' | ForEach-Object { $_.Trim() }
                    
                    foreach ($idx in $statusIndices)
                    {
                        if ([int]::TryParse($idx, [ref]$null) -and [int]$idx -ge 0 -and [int]$idx -lt $statuses.Count)
                        {
                            $selectedStatusIds += $statuses[[int]$idx].id
                        }
                    }
                    
                    if ($selectedStatusIds.Count -eq 0)
                    {
                        Write-Host "No valid status selected." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                        return
                    }
                    
                    # Get assets with the selected status(es)
                    $filteredAssets = @()
                    foreach ($statusId in $selectedStatusIds)
                    {
                        Write-Host "`nSearching for assets with status ID: $statusId..." -ForegroundColor Green
                        $statusAssets = Get-SnipeitAsset -status_id $statusId
                        $filteredAssets += $statusAssets
                    }
                    
                    Write-Host "`nFound $($filteredAssets.Count) assets with the selected status(es)." -ForegroundColor Yellow
                    
                    if ($filteredAssets.Count -gt 0)
                    {
                        Display-AssetTable -assets $filteredAssets -title "Assets with Selected Status"
                        
                        # Option to export to CSV
                        $exportOption = Read-Host "Would you like to export this list to CSV? (Y/N)"
                        if ($exportOption -eq "Y" -or $exportOption -eq "y")
                        {
                            Export-AssetsToCSV -assets $filteredAssets -fileNameSuffix "StatusReport"
                        }
                    }
                } else
                {
                    Write-Host "No status types found in Snipe-IT." -ForegroundColor Red
                }
            }
            
            "3"
            {
                # Option 3: Checked-out assets without recent activity
                $activityThreshold = Read-Host "Enter number of days of inactivity to consider as potentially missing"
                
                if (-not [int]::TryParse($activityThreshold, [ref]$null))
                {
                    Write-Host "Invalid number of days." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    return
                }
                
                $cutoffDate = (Get-Date).AddDays(-([int]$activityThreshold))
                $formattedDate = $cutoffDate.ToString("yyyy-MM-dd")
                
                Write-Host "`nSearching for checked-out assets without activity since $formattedDate..." -ForegroundColor Green
                
                # Get all checked-out assets (status typically 'Deployed')
                $deployedStatusId = (Get-SnipeitStatus | Where-Object { $_.name -match "Deployed|Assigned|Checked Out" } | Select-Object -First 1).id
                
                if (-not $deployedStatusId)
                {
                    Write-Host "Could not determine the 'Deployed' status ID. Please check status names in your Snipe-IT instance." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    return
                }
                
                $deployedAssets = Get-SnipeitAsset -status_id $deployedStatusId
                
                # Filter for assets with old or missing activity dates
                $inactiveAssets = $deployedAssets | Where-Object {
                    # Check if last_checkout is older than threshold and no recent activity
                    $lastActivity = if ($_.last_checkout -gt $_.last_update)
                    { $_.last_checkout 
                    } else
                    { $_.last_update 
                    }
                    
                    if ($lastActivity -is [string])
                    {
                        try
                        { [DateTime]::Parse($lastActivity) -lt $cutoffDate 
                        } catch
                        { $true 
                        }
                    } elseif ($lastActivity -is [DateTime])
                    {
                        $lastActivity -lt $cutoffDate
                    } else
                    {
                        $true # No activity date available
                    }
                }
                
                Write-Host "`nFound $($inactiveAssets.Count) checked-out assets without activity in the last $activityThreshold days." -ForegroundColor Yellow
                
                if ($inactiveAssets.Count -gt 0)
                {
                    Display-AssetTable -assets $inactiveAssets -title "Inactive Checked-Out Assets"
                    
                    # Option to export to CSV
                    $exportOption = Read-Host "Would you like to export this list to CSV? (Y/N)"
                    if ($exportOption -eq "Y" -or $exportOption -eq "y")
                    {
                        Export-AssetsToCSV -assets $inactiveAssets -fileNameSuffix "InactiveAssets_$formattedDate"
                    }
                }
            }
            
            "4"
            {
                # Option 4: Custom search with multiple criteria
                Write-Host "`nCustom Search - Enter criteria (leave blank to skip):" -ForegroundColor Green
                
                # Build search parameters
                $searchParams = @{}
                
                # Get Model options
                $getModels = Read-Host "Do you want to filter by model? (Y/N)"
                if ($getModels -eq "Y" -or $getModels -eq "y")
                {
                    $models = Get-SnipeitModel
                    if ($models -and $models.Count -gt 0)
                    {
                        Write-Host "Available Models:" -ForegroundColor Yellow
                        for ($i = 0; $i -lt $models.Count; $i++)
                        {
                            Write-Host "[$i] $($models[$i].name)" -ForegroundColor Cyan
                        }
                        
                        $modelIndex = Read-Host "Select model number"
                        if ([int]::TryParse($modelIndex, [ref]$null) -and [int]$modelIndex -ge 0 -and [int]$modelIndex -lt $models.Count)
                        {
                            $searchParams.Add("model_id", $models[[int]$modelIndex].id)
                        }
                    }
                }
                
                # Get Category options
                $getCategories = Read-Host "Do you want to filter by category? (Y/N)"
                if ($getCategories -eq "Y" -or $getCategories -eq "y")
                {
                    $categories = Get-SnipeitCategory
                    if ($categories -and $categories.Count -gt 0)
                    {
                        Write-Host "Available Categories:" -ForegroundColor Yellow
                        for ($i = 0; $i -lt $categories.Count; $i++)
                        {
                            Write-Host "[$i] $($categories[$i].name)" -ForegroundColor Cyan
                        }
                        
                        $categoryIndex = Read-Host "Select category number"
                        if ([int]::TryParse($categoryIndex, [ref]$null) -and [int]$categoryIndex -ge 0 -and [int]$categoryIndex -lt $categories.Count)
                        {
                            $searchParams.Add("category_id", $categories[[int]$categoryIndex].id)
                        }
                    }
                }
                
                # Get Location options
                $getLocations = Read-Host "Do you want to filter by location? (Y/N)"
                if ($getLocations -eq "Y" -or $getLocations -eq "y")
                {
                    $locations = Get-SnipeitLocation
                    if ($locations -and $locations.Count -gt 0)
                    {
                        Write-Host "Available Locations:" -ForegroundColor Yellow
                        for ($i = 0; $i -lt $locations.Count; $i++)
                        {
                            Write-Host "[$i] $($locations[$i].name)" -ForegroundColor Cyan
                        }
                        
                        $locationIndex = Read-Host "Select location number"
                        if ([int]::TryParse($locationIndex, [ref]$null) -and [int]$locationIndex -ge 0 -and [int]$locationIndex -lt $locations.Count)
                        {
                            $searchParams.Add("location_id", $locations[[int]$locationIndex].id)
                        }
                    }
                }
                
                # Other search parameters
                $assetTag = Read-Host "Asset Tag (partial match)"
                if (-not [string]::IsNullOrWhiteSpace($assetTag))
                {
                    $searchParams.Add("asset_tag", $assetTag)
                }
                
                $serialNumber = Read-Host "Serial Number (partial match)"
                if (-not [string]::IsNullOrWhiteSpace($serialNumber))
                {
                    $searchParams.Add("serial", $serialNumber)
                }
                
                $purchaseDate = Read-Host "Purchase Date older than (YYYY-MM-DD)"
                if (-not [string]::IsNullOrWhiteSpace($purchaseDate))
                {
                    try
                    {
                        $date = [DateTime]::ParseExact($purchaseDate, "yyyy-MM-dd", $null)
                        $searchParams.Add("purchase_date", $date.ToString("yyyy-MM-dd"))
                    } catch
                    {
                        Write-Host "Invalid date format. Ignoring purchase date filter." -ForegroundColor Yellow
                    }
                }
                
                # Perform the search
                Write-Host "`nSearching for assets with specified criteria..." -ForegroundColor Green
                
                # If no params specified, get all assets
                if ($searchParams.Count -eq 0)
                {
                    $resultAssets = Get-SnipeitAsset -all
                } else
                {
                    # Try to use splatting for the search
                    try
                    {
                        # Start with a basic search
                        $resultAssets = Get-SnipeitAsset @searchParams
                    } catch
                    {
                        Write-Host "Error performing search with multiple parameters: $_" -ForegroundColor Red
                        Write-Host "Falling back to sequential filtering..." -ForegroundColor Yellow
                        
                        # Fall back to getting all assets and filtering manually
                        $resultAssets = Get-SnipeitAsset -all
                        
                        # Apply filters sequentially
                        foreach ($key in $searchParams.Keys)
                        {
                            $value = $searchParams[$key]
                            
                            # Apply filter based on parameter type
                            switch ($key)
                            {
                                "model_id"
                                {
                                    $resultAssets = $resultAssets | Where-Object { $_.model.id -eq $value }
                                }
                                "category_id"
                                {
                                    $resultAssets = $resultAssets | Where-Object { $_.category.id -eq $value }
                                }
                                "location_id"
                                {
                                    $resultAssets = $resultAssets | Where-Object { $_.rtd_location.id -eq $value -or $_.location.id -eq $value }
                                }
                                "asset_tag"
                                {
                                    $resultAssets = $resultAssets | Where-Object { $_.asset_tag -like "*$value*" }
                                }
                                "serial"
                                {
                                    $resultAssets = $resultAssets | Where-Object { $_.serial -like "*$value*" }
                                }
                                "purchase_date"
                                {
                                    $dateCutoff = [DateTime]::ParseExact($value, "yyyy-MM-dd", $null)
                                    $resultAssets = $resultAssets | Where-Object { 
                                        if ($_.purchase_date)
                                        {
                                            try
                                            {
                                                $purchaseDate = if ($_.purchase_date -is [string])
                                                {
                                                    [DateTime]::Parse($_.purchase_date)
                                                } else
                                                {
                                                    $_.purchase_date
                                                }
                                                $purchaseDate -lt $dateCutoff
                                            } catch
                                            {
                                                $false
                                            }
                                        } else
                                        {
                                            $false
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
                Write-Host "`nFound $($resultAssets.Count) assets matching your criteria." -ForegroundColor Yellow
                
                if ($resultAssets.Count -gt 0)
                {
                    Display-AssetTable -assets $resultAssets -title "Custom Search Results"
                    
                    # Option to export to CSV
                    $exportOption = Read-Host "Would you like to export this list to CSV? (Y/N)"
                    if ($exportOption -eq "Y" -or $exportOption -eq "y")
                    {
                        Export-AssetsToCSV -assets $resultAssets -fileNameSuffix "CustomSearch"
                    }
                }
            }
            
            "5"
            {
                return
            }
            
            default
            {
                Write-Host "Invalid option. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
                return
            }
        }
    } catch
    {
        Write-Host "An error occurred: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Helper function to display assets in a table format
function Display-AssetTable
{
    param (
        [Parameter(Mandatory = $true)]
        [array]$assets,
        
        [Parameter(Mandatory = $false)]
        [string]$title = "Asset List"
    )
    
    # Determine the window width for better formatting
    $windowWidth = $Host.UI.RawUI.WindowSize.Width
    
    # Calculate column widths
    $tagWidth = [Math]::Min(12, $windowWidth * 0.15)
    $nameWidth = [Math]::Min(20, $windowWidth * 0.25)
    $modelWidth = [Math]::Min(15, $windowWidth * 0.2)
    $serialWidth = [Math]::Min(15, $windowWidth * 0.2)
    $statusWidth = [Math]::Min(12, $windowWidth * 0.15)
    
    # Create the header
    Write-Host "`n===== $title =====" -ForegroundColor Cyan
    Write-Host ("-" * $windowWidth) -ForegroundColor Gray
    $headerFormat = "| {0,-$tagWidth} | {1,-$nameWidth} | {2,-$modelWidth} | {3,-$serialWidth} | {4,-$statusWidth} |"
    Write-Host ($headerFormat -f "Asset Tag", "Name", "Model", "Serial", "Status") -ForegroundColor Cyan
    Write-Host ("-" * $windowWidth) -ForegroundColor Gray
    
    # Display data rows
    foreach ($asset in $assets)
    {
        $assetTag = if ($asset.asset_tag.Length -gt $tagWidth)
        { $asset.asset_tag.Substring(0, $tagWidth - 3) + "..." 
        } else
        { $asset.asset_tag 
        }
        $name = if ($asset.name.Length -gt $nameWidth)
        { $asset.name.Substring(0, $nameWidth - 3) + "..." 
        } else
        { $asset.name 
        }
        $model = if ($asset.model.name.Length -gt $modelWidth)
        { $asset.model.name.Substring(0, $modelWidth - 3) + "..." 
        } else
        { $asset.model.name 
        }
        $serial = if ($asset.serial.Length -gt $serialWidth)
        { $asset.serial.Substring(0, $serialWidth - 3) + "..." 
        } else
        { $asset.serial 
        }
        $status = if ($asset.status_label.name.Length -gt $statusWidth)
        { $asset.status_label.name.Substring(0, $statusWidth - 3) + "..." 
        } else
        { $asset.status_label.name 
        }
        
        Write-Host ($headerFormat -f $assetTag, $name, $model, $serial, $status) -ForegroundColor White
    }
    
    Write-Host ("-" * $windowWidth) -ForegroundColor Gray
    Write-Host "Total: $($assets.Count) assets" -ForegroundColor Green
}

# Helper function to export assets to CSV
function Export-AssetsToCSV
{
    param (
        [Parameter(Mandatory = $true)]
        [array]$assets,
        
        [Parameter(Mandatory = $false)]
        [string]$fileNameSuffix = "Export"
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\SnipeIT_Assets_${fileNameSuffix}_${timestamp}.csv"
    
    $savePath = Read-Host "Enter path to save CSV file (default: $defaultPath)"
    if ([string]::IsNullOrWhiteSpace($savePath))
    {
        $savePath = $defaultPath
    }
    
    try
    {
        # Create a custom object array with selected properties
        $exportData = $assets | ForEach-Object {
            $assetObj = [PSCustomObject]@{
                'Asset Tag' = $_.asset_tag
                'Name' = $_.name
                'Model' = $_.model.name
                'Category' = $_.category.name
                'Manufacturer' = $_.manufacturer.name
                'Serial' = $_.serial
                'Status' = $_.status_label.name
                'Location' = if ($_.rtd_location)
                { $_.rtd_location.name 
                } else
                { "N/A" 
                }
                'Assigned To' = if ($_.assigned_to)
                { $_.assigned_to.name 
                } else
                { "N/A" 
                }
                'Assignment Type' = if ($_.assigned_to)
                { $_.assigned_to.type 
                } else
                { "N/A" 
                }
                'Purchase Date' = $_.purchase_date
                'Purchase Cost' = $_.purchase_cost
                'Last Audit' = $_.last_audit
                'Last Checkout' = $_.last_checkout
                'Last Update' = $_.updated_at
                'Created' = $_.created_at
            }
            
            # Add custom fields if they exist
            if ($_.custom_fields)
            {
                foreach ($field in $_.custom_fields.PSObject.Properties)
                {
                    if ($field.Value.value)
                    {
                        $assetObj | Add-Member -NotePropertyName $field.Name -NotePropertyValue $field.Value.value
                    }
                }
            }
            
            return $assetObj
        }
        
        # Export to CSV
        $exportData | Export-Csv -Path $savePath -NoTypeInformation
        
        Write-Host "File successfully exported to: $savePath" -ForegroundColor Green
    } catch
    {
        Write-Host "Error exporting to CSV: $_" -ForegroundColor Red
    }
}

# Function to view a user's assets in Snipe-IT
function View-UserAssets
{
    Clear-Host
    Write-Host "===== View User Assets =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Find the user
        $searchTerm = Read-Host "Enter username, email, or employee number to search for the user (or 'q' to quit)"
        if ($searchTerm -eq 'q')
        {
            return
        }
        
        Write-Host "Searching for user..." -ForegroundColor Yellow
        $users = Get-SnipeitUser -search $searchTerm
        
        if (!$users -or $users.Count -eq 0)
        {
            Write-Host "No users found with the search term: $searchTerm" -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }
        
        # Step 2: Select the user if multiple found
        $selectedUser = $null
        
        if ($users.Count -gt 1)
        {
            Write-Host "`nMultiple users found:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $users.Count; $i++)
            {
                Write-Host "[$i] $($users[$i].name) ($($users[$i].username)) - $($users[$i].email)" -ForegroundColor Cyan
            }
            
            $userIndex = Read-Host "Enter the number of the user to view assets (or 'q' to quit)"
            if ($userIndex -eq 'q')
            {
                return
            }
            
            if ([int]::TryParse($userIndex, [ref]$null) -and [int]$userIndex -ge 0 -and [int]$userIndex -lt $users.Count)
            {
                $selectedUser = $users[[int]$userIndex]
            } else
            {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
                return
            }
        } else
        {
            $selectedUser = $users
        }
        
        # Step 3: Get user's assets
        Clear-Host
        Write-Host "===== Assets for $($selectedUser.name) =====" -ForegroundColor Cyan
        Write-Host "User ID: $($selectedUser.id)" -ForegroundColor Yellow
        Write-Host "Username: $($selectedUser.username)" -ForegroundColor Yellow
        Write-Host "Email: $($selectedUser.email)" -ForegroundColor Yellow
        
        Write-Host "`nRetrieving assets..." -ForegroundColor Green
        
        # Use the user_id parameter to get assets assigned to this user
        $assets = Get-SnipeitAsset -user_id $selectedUser.id
        
        if (!$assets -or $assets.Count -eq 0)
        {
            Write-Host "No assets currently assigned to this user." -ForegroundColor Yellow
        } else
        {
            # Display assets in a tabular format
            Write-Host "`nAssets assigned to $($selectedUser.name):" -ForegroundColor Green
            Write-Host "-----------------------------------------------------------" -ForegroundColor Gray
            Write-Host "| Asset Tag | Name | Model | Category | Status | Checkout Date |" -ForegroundColor Cyan
            Write-Host "-----------------------------------------------------------" -ForegroundColor Gray
            
            foreach ($asset in $assets)
            {
                $assetTag = $asset.asset_tag
                $name = if ($asset.name)
                { $asset.name 
                } else
                { "N/A" 
                }
                $model = if ($asset.model)
                { $asset.model.name 
                } else
                { "N/A" 
                }
                $category = if ($asset.category)
                { $asset.category.name 
                } else
                { "N/A" 
                }
                $status = if ($asset.status_label)
                { $asset.status_label.name 
                } else
                { "N/A" 
                }
                $checkoutDate = if ($asset.last_checkout)
                { 
                    # Handle different date formats that might be returned
                    if ($asset.last_checkout -is [string])
                    {
                        try
                        { $asset.last_checkout.Substring(0, 10) 
                        } catch
                        { "N/A" 
                        }
                    } elseif ($asset.last_checkout -is [DateTime])
                    {
                        $asset.last_checkout.ToString("yyyy-MM-dd")
                    } else
                    {
                        "N/A"
                    }
                } else
                { "N/A" 
                }
                
                Write-Host "| $assetTag | $name | $model | $category | $status | $checkoutDate |" -ForegroundColor White
            }
            Write-Host "-----------------------------------------------------------" -ForegroundColor Gray
            Write-Host "Total Assets: $($assets.Count)" -ForegroundColor Green
        }
        
        # Additional options
        Write-Host "`nOptions:" -ForegroundColor Cyan
        Write-Host "[1] View asset details" -ForegroundColor Yellow
        Write-Host "[2] Check in an asset" -ForegroundColor Yellow
        Write-Host "[q] Return to menu" -ForegroundColor Yellow
        
        $option = Read-Host "Select an option"
        
        switch ($option)
        {
            "1"
            {
                if (!$assets -or $assets.Count -eq 0)
                {
                    Write-Host "No assets available to view." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    break
                }
                
                $assetTag = Read-Host "Enter asset tag to view details"
                $selectedAsset = $assets | Where-Object { $_.asset_tag -eq $assetTag }
                
                if (!$selectedAsset)
                {
                    Write-Host "Asset not found." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    break
                }
                
                # Display detailed asset info
                Clear-Host
                Write-Host "===== Asset Details =====" -ForegroundColor Cyan
                Write-Host "Asset Tag: $($selectedAsset.asset_tag)" -ForegroundColor Yellow
                Write-Host "Name: $($selectedAsset.name)" -ForegroundColor Yellow
                Write-Host "Model: $($selectedAsset.model.name)" -ForegroundColor Yellow
                Write-Host "Category: $($selectedAsset.category.name)" -ForegroundColor Yellow
                Write-Host "Status: $($selectedAsset.status_label.name)" -ForegroundColor Yellow
                Write-Host "Serial: $($selectedAsset.serial)" -ForegroundColor Yellow
                
                if ($selectedAsset.purchase_date)
                {
                    $purchaseDate = if ($selectedAsset.purchase_date -is [string])
                    {
                        try
                        { $selectedAsset.purchase_date.Substring(0, 10) 
                        } catch
                        { $selectedAsset.purchase_date 
                        }
                    } elseif ($selectedAsset.purchase_date -is [DateTime])
                    {
                        $selectedAsset.purchase_date.ToString("yyyy-MM-dd")
                    } else
                    {
                        $selectedAsset.purchase_date
                    }
                    Write-Host "Purchase Date: $purchaseDate" -ForegroundColor Yellow
                }
                
                if ($selectedAsset.last_checkout)
                {
                    $checkoutDate = if ($selectedAsset.last_checkout -is [string])
                    {
                        try
                        { $selectedAsset.last_checkout.Substring(0, 10) 
                        } catch
                        { $selectedAsset.last_checkout 
                        }
                    } elseif ($selectedAsset.last_checkout -is [DateTime])
                    {
                        $selectedAsset.last_checkout.ToString("yyyy-MM-dd")
                    } else
                    {
                        $selectedAsset.last_checkout
                    }
                    Write-Host "Last Checkout: $checkoutDate" -ForegroundColor Yellow
                }
                
                if ($selectedAsset.notes)
                {
                    Write-Host "Notes: $($selectedAsset.notes)" -ForegroundColor Yellow
                }
                
                # Display custom fields if any
                if ($selectedAsset.custom_fields -and $selectedAsset.custom_fields.Count -gt 0)
                {
                    Write-Host "`nCustom Fields:" -ForegroundColor Green
                    foreach ($field in $selectedAsset.custom_fields.PSObject.Properties)
                    {
                        if ($field.Value.value)
                        {
                            Write-Host "$($field.Name): $($field.Value.value)" -ForegroundColor Yellow
                        }
                    }
                }
                
                Write-Host "`nPress any key to continue..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "2"
            {
                if (!$assets -or $assets.Count -eq 0)
                {
                    Write-Host "No assets available to check in." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    break
                }
                
                $assetTag = Read-Host "Enter asset tag to check in"
                $selectedAsset = $assets | Where-Object { $_.asset_tag -eq $assetTag }
                
                if (!$selectedAsset)
                {
                    Write-Host "Asset not found." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    break
                }
                
                $notes = Read-Host "Enter check-in notes (optional)"
                
                $confirmation = Read-Host "Are you sure you want to check in this asset? (Y/N)"
                if ($confirmation -ne "Y" -and $confirmation -ne "y")
                {
                    Write-Host "Check-in cancelled." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                    break
                }
                
                # Check in the asset
                $checkInParams = @{
                    id = $selectedAsset.id
                }
                
                if (![string]::IsNullOrWhiteSpace($notes))
                {
                    $checkInParams.Add("note", $notes)
                }
                
                try
                {
                    $checkInResult = Reset-SnipeitAssetOwner @checkInParams
                    if ($checkInResult)
                    {
                        Write-Host "Asset checked in successfully!" -ForegroundColor Green
                    } else
                    {
                        Write-Host "Failed to check in asset." -ForegroundColor Red
                    }
                } catch
                {
                    Write-Host "An error occurred during check-in: $_" -ForegroundColor Red
                }
                
                Start-Sleep -Seconds 2
            }
            "q"
            {
                # Return to menu
                return
            }
            default
            {
                Write-Host "Invalid option." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } catch
    {
        Write-Host "An error occurred: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to bulk import users into Snipe-IT from a CSV file
function Import-BulkUsers
{
    Clear-Host
    Write-Host "===== Bulk Import Users =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Select CSV file
        Write-Host "This function will import multiple users from a CSV file." -ForegroundColor Yellow
        Write-Host "The CSV file should have the following headers:" -ForegroundColor Yellow
        Write-Host "first_name, last_name, username, email (required)" -ForegroundColor Yellow
        Write-Host "Optional headers: password, phone, jobtitle, employee_num, department_id, location_id, notes" -ForegroundColor Yellow
        
        # Ask for CSV path
        $csvPath = Read-Host "`nEnter the full path to the CSV file (or 'q' to quit)"
        if ($csvPath -eq 'q')
        {
            return
        }
        
        # Check if file exists
        if (-not (Test-Path $csvPath))
        {
            Write-Host "File not found: $csvPath" -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
        
        # Step 2: Parse the CSV
        Write-Host "`nReading CSV file..." -ForegroundColor Green
        try
        {
            $users = Import-Csv -Path $csvPath
        } catch
        {
            Write-Host "Error reading CSV file: $_" -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
        
        # Step 3: Validate the CSV data
        $requiredHeaders = @('first_name', 'last_name', 'username', 'email')
        $csvHeaders = $users[0].PSObject.Properties.Name
        
        $missingHeaders = $requiredHeaders | Where-Object { $_ -notin $csvHeaders }
        if ($missingHeaders.Count -gt 0)
        {
            Write-Host "Error: CSV is missing required headers: $($missingHeaders -join ', ')" -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
        
        # Step 4: Display preview of the data
        Write-Host "`nPreview of CSV data (first 5 records or less):" -ForegroundColor Green
        $previewCount = [Math]::Min(5, $users.Count)
        
        for ($i = 0; $i -lt $previewCount; $i++)
        {
            Write-Host "Record $($i+1):" -ForegroundColor Yellow
            Write-Host "  First Name: $($users[$i].first_name)" -ForegroundColor White
            Write-Host "  Last Name: $($users[$i].last_name)" -ForegroundColor White
            Write-Host "  Username: $($users[$i].username)" -ForegroundColor White
            Write-Host "  Email: $($users[$i].email)" -ForegroundColor White
            Write-Host "  ----------------------------" -ForegroundColor Gray
        }
        
        Write-Host "`nTotal records in CSV: $($users.Count)" -ForegroundColor Green
        
        # Step 5: Confirm import
        $confirmation = Read-Host "`nDo you want to import these users? (Y/N)"
        if ($confirmation -ne "Y" -and $confirmation -ne "y")
        {
            Write-Host "Import cancelled." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return
        }
        
        # Step 6: Process each user
        $successCount = 0
        $errorCount = 0
        $errors = @()
        
        Write-Host "`nImporting users..." -ForegroundColor Green
        
        # Create a progress bar
        $progressParams = @{
            Activity = "Importing Users"
            Status = "Processing users..."
            PercentComplete = 0
        }
        
        for ($i = 0; $i -lt $users.Count; $i++)
        {
            $user = $users[$i]
            
            # Update progress
            $progressParams.PercentComplete = [Math]::Round(($i / $users.Count) * 100)
            $progressParams.Status = "Processing user $($i+1) of $($users.Count): $($user.username)"
            Write-Progress @progressParams
            
            # Build parameters for New-SnipeitUser
            $userParams = @{
                first_name = $user.first_name
                last_name = $user.last_name
                username = $user.username
                email = $user.email
            }
            
            # Add optional parameters if present
            $optionalParams = @('password', 'phone', 'jobtitle', 'employee_num', 'department_id', 'location_id', 'notes')
            
            foreach ($param in $optionalParams)
            {
                if ($user.PSObject.Properties.Name -contains $param -and ![string]::IsNullOrWhiteSpace($user.$param))
                {
                    $userParams.Add($param, $user.$param)
                }
            }
            
            # Create the user
            try
            {
                $newUser = New-SnipeitUser @userParams
                if ($newUser)
                {
                    $successCount++
                } else
                {
                    $errorCount++
                    $errors += "User $($user.username): Unknown error"
                }
            } catch
            {
                $errorCount++
                $errors += "User $($user.username): $($_.Exception.Message)"
            }
            
            # Add a small delay to prevent API throttling
            Start-Sleep -Milliseconds 200
        }
        
        # Complete the progress bar
        Write-Progress -Activity "Importing Users" -Completed
        
        # Step 7: Display results
        Clear-Host
        Write-Host "===== Import Results =====" -ForegroundColor Cyan
        Write-Host "Total users processed: $($users.Count)" -ForegroundColor White
        Write-Host "Successfully imported: $successCount" -ForegroundColor Green
        Write-Host "Errors: $errorCount" -ForegroundColor $(if ($errorCount -gt 0)
            { "Red" 
            } else
            { "Green" 
            })
        
        if ($errorCount -gt 0)
        {
            Write-Host "`nError Details:" -ForegroundColor Yellow
            foreach ($error in $errors)
            {
                Write-Host "- $error" -ForegroundColor Red
            }
        }
        
        # Offer to create a template CSV
        Write-Host "`nWould you like to:" -ForegroundColor Cyan
        Write-Host "1. Create a template CSV file for future imports" -ForegroundColor Yellow
        Write-Host "2. Return to menu" -ForegroundColor Yellow
        
        $option = Read-Host "Select an option"
        
        if ($option -eq "1")
        {
            $templatePath = Read-Host "Enter a path to save the template (e.g., C:\Users\YourName\Desktop\user_template.csv)"
            
            if ([string]::IsNullOrWhiteSpace($templatePath))
            {
                $templatePath = Join-Path -Path $env:TEMP -ChildPath "snipeit_user_template.csv"
            }
            
            $template = @"
first_name,last_name,username,email,password,phone,jobtitle,employee_num,department_id,location_id,notes
John,Doe,jdoe,john.doe@example.com,Password123,555-1234,IT Manager,EMP001,1,1,"Example notes"
Jane,Smith,jsmith,jane.smith@example.com,Password456,555-5678,Developer,EMP002,2,1,"Another example"
"@
            
            try
            {
                $template | Out-File -FilePath $templatePath -Encoding UTF8
                Write-Host "Template saved to: $templatePath" -ForegroundColor Green
            } catch
            {
                Write-Host "Error saving template: $_" -ForegroundColor Red
            }
        }
    } catch
    {
        Write-Host "An unexpected error occurred: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to generate a user import template CSV
function New-UserImportTemplate
{
    Clear-Host
    Write-Host "===== Generate User Import Template =====" -ForegroundColor Cyan
    
    try
    {
        # Get the save location
        $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\snipeit_user_template.csv"
        Write-Host "This will create a template CSV file for importing users into Snipe-IT." -ForegroundColor Yellow
        $savePath = Read-Host "Enter a path to save the template (default: $defaultPath)"
        
        if ([string]::IsNullOrWhiteSpace($savePath))
        {
            $savePath = $defaultPath
        }
        
        # Create the template content
        $templateContent = @"
first_name,last_name,username,email,password,phone,jobtitle,employee_num,department_id,location_id,notes
John,Doe,jdoe,john.doe@example.com,Password123,555-1234,IT Manager,EMP001,1,1,"Example notes"
Jane,Smith,jsmith,jane.smith@example.com,Password456,555-5678,Developer,EMP002,2,1,"Another example"
"@
        
        # Save the template
        try
        {
            $templateContent | Out-File -FilePath $savePath -Encoding UTF8
            Write-Host "`nTemplate saved successfully to: $savePath" -ForegroundColor Green
            
            # Provide additional information about department and location IDs
            Write-Host "`nTemplate Information:" -ForegroundColor Cyan
            Write-Host "- Required fields: first_name, last_name, username, email" -ForegroundColor Yellow
            Write-Host "- The password field is optional. If omitted, a random password will be generated." -ForegroundColor Yellow
            Write-Host "- For department_id and location_id fields, you need to use the numeric IDs from Snipe-IT." -ForegroundColor Yellow
            
            # Option to list available departments and locations
            $showDetails = Read-Host "`nWould you like to view available department and location IDs? (Y/N)"
            
            if ($showDetails -eq "Y" -or $showDetails -eq "y")
            {
                # Get and display departments
                try
                {
                    Write-Host "`nRetrieving departments..." -ForegroundColor Green
                    $departments = Get-SnipeitDepartment
                    
                    if ($departments -and $departments.Count -gt 0)
                    {
                        Write-Host "`nAvailable Departments:" -ForegroundColor Cyan
                        Write-Host "ID | Name" -ForegroundColor Yellow
                        Write-Host "-------------------------------------------" -ForegroundColor Gray
                        
                        foreach ($dept in $departments)
                        {
                            Write-Host "$($dept.id) | $($dept.name)" -ForegroundColor White
                        }
                    } else
                    {
                        Write-Host "No departments found." -ForegroundColor Yellow
                    }
                } catch
                {
                    Write-Host "Error retrieving departments: $_" -ForegroundColor Red
                }
                
                # Get and display locations
                try
                {
                    Write-Host "`nRetrieving locations..." -ForegroundColor Green
                    $locations = Get-SnipeitLocation
                    
                    if ($locations -and $locations.Count -gt 0)
                    {
                        Write-Host "`nAvailable Locations:" -ForegroundColor Cyan
                        Write-Host "ID | Name" -ForegroundColor Yellow
                        Write-Host "-------------------------------------------" -ForegroundColor Gray
                        
                        foreach ($loc in $locations)
                        {
                            Write-Host "$($loc.id) | $($loc.name)" -ForegroundColor White
                        }
                    } else
                    {
                        Write-Host "No locations found." -ForegroundColor Yellow
                    }
                } catch
                {
                    Write-Host "Error retrieving locations: $_" -ForegroundColor Red
                }
            }
            
            # Option to open the file
            $openFile = Read-Host "`nWould you like to open the template file? (Y/N)"
            
            if ($openFile -eq "Y" -or $openFile -eq "y")
            {
                try
                {
                    Invoke-Item -Path $savePath
                } catch
                {
                    Write-Host "Unable to open the file: $_" -ForegroundColor Red
                }
            }
        } catch
        {
            Write-Host "Error saving template: $_" -ForegroundColor Red
        }
    } catch
    {
        Write-Host "An error occurred: $_" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
# Function to update existing user in Snipe-IT
function Update-SnipeITUserInfo
{
    Clear-Host
    Write-Host "===== Update User Information =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Find the user
        $searchTerm = Read-Host "Enter username, email, or employee number to search for the user (or 'q' to quit)"
        if ($searchTerm -eq 'q')
        {
            return
        }
        
        Write-Host "Searching for user..." -ForegroundColor Yellow
        $users = Get-SnipeitUser -search $searchTerm
        
        if (!$users -or $users.Count -eq 0)
        {
            Write-Host "No users found with the search term: $searchTerm" -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }
        
        # Step 2: Select the user if multiple found
        $selectedUser = $null
        
        if ($users.Count -gt 1)
        {
            Write-Host "`nMultiple users found:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $users.Count; $i++)
            {
                Write-Host "[$i] $($users[$i].name) ($($users[$i].username)) - $($users[$i].email)" -ForegroundColor Cyan
            }
            
            $userIndex = Read-Host "Enter the number of the user to update (or 'q' to quit)"
            if ($userIndex -eq 'q')
            {
                return
            }
            
            if ([int]::TryParse($userIndex, [ref]$null) -and [int]$userIndex -ge 0 -and [int]$userIndex -lt $users.Count)
            {
                $selectedUser = $users[[int]$userIndex]
            } else
            {
                Write-Host "Invalid selection." -ForegroundColor Red
                Start-Sleep -Seconds 2
                return
            }
        } else
        {
            $selectedUser = $users
        }
        
        # Step 3: Display current user info
        Clear-Host
        Write-Host "===== Update User Information =====" -ForegroundColor Cyan
        Write-Host "Current User Information:" -ForegroundColor Green
        Write-Host "ID: $($selectedUser.id)" -ForegroundColor Yellow
        Write-Host "First Name: $($selectedUser.first_name)" -ForegroundColor Yellow
        Write-Host "Last Name: $($selectedUser.last_name)" -ForegroundColor Yellow
        Write-Host "Username: $($selectedUser.username)" -ForegroundColor Yellow
        Write-Host "Email: $($selectedUser.email)" -ForegroundColor Yellow
        
        if ($selectedUser.phone)
        {
            Write-Host "Phone: $($selectedUser.phone)" -ForegroundColor Yellow
        } else
        {
            Write-Host "Phone: Not set" -ForegroundColor Yellow
        }
        
        if ($selectedUser.jobtitle)
        {
            Write-Host "Job Title: $($selectedUser.jobtitle)" -ForegroundColor Yellow
        } else
        {
            Write-Host "Job Title: Not set" -ForegroundColor Yellow
        }
        
        if ($selectedUser.employee_num)
        {
            Write-Host "Employee Number: $($selectedUser.employee_num)" -ForegroundColor Yellow
        } else
        {
            Write-Host "Employee Number: Not set" -ForegroundColor Yellow
        }
        
        if ($selectedUser.department)
        {
            Write-Host "Department: $($selectedUser.department.name)" -ForegroundColor Yellow
        } else
        {
            Write-Host "Department: Not set" -ForegroundColor Yellow
        }
        
        if ($selectedUser.location)
        {
            Write-Host "Location: $($selectedUser.location.name)" -ForegroundColor Yellow
        } else
        {
            Write-Host "Location: Not set" -ForegroundColor Yellow
        }
        
        # Step 4: Collect updated information
        Write-Host "`nEnter new information (leave blank to keep current value):" -ForegroundColor Green
        
        # Initialize parameters object
        $updateParams = @{
            id = $selectedUser.id
        }
        
        # Collect updated values
        $first_name = Read-Host "First Name [$($selectedUser.first_name)]"
        if (![string]::IsNullOrWhiteSpace($first_name))
        {
            $updateParams.Add("first_name", $first_name)
        }
        
        $last_name = Read-Host "Last Name [$($selectedUser.last_name)]"
        if (![string]::IsNullOrWhiteSpace($last_name))
        {
            $updateParams.Add("last_name", $last_name)
        }
        
        $username = Read-Host "Username [$($selectedUser.username)]"
        if (![string]::IsNullOrWhiteSpace($username))
        {
            $updateParams.Add("username", $username)
        }
        
        $email = Read-Host "Email [$($selectedUser.email)]"
        if (![string]::IsNullOrWhiteSpace($email))
        {
            $updateParams.Add("email", $email)
        }
        
        $phone = Read-Host "Phone [$($selectedUser.phone)]"
        if (![string]::IsNullOrWhiteSpace($phone))
        {
            $updateParams.Add("phone", $phone)
        }
        
        $jobtitle = Read-Host "Job Title [$($selectedUser.jobtitle)]"
        if (![string]::IsNullOrWhiteSpace($jobtitle))
        {
            $updateParams.Add("jobtitle", $jobtitle)
        }
        
        $employee_num = Read-Host "Employee Number [$($selectedUser.employee_num)]"
        if (![string]::IsNullOrWhiteSpace($employee_num))
        {
            $updateParams.Add("employee_num", $employee_num)
        }
        
        # Get departments for selection
        $updateDepartment = Read-Host "Update department? (Y/N)"
        if ($updateDepartment -eq "Y" -or $updateDepartment -eq "y")
        {
            $departments = Get-SnipeitDepartment
            if ($departments -and $departments.Count -gt 0)
            {
                Write-Host "Available Departments:" -ForegroundColor Green
                for ($i = 0; $i -lt $departments.Count; $i++)
                {
                    Write-Host "[$i] $($departments[$i].name)" -ForegroundColor Yellow
                }
                $dept_selection = Read-Host "Select department number (or press Enter to clear)"
                
                if (![string]::IsNullOrWhiteSpace($dept_selection) -and 
                    [int]::TryParse($dept_selection, [ref]$null) -and 
                    [int]$dept_selection -ge 0 -and 
                    [int]$dept_selection -lt $departments.Count)
                {
                    $updateParams.Add("department_id", $departments[[int]$dept_selection].id)
                } else
                {
                    # Clear department
                    $updateParams.Add("department_id", $null)
                }
            }
        }
        
        # Get locations for selection
        $updateLocation = Read-Host "Update location? (Y/N)"
        if ($updateLocation -eq "Y" -or $updateLocation -eq "y")
        {
            $locations = Get-SnipeitLocation
            if ($locations -and $locations.Count -gt 0)
            {
                Write-Host "Available Locations:" -ForegroundColor Green
                for ($i = 0; $i -lt $locations.Count; $i++)
                {
                    Write-Host "[$i] $($locations[$i].name)" -ForegroundColor Yellow
                }
                $loc_selection = Read-Host "Select location number (or press Enter to clear)"
                
                if (![string]::IsNullOrWhiteSpace($loc_selection) -and 
                    [int]::TryParse($loc_selection, [ref]$null) -and 
                    [int]$loc_selection -ge 0 -and 
                    [int]$loc_selection -lt $locations.Count)
                {
                    $updateParams.Add("location_id", $locations[[int]$loc_selection].id)
                } else
                {
                    # Clear location
                    $updateParams.Add("location_id", $null)
                }
            }
        }
        
        # Reset password option
        $resetPassword = Read-Host "Reset password? (Y/N)"
        if ($resetPassword -eq "Y" -or $resetPassword -eq "y")
        {
            $newPassword = Read-Host "Enter new password"
            if (![string]::IsNullOrWhiteSpace($newPassword))
            {
                $updateParams.Add("password", $newPassword)
            }
        }
        
        # Step 5: Confirm and update
        if ($updateParams.Count -le 1)
        {
            Write-Host "No changes were specified." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return
        }
        
        Write-Host "`nThe following changes will be made:" -ForegroundColor Cyan
        foreach ($key in $updateParams.Keys)
        {
            if ($key -ne "id")
            {
                if ($key -eq "department_id" -and $updateParams[$key] -eq $null)
                {
                    Write-Host "Department: [Clear]" -ForegroundColor Yellow
                } elseif ($key -eq "department_id")
                {
                    $deptName = $departments | Where-Object { $_.id -eq $updateParams[$key] } | Select-Object -ExpandProperty name
                    Write-Host "Department: $deptName" -ForegroundColor Yellow
                } elseif ($key -eq "location_id" -and $updateParams[$key] -eq $null)
                {
                    Write-Host "Location: [Clear]" -ForegroundColor Yellow
                } elseif ($key -eq "location_id")
                {
                    $locName = $locations | Where-Object { $_.id -eq $updateParams[$key] } | Select-Object -ExpandProperty name
                    Write-Host "Location: $locName" -ForegroundColor Yellow
                } elseif ($key -eq "password")
                {
                    Write-Host "Password: [Reset]" -ForegroundColor Yellow
                } else
                {
                    Write-Host "$($key): $($updateParams[$key])" -ForegroundColor Yellow
                }
            }
        }
        
        $confirm = Read-Host "`nUpdate this user? (Y/N)"
        if ($confirm -ne "Y" -and $confirm -ne "y")
        {
            Write-Host "User update cancelled." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return
        }
        
        # Update the user
        Write-Host "`nUpdating user..." -ForegroundColor Green
        $updatedUser = Set-SnipeitUser @updateParams
        
        if ($updatedUser)
        {
            Write-Host "User updated successfully!" -ForegroundColor Green
            Write-Host "Updated User: $($updatedUser.name)" -ForegroundColor Yellow
        } else
        {
            Write-Host "Failed to update user." -ForegroundColor Red
        }
    } catch
    {
        Write-Host "An error occurred while updating the user: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
# Function to add a new user to Snipe-IT
function Add-SnipeITUser
{
    Clear-Host
    Write-Host "===== Add New User =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Collect user information
        Write-Host "Enter user information (press Enter to skip optional fields):" -ForegroundColor Green
        
        # Required fields
        $first_name = Read-Host "First Name (required)"
        if ([string]::IsNullOrWhiteSpace($first_name))
        {
            Write-Host "First name is required." -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }
        
        $last_name = Read-Host "Last Name (required)"
        if ([string]::IsNullOrWhiteSpace($last_name))
        {
            Write-Host "Last name is required." -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }
        
        $username = Read-Host "Username (required)"
        if ([string]::IsNullOrWhiteSpace($username))
        {
            Write-Host "Username is required." -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }
        
        $email = Read-Host "Email (required)"
        if ([string]::IsNullOrWhiteSpace($email))
        {
            Write-Host "Email is required." -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }
        
        # Optional fields
        $password = Read-Host "Password (leave blank to generate random)"
        $phone = Read-Host "Phone Number"
        $jobtitle = Read-Host "Job Title"
        $employee_num = Read-Host "Employee Number"
        
        # Get available departments for selection
        Write-Host "`nRetrieving departments..." -ForegroundColor Yellow
        $departments = Get-SnipeitDepartment
        if ($departments -and $departments.Count -gt 0)
        {
            Write-Host "Available Departments:" -ForegroundColor Green
            for ($i = 0; $i -lt $departments.Count; $i++)
            {
                Write-Host "[$i] $($departments[$i].name)" -ForegroundColor Yellow
            }
            $dept_selection = Read-Host "Select department number (press Enter to skip)"
            
            if (![string]::IsNullOrWhiteSpace($dept_selection) -and 
                [int]::TryParse($dept_selection, [ref]$null) -and 
                [int]$dept_selection -ge 0 -and 
                [int]$dept_selection -lt $departments.Count)
            {
                $department_id = $departments[[int]$dept_selection].id
            }
        } else
        {
            Write-Host "No departments found." -ForegroundColor Yellow
        }
        
        # Get available locations for selection
        Write-Host "`nRetrieving locations..." -ForegroundColor Yellow
        $locations = Get-SnipeitLocation
        if ($locations -and $locations.Count -gt 0)
        {
            Write-Host "Available Locations:" -ForegroundColor Green
            for ($i = 0; $i -lt $locations.Count; $i++)
            {
                Write-Host "[$i] $($locations[$i].name)" -ForegroundColor Yellow
            }
            $loc_selection = Read-Host "Select location number (press Enter to skip)"
            
            if (![string]::IsNullOrWhiteSpace($loc_selection) -and 
                [int]::TryParse($loc_selection, [ref]$null) -and 
                [int]$loc_selection -ge 0 -and 
                [int]$loc_selection -lt $locations.Count)
            {
                $location_id = $locations[[int]$loc_selection].id
            }
        } else
        {
            Write-Host "No locations found." -ForegroundColor Yellow
        }
        
        # Prepare parameters for new user
        $userParams = @{
            first_name = $first_name
            last_name = $last_name
            username = $username
            email = $email
        }
        
        # Add optional parameters if provided
        if (![string]::IsNullOrWhiteSpace($password))
        {
            $userParams.Add("password", $password)
        }
        
        if (![string]::IsNullOrWhiteSpace($phone))
        {
            $userParams.Add("phone", $phone)
        }
        
        if (![string]::IsNullOrWhiteSpace($jobtitle))
        {
            $userParams.Add("jobtitle", $jobtitle)
        }
        
        if (![string]::IsNullOrWhiteSpace($employee_num))
        {
            $userParams.Add("employee_num", $employee_num)
        }
        
        if ($department_id)
        {
            $userParams.Add("department_id", $department_id)
        }
        
        if ($location_id)
        {
            $userParams.Add("location_id", $location_id)
        }
        
        # Show summary and confirm
        Clear-Host
        Write-Host "===== User Summary =====" -ForegroundColor Cyan
        Write-Host "First Name: $first_name" -ForegroundColor Yellow
        Write-Host "Last Name: $last_name" -ForegroundColor Yellow
        Write-Host "Username: $username" -ForegroundColor Yellow
        Write-Host "Email: $email" -ForegroundColor Yellow
        
        if (![string]::IsNullOrWhiteSpace($phone))
        {
            Write-Host "Phone: $phone" -ForegroundColor Yellow
        }
        
        if (![string]::IsNullOrWhiteSpace($jobtitle))
        {
            Write-Host "Job Title: $jobtitle" -ForegroundColor Yellow
        }
        
        if (![string]::IsNullOrWhiteSpace($employee_num))
        {
            Write-Host "Employee Number: $employee_num" -ForegroundColor Yellow
        }
        
        if ($department_id)
        {
            Write-Host "Department: $($departments[[int]$dept_selection].name)" -ForegroundColor Yellow
        }
        
        if ($location_id)
        {
            Write-Host "Location: $($locations[[int]$loc_selection].name)" -ForegroundColor Yellow
        }
        
        $confirmation = Read-Host "`nCreate this user? (Y/N)"
        if ($confirmation -ne "Y" -and $confirmation -ne "y")
        {
            Write-Host "User creation cancelled." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return
        }
        
        # Create the user
        Write-Host "`nCreating user..." -ForegroundColor Green
        $newUser = New-SnipeitUser @userParams
        
        if ($newUser)
        {
            Write-Host "User created successfully!" -ForegroundColor Green
            Write-Host "User ID: $($newUser.id)" -ForegroundColor Yellow
            Write-Host "Full Name: $($newUser.name)" -ForegroundColor Yellow
            Write-Host "Email: $($newUser.email)" -ForegroundColor Yellow
        } else
        {
            Write-Host "Failed to create user." -ForegroundColor Red
        }
    } catch
    {
        Write-Host "An error occurred while creating the user: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}




# Function to generate an Asset Audit Report
function Generate-AssetAuditReport
{
    Clear-Host
    Write-Host "===== Asset Audit Report =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Define report parameters
        Write-Host "This report shows the audit status of assets in your inventory." -ForegroundColor Yellow
        Write-Host "Choose report parameters:" -ForegroundColor Green
        
        # Get audit interval setting - use default of 12 months if not specified
        $auditIntervalMonths = Read-Host "Enter audit interval in months (default: 12)"
        if ([string]::IsNullOrWhiteSpace($auditIntervalMonths) -or -not [int]::TryParse($auditIntervalMonths, [ref]$null))
        {
            $auditIntervalMonths = 12
        }
        
        # Choose asset types to include
        Write-Host "`nSelect assets to include:" -ForegroundColor Yellow
        Write-Host "1. All assets" -ForegroundColor Cyan
        Write-Host "2. Deployed assets only" -ForegroundColor Cyan
        Write-Host "3. Ready-to-deploy assets only" -ForegroundColor Cyan
        Write-Host "4. Assets by category" -ForegroundColor Cyan
        Write-Host "5. Assets by location" -ForegroundColor Cyan
        
        $assetFilter = Read-Host "Enter option (1-5)"
        
        # Additional filter parameters based on selection
        $filterParams = @{}
        $filterDescription = "All assets"
        
        switch ($assetFilter)
        {
            "2"
            {
                # Get deployed status ID
                $deployedStatusId = (Get-SnipeitStatus | Where-Object { $_.name -match "Deployed|Assigned|Checked Out" } | Select-Object -First 1).id
                if ($deployedStatusId)
                {
                    $filterParams.Add("status_id", $deployedStatusId)
                    $filterDescription = "Deployed assets only"
                } else
                {
                    Write-Host "Could not determine deployed status ID. Including all assets." -ForegroundColor Yellow
                }
            }
            "3"
            {
                # Get ready-to-deploy status ID
                $readyStatusId = (Get-SnipeitStatus | Where-Object { $_.name -match "Ready|Available" } | Select-Object -First 1).id
                if ($readyStatusId)
                {
                    $filterParams.Add("status_id", $readyStatusId)
                    $filterDescription = "Ready-to-deploy assets only"
                } else
                {
                    Write-Host "Could not determine ready status ID. Including all assets." -ForegroundColor Yellow
                }
            }
            "4"
            {
                # Filter by category
                $categories = Get-SnipeitCategory
                if ($categories)
                {
                    Write-Host "`nAvailable Categories:" -ForegroundColor Yellow
                    for ($i = 0; $i -lt $categories.Count; $i++)
                    {
                        Write-Host "[$i] $($categories[$i].name)" -ForegroundColor Cyan
                    }
                    
                    $categoryIndex = Read-Host "Select category number"
                    if ([int]::TryParse($categoryIndex, [ref]$null) -and [int]$categoryIndex -ge 0 -and [int]$categoryIndex -lt $categories.Count)
                    {
                        $selectedCategory = $categories[[int]$categoryIndex]
                        $filterParams.Add("category_id", $selectedCategory.id)
                        $filterDescription = "Assets in category: $($selectedCategory.name)"
                    } else
                    {
                        Write-Host "Invalid selection. Including all assets." -ForegroundColor Yellow
                    }
                } else
                {
                    Write-Host "Could not retrieve categories. Including all assets." -ForegroundColor Yellow
                }
            }
            "5"
            {
                # Filter by location
                $locations = Get-SnipeitLocation
                if ($locations)
                {
                    Write-Host "`nAvailable Locations:" -ForegroundColor Yellow
                    for ($i = 0; $i -lt $locations.Count; $i++)
                    {
                        Write-Host "[$i] $($locations[$i].name)" -ForegroundColor Cyan
                    }
                    
                    $locationIndex = Read-Host "Select location number"
                    if ([int]::TryParse($locationIndex, [ref]$null) -and [int]$locationIndex -ge 0 -and [int]$locationIndex -lt $locations.Count)
                    {
                        $selectedLocation = $locations[[int]$locationIndex]
                        $filterParams.Add("location_id", $selectedLocation.id)
                        $filterDescription = "Assets at location: $($selectedLocation.name)"
                    } else
                    {
                        Write-Host "Invalid selection. Including all assets." -ForegroundColor Yellow
                    }
                } else
                {
                    Write-Host "Could not retrieve locations. Including all assets." -ForegroundColor Yellow
                }
            }
        }
        
        # Step 2: Retrieve assets based on filters
        Write-Host "`nRetrieving assets for the report..." -ForegroundColor Green
        
        # Get all assets matching the filter
        $assets = if ($filterParams.Count -gt 0)
        {
            try
            {
                # Try with splatting
                Get-SnipeitAsset @filterParams -all
            } catch
            {
                Write-Host "Error applying filters. Getting all assets instead." -ForegroundColor Yellow
                Get-SnipeitAsset -all
            }
        } else
        {
            Get-SnipeitAsset -all
        }
        
        if (-not $assets -or $assets.Count -eq 0)
        {
            Write-Host "No assets found with the specified criteria." -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
        
        # Step 3: Analyze audit status
        Write-Host "`nAnalyzing audit status of $($assets.Count) assets..." -ForegroundColor Green
        
        $today = Get-Date
        $auditIntervalDays = [int]$auditIntervalMonths * 30
        $cutoffDate = $today.AddDays(-$auditIntervalDays)
        
        # Categorize assets by audit status
        $assetsNeverAudited = @()
        $assetsOverdue = @()
        $assetsDueSoon = @()
        $assetsCompliant = @()
        
        foreach ($asset in $assets)
        {
            if (-not $asset.last_audit)
            {
                # Asset has never been audited
                $assetsNeverAudited += $asset
            } else
            {
                # Parse the last audit date
                try
                {
                    $lastAuditDate = if ($asset.last_audit -is [string])
                    {
                        [DateTime]::Parse($asset.last_audit)
                    } else
                    {
                        $asset.last_audit
                    }
                    
                    $daysSinceAudit = ($today - $lastAuditDate).TotalDays
                    $daysUntilDue = $auditIntervalDays - $daysSinceAudit
                    
                    if ($daysSinceAudit -gt $auditIntervalDays)
                    {
                        # Audit is overdue
                        $asset | Add-Member -NotePropertyName DaysSinceAudit -NotePropertyValue ([Math]::Round($daysSinceAudit)) -Force
                        $assetsOverdue += $asset
                    } elseif ($daysUntilDue -le 30)
                    {
                        # Audit is due within 30 days
                        $asset | Add-Member -NotePropertyName DaysUntilDue -NotePropertyValue ([Math]::Round($daysUntilDue)) -Force
                        $assetsDueSoon += $asset
                    } else
                    {
                        # Asset is compliant
                        $asset | Add-Member -NotePropertyName DaysUntilDue -NotePropertyValue ([Math]::Round($daysUntilDue)) -Force
                        $assetsCompliant += $asset
                    }
                } catch
                {
                    # If date parsing fails, consider it never audited
                    $assetsNeverAudited += $asset
                }
            }
        }
        
        # Step 4: Generate summary statistics
        $totalAssets = $assets.Count
        $neverAuditedCount = $assetsNeverAudited.Count
        $overdueCount = $assetsOverdue.Count
        $dueSoonCount = $assetsDueSoon.Count
        $compliantCount = $assetsCompliant.Count
        
        $neverAuditedPercent = if ($totalAssets -gt 0)
        { [Math]::Round(($neverAuditedCount / $totalAssets) * 100, 2) 
        } else
        { 0 
        }
        $overduePercent = if ($totalAssets -gt 0)
        { [Math]::Round(($overdueCount / $totalAssets) * 100, 2) 
        } else
        { 0 
        }
        $dueSoonPercent = if ($totalAssets -gt 0)
        { [Math]::Round(($dueSoonCount / $totalAssets) * 100, 2) 
        } else
        { 0 
        }
        $compliantPercent = if ($totalAssets -gt 0)
        { [Math]::Round(($compliantCount / $totalAssets) * 100, 2) 
        } else
        { 0 
        }
        
        # Step 5: Display the report
        Clear-Host
        Write-Host "===== Asset Audit Report =====" -ForegroundColor Cyan
        Write-Host "Report Date: $($today.ToString('yyyy-MM-dd'))" -ForegroundColor Yellow
        Write-Host "Assets Included: $filterDescription" -ForegroundColor Yellow
        Write-Host "Audit Interval: $auditIntervalMonths months" -ForegroundColor Yellow
        Write-Host "Total Assets: $totalAssets" -ForegroundColor Yellow
        
        Write-Host "`nAudit Status Summary:" -ForegroundColor Green
        Write-Host "Never Audited:  $neverAuditedCount ($neverAuditedPercent%)" -ForegroundColor $(if ($neverAuditedCount -gt 0)
            { "Red" 
            } else
            { "Green" 
            })
        Write-Host "Overdue:        $overdueCount ($overduePercent%)" -ForegroundColor $(if ($overdueCount -gt 0)
            { "Red" 
            } else
            { "Green" 
            })
        Write-Host "Due Soon:       $dueSoonCount ($dueSoonPercent%)" -ForegroundColor Yellow
        Write-Host "Compliant:      $compliantCount ($compliantPercent%)" -ForegroundColor Green
        
        # Overall compliance rate
        $complianceRate = $compliantPercent
        Write-Host "`nOverall Compliance Rate: $complianceRate%" -ForegroundColor $(
            if ($complianceRate -ge 90)
            { "Green" 
            } elseif ($complianceRate -ge 70)
            { "Yellow" 
            } else
            { "Red" 
            }
        )
        
        # Step 6: Option to display detailed lists
        Write-Host "`nView detailed lists:" -ForegroundColor Cyan
        Write-Host "1. Never Audited Assets" -ForegroundColor Cyan
        Write-Host "2. Overdue Assets" -ForegroundColor Cyan
        Write-Host "3. Assets Due Soon" -ForegroundColor Cyan
        Write-Host "4. Export Full Report to CSV" -ForegroundColor Cyan
        Write-Host "5. Return to menu" -ForegroundColor Cyan
        
        $detailOption = Read-Host "Enter option (1-5)"
        
        switch ($detailOption)
        {
            "1"
            {
                if ($assetsNeverAudited.Count -gt 0)
                {
                    Write-Host "`nAssets Never Audited:" -ForegroundColor Red
                    $assetsNeverAudited | Select-Object asset_tag, name, @{Name="Model";Expression={$_.model.name}}, @{Name="Category";Expression={$_.category.name}}, serial | Format-Table -AutoSize
                    
                    $exportOption = Read-Host "Export this list to CSV? (Y/N)"
                    if ($exportOption -eq "Y" -or $exportOption -eq "y")
                    {
                        Export-AssetsToCSV -assets $assetsNeverAudited -fileNameSuffix "NeverAudited"
                    }
                } else
                {
                    Write-Host "No assets in this category." -ForegroundColor Green
                }
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-AssetAuditReport
            }
            "2"
            {
                if ($assetsOverdue.Count -gt 0)
                {
                    Write-Host "`nOverdue Assets:" -ForegroundColor Red
                    $assetsOverdue | Sort-Object DaysSinceAudit -Descending | Select-Object asset_tag, name, @{Name="Model";Expression={$_.model.name}}, @{Name="Days Since Audit";Expression={$_.DaysSinceAudit}}, @{Name="Last Audit";Expression={
                            if ($_.last_audit -is [string])
                            {
                                try
                                { [DateTime]::Parse($_.last_audit).ToString("yyyy-MM-dd") 
                                } catch
                                { "Invalid date" 
                                }
                            } elseif ($_.last_audit -is [DateTime])
                            {
                                $_.last_audit.ToString("yyyy-MM-dd")
                            } else
                            { "Unknown" 
                            }
                        }
                    } | Format-Table -AutoSize
                    
                    $exportOption = Read-Host "Export this list to CSV? (Y/N)"
                    if ($exportOption -eq "Y" -or $exportOption -eq "y")
                    {
                        Export-AssetsToCSV -assets $assetsOverdue -fileNameSuffix "AuditOverdue"
                    }
                } else
                {
                    Write-Host "No assets in this category." -ForegroundColor Green
                }
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-AssetAuditReport
            }
            "3"
            {
                if ($assetsDueSoon.Count -gt 0)
                {
                    Write-Host "`nAssets Due for Audit Soon:" -ForegroundColor Yellow
                    $assetsDueSoon | Sort-Object DaysUntilDue | Select-Object asset_tag, name, @{Name="Model";Expression={$_.model.name}}, @{Name="Days Until Due";Expression={$_.DaysUntilDue}}, @{Name="Last Audit";Expression={
                            if ($_.last_audit -is [string])
                            {
                                try
                                { [DateTime]::Parse($_.last_audit).ToString("yyyy-MM-dd") 
                                } catch
                                { "Invalid date" 
                                }
                            } elseif ($_.last_audit -is [DateTime])
                            {
                                $_.last_audit.ToString("yyyy-MM-dd")
                            } else
                            { "Unknown" 
                            }
                        }
                    } | Format-Table -AutoSize
                    
                    $exportOption = Read-Host "Export this list to CSV? (Y/N)"
                    if ($exportOption -eq "Y" -or $exportOption -eq "y")
                    {
                        Export-AssetsToCSV -assets $assetsDueSoon -fileNameSuffix "AuditDueSoon"
                    }
                } else
                {
                    Write-Host "No assets in this category." -ForegroundColor Green
                }
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-AssetAuditReport
            }
            "4"
            {
                # Create a comprehensive report object
                $reportData = @()
                
                # Add all assets with audit status
                foreach ($asset in $assets)
                {
                    $auditStatus = "Unknown"
                    $daysSinceAudit = $null
                    $daysUntilDue = $null
                    $lastAuditDate = "Never"
                    
                    if ($asset.last_audit)
                    {
                        try
                        {
                            $lastAudit = if ($asset.last_audit -is [string])
                            {
                                [DateTime]::Parse($asset.last_audit)
                            } else
                            {
                                $asset.last_audit
                            }
                            
                            $lastAuditDate = $lastAudit.ToString("yyyy-MM-dd")
                            $daysSinceAudit = [Math]::Round(($today - $lastAudit).TotalDays)
                            $daysUntilDue = [Math]::Round($auditIntervalDays - $daysSinceAudit)
                            
                            if ($daysSinceAudit -gt $auditIntervalDays)
                            {
                                $auditStatus = "Overdue"
                            } elseif ($daysUntilDue -le 30)
                            {
                                $auditStatus = "Due Soon"
                            } else
                            {
                                $auditStatus = "Compliant"
                            }
                        } catch
                        {
                            $auditStatus = "Invalid Date"
                        }
                    } else
                    {
                        $auditStatus = "Never Audited"
                    }
                    
                    # Create report object
                    $reportData += [PSCustomObject]@{
                        'Asset Tag' = $asset.asset_tag
                        'Name' = $asset.name
                        'Model' = $asset.model.name
                        'Category' = $asset.category.name
                        'Serial' = $asset.serial
                        'Status' = $asset.status_label.name
                        'Location' = if ($asset.rtd_location)
                        { $asset.rtd_location.name 
                        } else
                        { "N/A" 
                        }
                        'Assigned To' = if ($asset.assigned_to)
                        { $asset.assigned_to.name 
                        } else
                        { "N/A" 
                        }
                        'Last Audit Date' = $lastAuditDate
                        'Days Since Audit' = $daysSinceAudit
                        'Days Until Due' = $daysUntilDue
                        'Audit Status' = $auditStatus
                        'Purchase Date' = if ($asset.purchase_date)
                        {
                            if ($asset.purchase_date -is [string])
                            {
                                try
                                { [DateTime]::Parse($asset.purchase_date).ToString("yyyy-MM-dd") 
                                } catch
                                { "Invalid date" 
                                }
                            } elseif ($asset.purchase_date -is [DateTime])
                            {
                                $asset.purchase_date.ToString("yyyy-MM-dd")
                            } else
                            { "Unknown" 
                            }
                        } else
                        { "Unknown" 
                        }
                    }
                }
                
                # Export the report
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\SnipeIT_AuditReport_${timestamp}.csv"
                
                $savePath = Read-Host "Enter path to save CSV file (default: $defaultPath)"
                if ([string]::IsNullOrWhiteSpace($savePath))
                {
                    $savePath = $defaultPath
                }
                
                try
                {
                    $reportData | Export-Csv -Path $savePath -NoTypeInformation
                    Write-Host "Report exported to: $savePath" -ForegroundColor Green
                } catch
                {
                    Write-Host "Error exporting report: $_" -ForegroundColor Red
                }
            }
            "5"
            {
                return
            }
        }
    } catch
    {
        Write-Host "An error occurred while generating the report: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to generate a License Compliance Report
function Generate-LicenseComplianceReport
{
    Clear-Host
    Write-Host "===== License Compliance Report =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Retrieve all licenses
        Write-Host "Retrieving license information..." -ForegroundColor Yellow
        $licenses = Get-SnipeitLicense -all
        
        if (-not $licenses -or $licenses.Count -eq 0)
        {
            Write-Host "No licenses found in the system." -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
        
        Write-Host "Found $($licenses.Count) licenses." -ForegroundColor Green
        
        # Step 2: Analyze license utilization and compliance
        Write-Host "`nAnalyzing license utilization and compliance..." -ForegroundColor Yellow
        
        $licenseReport = @()
        $today = Get-Date
        $totalSeats = 0
        $totalUtilized = 0
        $expiredLicenseCount = 0
        $expiringLicenseCount = 0
        $overdeployedLicenseCount = 0
        
        foreach ($license in $licenses)
        {
            # Parse expiration date if present
            $expirationDate = $null
            $daysUntilExpiry = $null
            $isExpired = $false
            $isExpiring = $false
            
            if ($license.expiration_date)
            {
                try
                {
                    $expirationDate = if ($license.expiration_date -is [string])
                    {
                        [DateTime]::Parse($license.expiration_date)
                    } else
                    {
                        $license.expiration_date
                    }
                    
                    $daysUntilExpiry = [Math]::Round(($expirationDate - $today).TotalDays)
                    $isExpired = $daysUntilExpiry -lt 0
                    $isExpiring = $daysUntilExpiry -ge 0 -and $daysUntilExpiry -le 90
                    
                    if ($isExpired)
                    { $expiredLicenseCount++ 
                    }
                    if ($isExpiring)
                    { $expiringLicenseCount++ 
                    }
                } catch
                {
                    # If date parsing fails, treat as no expiration
                    $expirationDate = $null
                }
            }
            
            # Get seats information
            $totalSeatsForLicense = $license.seats
            $totalSeats += $totalSeatsForLicense
            
            # Try to get utilized seats count
            $utilizedSeats = 0
            if ($license.free_seats_count -ne $null)
            {
                $utilizedSeats = $totalSeatsForLicense - $license.free_seats_count
            } else
            {
                # If free_seats_count is not available, try to get it from license seats
                try
                {
                    $licenseSeats = Get-SnipeitLicenseSeat -license_id $license.id
                    if ($licenseSeats)
                    {
                        $utilizedSeats = $licenseSeats.Count
                    }
                } catch
                {
                    Write-Host "Could not retrieve seat information for license: $($license.name)" -ForegroundColor Yellow
                }
            }
            
            $totalUtilized += $utilizedSeats
            
            # Calculate utilization percentage
            $utilizationPercentage = if ($totalSeatsForLicense -gt 0)
            {
                [Math]::Round(($utilizedSeats / $totalSeatsForLicense) * 100, 2)
            } else
            {
                0
            }
            
            # Determine compliance status
            $isOverdeployed = $utilizedSeats -gt $totalSeatsForLicense
            if ($isOverdeployed)
            { $overdeployedLicenseCount++ 
            }
            
            $complianceStatus = if ($isExpired)
            {
                "Expired"
            } elseif ($isOverdeployed)
            {
                "Over-deployed"
            } elseif ($isExpiring)
            {
                "Expiring Soon"
            } else
            {
                "Compliant"
            }
            
            # Create report object
            $licenseReport += [PSCustomObject]@{
                'License' = $license.name
                'License Key' = if ($license.license_key)
                { 
                    # Mask key partially for security
                    $masked = if ($license.license_key.Length -gt 8)
                    {
                        $license.license_key.Substring(0, 4) + "..." + $license.license_key.Substring($license.license_key.Length - 4)
                    } else
                    {
                        "****" + $license.license_key.Substring([Math]::Max(0, $license.license_key.Length - 4))
                    }
                    $masked
                } else
                { "N/A" 
                }
                'Product' = if ($license.product)
                { $license.product.name 
                } else
                { "N/A" 
                }
                'Manufacturer' = if ($license.manufacturer)
                { $license.manufacturer.name 
                } else
                { "N/A" 
                }
                'Total Seats' = $totalSeatsForLicense
                'Used Seats' = $utilizedSeats
                'Available Seats' = $totalSeatsForLicense - $utilizedSeats
                'Utilization %' = $utilizationPercentage
                'Expiration Date' = if ($expirationDate)
                { $expirationDate.ToString("yyyy-MM-dd") 
                } else
                { "Never" 
                }
                'Days to Expiry' = $daysUntilExpiry
                'Purchase Date' = if ($license.purchase_date)
                {
                    if ($license.purchase_date -is [string])
                    {
                        try
                        { [DateTime]::Parse($license.purchase_date).ToString("yyyy-MM-dd") 
                        } catch
                        { "Invalid date" 
                        }
                    } elseif ($license.purchase_date -is [DateTime])
                    {
                        $license.purchase_date.ToString("yyyy-MM-dd")
                    } else
                    { "Unknown" 
                    }
                } else
                { "Unknown" 
                }
                'Purchase Cost' = $license.purchase_cost
                'Status' = $complianceStatus
            }
        }
        
        # Calculate overall statistics
        $totalLicenses = $licenses.Count
        $totalCompliantLicenses = $licenseReport | Where-Object { $_.Status -eq "Compliant" } | Measure-Object | Select-Object -ExpandProperty Count
        $overallUtilizationPercentage = if ($totalSeats -gt 0)
        {
            [Math]::Round(($totalUtilized / $totalSeats) * 100, 2)
        } else
        {
            0
        }
        
        # Step 3: Display summary
        Clear-Host
        Write-Host "===== License Compliance Report =====" -ForegroundColor Cyan
        Write-Host "Report Date: $($today.ToString('yyyy-MM-dd'))" -ForegroundColor Yellow
        
        Write-Host "`nSummary Statistics:" -ForegroundColor Green
        Write-Host "Total Licenses:              $totalLicenses" -ForegroundColor White
        Write-Host "Total Seats:                 $totalSeats" -ForegroundColor White
        Write-Host "Utilized Seats:              $totalUtilized" -ForegroundColor White
        Write-Host "Overall Utilization:         $overallUtilizationPercentage%" -ForegroundColor White
        Write-Host "Compliant Licenses:          $totalCompliantLicenses/$totalLicenses" -ForegroundColor $(if ($totalCompliantLicenses -eq $totalLicenses)
            { "Green" 
            } else
            { "Yellow" 
            })
        Write-Host "Expired Licenses:            $expiredLicenseCount" -ForegroundColor $(if ($expiredLicenseCount -gt 0)
            { "Red" 
            } else
            { "Green" 
            })
        Write-Host "Expiring within 90 days:     $expiringLicenseCount" -ForegroundColor $(if ($expiringLicenseCount -gt 0)
            { "Yellow" 
            } else
            { "Green" 
            })
        Write-Host "Over-deployed Licenses:      $overdeployedLicenseCount" -ForegroundColor $(if ($overdeployedLicenseCount -gt 0)
            { "Red" 
            } else
            { "Green" 
            })
        
        # Step 4: Display compliance issues
        if ($expiredLicenseCount -gt 0 -or $expiringLicenseCount -gt 0 -or $overdeployedLicenseCount -gt 0)
        {
            Write-Host "`nCompliance Issues:" -ForegroundColor Red
            
            # Show expired licenses
            if ($expiredLicenseCount -gt 0)
            {
                Write-Host "`nExpired Licenses:" -ForegroundColor Red
                $licenseReport | Where-Object { $_.Status -eq "Expired" } | 
                    Select-Object License, 'Expiration Date', 'Total Seats', 'Used Seats', 'Manufacturer' | 
                    Format-Table -AutoSize
            }
            
            # Show expiring licenses
            if ($expiringLicenseCount -gt 0)
            {
                Write-Host "`nLicenses Expiring Soon:" -ForegroundColor Yellow
                $licenseReport | Where-Object { $_.Status -eq "Expiring Soon" } | 
                    Sort-Object 'Days to Expiry' |
                    Select-Object License, 'Expiration Date', 'Days to Expiry', 'Total Seats', 'Used Seats', 'Manufacturer' | 
                    Format-Table -AutoSize
            }
            
            # Show over-deployed licenses
            if ($overdeployedLicenseCount -gt 0)
            {
                Write-Host "`nOver-deployed Licenses:" -ForegroundColor Red
                $licenseReport | Where-Object { $_.Status -eq "Over-deployed" } | 
                    Select-Object License, 'Total Seats', 'Used Seats', 'Available Seats', 'Utilization %', 'Manufacturer' | 
                    Format-Table -AutoSize
            }
        } else
        {
            Write-Host "`nNo compliance issues found. All licenses are compliant." -ForegroundColor Green
        }
        
        # Step 5: Provide options for detailed views or export
        Write-Host "`nOptions:" -ForegroundColor Cyan
        Write-Host "1. View All Licenses" -ForegroundColor Cyan
        Write-Host "2. View High Utilization Licenses (>80%)" -ForegroundColor Cyan
        Write-Host "3. View License Details by Manufacturer" -ForegroundColor Cyan
        Write-Host "4. Export Full Report to CSV" -ForegroundColor Cyan
        Write-Host "5. Return to Menu" -ForegroundColor Cyan
        
        $option = Read-Host "Select an option (1-5)"
        
        switch ($option)
        {
            "1"
            {
                # View all licenses
                Write-Host "`nAll Licenses:" -ForegroundColor Green
                $licenseReport | Sort-Object Status, License | 
                    Select-Object License, Status, 'Total Seats', 'Used Seats', 'Available Seats', 'Utilization %', 'Expiration Date' | 
                    Format-Table -AutoSize
                
                $exportOption = Read-Host "Export this list to CSV? (Y/N)"
                if ($exportOption -eq "Y" -or $exportOption -eq "y")
                {
                    Export-LicenseReportToCSV -report $licenseReport -fileNameSuffix "AllLicenses"
                }
                
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-LicenseComplianceReport
            }
            "2"
            {
                # View high utilization licenses
                $highUtilizationLicenses = $licenseReport | Where-Object { $_.'Utilization %' -gt 80 }
                
                if ($highUtilizationLicenses.Count -gt 0)
                {
                    Write-Host "`nHigh Utilization Licenses (>80%):" -ForegroundColor Yellow
                    $highUtilizationLicenses | Sort-Object 'Utilization %' -Descending | 
                        Select-Object License, 'Total Seats', 'Used Seats', 'Available Seats', 'Utilization %', 'Manufacturer' | 
                        Format-Table -AutoSize
                    
                    $exportOption = Read-Host "Export this list to CSV? (Y/N)"
                    if ($exportOption -eq "Y" -or $exportOption -eq "y")
                    {
                        Export-LicenseReportToCSV -report $highUtilizationLicenses -fileNameSuffix "HighUtilization"
                    }
                } else
                {
                    Write-Host "No licenses with high utilization found." -ForegroundColor Green
                }
                
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-LicenseComplianceReport
            }
            "3"
            {
                # View licenses by manufacturer
                $manufacturers = $licenseReport | Select-Object -ExpandProperty Manufacturer -Unique | Where-Object { $_ -ne "N/A" }
                
                if ($manufacturers.Count -gt 0)
                {
                    Write-Host "`nAvailable Manufacturers:" -ForegroundColor Green
                    for ($i = 0; $i -lt $manufacturers.Count; $i++)
                    {
                        Write-Host "[$i] $($manufacturers[$i])" -ForegroundColor Cyan
                    }
                    
                    $mfgIndex = Read-Host "Select manufacturer number"
                    if ([int]::TryParse($mfgIndex, [ref]$null) -and [int]$mfgIndex -ge 0 -and [int]$mfgIndex -lt $manufacturers.Count)
                    {
                        $selectedMfg = $manufacturers[[int]$mfgIndex]
                        $mfgLicenses = $licenseReport | Where-Object { $_.Manufacturer -eq $selectedMfg }
                        
                        Write-Host "`nLicenses for $selectedMfg`:" -ForegroundColor Green
                        $mfgLicenses | Sort-Object Status, License | 
                            Select-Object License, Status, 'Total Seats', 'Used Seats', 'Available Seats', 'Utilization %', 'Expiration Date' | 
                            Format-Table -AutoSize
                        
                        $exportOption = Read-Host "Export this list to CSV? (Y/N)"
                        if ($exportOption -eq "Y" -or $exportOption -eq "y")
                        {
                            Export-LicenseReportToCSV -report $mfgLicenses -fileNameSuffix "Manufacturer_$($selectedMfg -replace '[^a-zA-Z0-9]', '_')"
                        }
                    } else
                    {
                        Write-Host "Invalid selection." -ForegroundColor Red
                    }
                } else
                {
                    Write-Host "No manufacturer information available for licenses." -ForegroundColor Yellow
                }
                
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-LicenseComplianceReport
            }
            "4"
            {
                # Export full report to CSV
                Export-LicenseReportToCSV -report $licenseReport -fileNameSuffix "FullReport"
                
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-LicenseComplianceReport
            }
            "5"
            {
                return
            }
            default
            {
                Write-Host "Invalid option. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Generate-LicenseComplianceReport
            }
        }
    } catch
    {
        Write-Host "An error occurred while generating the report: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}


# Helper function to export license report to CSV
function Export-LicenseReportToCSV
{
    param (
        [Parameter(Mandatory = $true)]
        [array]$report,
        
        [Parameter(Mandatory = $false)]
        [string]$fileNameSuffix = "Export"
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\SnipeIT_LicenseReport_${fileNameSuffix}_${timestamp}.csv"
    
    $savePath = Read-Host "Enter path to save CSV file (default: $defaultPath)"
    if ([string]::IsNullOrWhiteSpace($savePath))
    {
        $savePath = $defaultPath
    }
    
    try
    {
        # Export to CSV
        $report | Export-Csv -Path $savePath -NoTypeInformation
        
        Write-Host "File successfully exported to: $savePath" -ForegroundColor Green
    } catch
    {
        Write-Host "Error exporting to CSV: $_" -ForegroundColor Red
    }
}


# Function to generate an Activity Log Report
function Generate-ActivityLogReport
{
    Clear-Host
    Write-Host "===== Activity Log Report =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Define report parameters
        Write-Host "This report shows recent activity in your Snipe-IT system." -ForegroundColor Yellow
        Write-Host "Choose report parameters:" -ForegroundColor Green
        
        # Ask for date range
        $defaultDays = 30
        $daysInput = Read-Host "Enter number of days of activity to include (default: $defaultDays)"
        $days = if ([string]::IsNullOrWhiteSpace($daysInput) -or -not [int]::TryParse($daysInput, [ref]$null))
        {
            $defaultDays
        } else
        {
            [int]$daysInput
        }
        
        $endDate = Get-Date
        $startDate = $endDate.AddDays(-$days)
        
        # Activity type filter
        Write-Host "`nChoose activity types to include:" -ForegroundColor Green
        Write-Host "1. All activities" -ForegroundColor Cyan
        Write-Host "2. Checkout/Checkin events only" -ForegroundColor Cyan
        Write-Host "3. Create/Update/Delete events only" -ForegroundColor Cyan
        Write-Host "4. Audit events only" -ForegroundColor Cyan
        
        $typeFilter = Read-Host "Enter option (1-4)"
        $activityTypeFilter = "all"
        
        switch ($typeFilter)
        {
            "2"
            { $activityTypeFilter = "checkout" 
            }
            "3"
            { $activityTypeFilter = "modify" 
            }
            "4"
            { $activityTypeFilter = "audit" 
            }
            default
            { $activityTypeFilter = "all" 
            }
        }
        
        # Step 2: Retrieve activity data
        Write-Host "`nRetrieving activity data..." -ForegroundColor Yellow
        
        $activities = Get-SnipeitActivity -all
        
        if (-not $activities -or $activities.Count -eq 0)
        {
            Write-Host "No activity records found in the system." -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
        
        Write-Host "Found $($activities.Count) activity records." -ForegroundColor Green
        
        # Step 3: Filter activity logs
        Write-Host "`nFiltering activity logs..." -ForegroundColor Yellow
        
        $filteredActivities = $activities | Where-Object {
            # Filter by date
            $activityDate = if ($_.created_at -is [string])
            {
                try
                { [DateTime]::Parse($_.created_at) 
                } catch
                { $null 
                }
            } elseif ($_.created_at -is [DateTime])
            {
                $_.created_at
            } else
            { $null 
            }
            
            $isInDateRange = $activityDate -and $activityDate -ge $startDate -and $activityDate -le $endDate
            
            # Filter by activity type
            $isMatchingType = $true
            if ($activityTypeFilter -ne "all")
            {
                if ($activityTypeFilter -eq "checkout")
                {
                    $isMatchingType = $_.action_type -match "checkout|checkin"
                } elseif ($activityTypeFilter -eq "modify")
                {
                    $isMatchingType = $_.action_type -match "create|update|delete"
                } elseif ($activityTypeFilter -eq "audit")
                {
                    $isMatchingType = $_.action_type -match "audit"
                }
            }
            
            $isInDateRange -and $isMatchingType
        }
        
        if (-not $filteredActivities -or $filteredActivities.Count -eq 0)
        {
            Write-Host "No activity records found matching your criteria." -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
        
        Write-Host "Found $($filteredActivities.Count) activities matching your criteria." -ForegroundColor Green
        
        # Step 4: Process activities into report format
        $activityReport = @()
        
        foreach ($activity in $filteredActivities)
        {
            # Parse date
            $activityDate = if ($activity.created_at -is [string])
            {
                try
                { [DateTime]::Parse($activity.created_at) 
                } catch
                { $endDate 
                }
            } elseif ($activity.created_at -is [DateTime])
            {
                $activity.created_at
            } else
            { $endDate 
            }
            
            # Format activity type for display
            $actionType = $activity.action_type
            $formattedAction = switch -Regex ($actionType)
            {
                "checkout"
                { "Checked Out" 
                }
                "checkin"
                { "Checked In" 
                }
                "create"
                { "Created" 
                }
                "update"
                { "Updated" 
                }
                "delete"
                { "Deleted" 
                }
                "audit"
                { "Audited" 
                }
                default
                { $actionType 
                }
            }
            
            # Format target type for display
            $targetType = $activity.item.type
            $formattedTargetType = switch -Regex ($targetType)
            {
                "asset"
                { "Asset" 
                }
                "user"
                { "User" 
                }
                "license"
                { "License" 
                }
                "accessory"
                { "Accessory" 
                }
                "component"
                { "Component" 
                }
                "consumable"
                { "Consumable" 
                }
                "location"
                { "Location" 
                }
                default
                { $targetType 
                }
            }
            
            # Create report object
            $activityReport += [PSCustomObject]@{
                'Date' = $activityDate.ToString("yyyy-MM-dd HH:mm:ss")
                'User' = if ($activity.admin)
                { $activity.admin.name 
                } else
                { "System" 
                }
                'Action' = $formattedAction
                'Item Type' = $formattedTargetType
                'Item Name' = if ($activity.item)
                { $activity.item.name 
                } else
                { "N/A" 
                }
                'Item Tag' = if ($activity.item.asset_tag)
                { $activity.item.asset_tag 
                } else
                { "" 
                }
                'Target' = if ($activity.target)
                { $activity.target.name 
                } else
                { "N/A" 
                }
                'Notes' = $activity.notes
                'Changed' = if ($activity.log_meta)
                {
                    # Try to extract field changes from log_meta
                    try
                    {
                        if ($activity.log_meta -is [string])
                        {
                            "Changed data"
                        } else
                        {
                            $changedFields = $activity.log_meta | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                            $changedFields -join ", "
                        }
                    } catch
                    {
                        "Changed data"
                    }
                } else
                { "" 
                }
            }
        }
        
        # Sort by date (newest first)
        $activityReport = $activityReport | Sort-Object { [DateTime]::Parse($_.Date) } -Descending
        
        # Step 5: Display summary and activity statistics
        Clear-Host
        Write-Host "===== Activity Log Report =====" -ForegroundColor Cyan
        Write-Host "Report Date: $($endDate.ToString('yyyy-MM-dd'))" -ForegroundColor Yellow
        Write-Host "Period: $($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd')) ($days days)" -ForegroundColor Yellow
        
        # Calculate statistics
        $totalActivities = $activityReport.Count
        $activityByUser = $activityReport | Group-Object User | Sort-Object Count -Descending
        $activityByType = $activityReport | Group-Object Action | Sort-Object Count -Descending
        $activityByItemType = $activityReport | Group-Object 'Item Type' | Sort-Object Count -Descending
        
        # Display statistics
        Write-Host "`nActivity Statistics:" -ForegroundColor Green
        Write-Host "Total Activities: $totalActivities" -ForegroundColor White
        
        Write-Host "`nTop 5 Users by Activity:" -ForegroundColor Green
        $activityByUser | Select-Object -First 5 | ForEach-Object {
            Write-Host "$($_.Name): $($_.Count) activities" -ForegroundColor White
        }
        
        Write-Host "`nActivity by Action Type:" -ForegroundColor Green
        $activityByType | ForEach-Object {
            Write-Host "$($_.Name): $($_.Count) activities" -ForegroundColor White
        }
        
        Write-Host "`nActivity by Item Type:" -ForegroundColor Green
        $activityByItemType | ForEach-Object {
            Write-Host "$($_.Name): $($_.Count) activities" -ForegroundColor White
        }
        
        # Step 6: Provide options for detailed views or export
        Write-Host "`nOptions:" -ForegroundColor Cyan
        Write-Host "1. View Recent Activities (Last 20)" -ForegroundColor Cyan
        Write-Host "2. View Activities by User" -ForegroundColor Cyan
        Write-Host "3. View Activities by Item Type" -ForegroundColor Cyan
        Write-Host "4. Export Full Report to CSV" -ForegroundColor Cyan
        Write-Host "5. Return to Menu" -ForegroundColor Cyan
        
        $option = Read-Host "Select an option (1-5)"
        
        switch ($option)
        {
            "1"
            {
                # View recent activities
                Write-Host "`nMost Recent Activities:" -ForegroundColor Green
                $activityReport | Select-Object -First 20 | 
                    Select-Object Date, User, Action, 'Item Type', 'Item Name', 'Target', 'Notes' | 
                    Format-Table -AutoSize
                
                $exportOption = Read-Host "Export this list to CSV? (Y/N)"
                if ($exportOption -eq "Y" -or $exportOption -eq "y")
                {
                    Export-ActivityReportToCSV -report ($activityReport | Select-Object -First 20) -fileNameSuffix "RecentActivities"
                }
                
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-ActivityLogReport
            }
            "2"
            {
                # View activities by user
                if ($activityByUser.Count -gt 0)
                {
                    Write-Host "`nSelect a user to view their activities:" -ForegroundColor Green
                    for ($i = 0; $i -lt [Math]::Min(10, $activityByUser.Count); $i++)
                    {
                        Write-Host "[$i] $($activityByUser[$i].Name) ($($activityByUser[$i].Count) activities)" -ForegroundColor Cyan
                    }
                    
                    $userIndex = Read-Host "Enter user number (0-$([Math]::Min(9, $activityByUser.Count-1)))"
                    if ([int]::TryParse($userIndex, [ref]$null) -and [int]$userIndex -ge 0 -and [int]$userIndex -lt [Math]::Min(10, $activityByUser.Count))
                    {
                        $selectedUser = $activityByUser[[int]$userIndex].Name
                        $userActivities = $activityReport | Where-Object { $_.User -eq $selectedUser }
                        
                        Write-Host "`nActivities for $selectedUser`:" -ForegroundColor Green
                        $userActivities | 
                            Select-Object Date, Action, 'Item Type', 'Item Name', 'Target', 'Notes' | 
                            Format-Table -AutoSize
                        
                        $exportOption = Read-Host "Export this list to CSV? (Y/N)"
                        if ($exportOption -eq "Y" -or $exportOption -eq "y")
                        {
                            Export-ActivityReportToCSV -report $userActivities -fileNameSuffix "User_$($selectedUser -replace '[^a-zA-Z0-9]', '_')"
                        }
                    } else
                    {
                        Write-Host "Invalid selection." -ForegroundColor Red
                    }
                } else
                {
                    Write-Host "No user activity data available." -ForegroundColor Yellow
                }
                
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-ActivityLogReport
            }
            "3"
            {
                # View activities by item type
                if ($activityByItemType.Count -gt 0)
                {
                    Write-Host "`nSelect an item type to view activities:" -ForegroundColor Green
                    for ($i = 0; $i -lt $activityByItemType.Count; $i++)
                    {
                        Write-Host "[$i] $($activityByItemType[$i].Name) ($($activityByItemType[$i].Count) activities)" -ForegroundColor Cyan
                    }
                    
                    $typeIndex = Read-Host "Enter item type number (0-$($activityByItemType.Count-1))"
                    if ([int]::TryParse($typeIndex, [ref]$null) -and [int]$typeIndex -ge 0 -and [int]$typeIndex -lt $activityByItemType.Count)
                    {
                        $selectedType = $activityByItemType[[int]$typeIndex].Name
                        $typeActivities = $activityReport | Where-Object { $_.'Item Type' -eq $selectedType }
                        
                        Write-Host "`nActivities for item type '$selectedType':" -ForegroundColor Green
                        $typeActivities | 
                            Select-Object Date, User, Action, 'Item Name', 'Target', 'Notes' | 
                            Format-Table -AutoSize
                        
                        $exportOption = Read-Host "Export this list to CSV? (Y/N)"
                        if ($exportOption -eq "Y" -or $exportOption -eq "y")
                        {
                            Export-ActivityReportToCSV -report $typeActivities -fileNameSuffix "ItemType_$($selectedType -replace '[^a-zA-Z0-9]', '_')"
                        }
                    } else
                    {
                        Write-Host "Invalid selection." -ForegroundColor Red
                    }
                } else
                {
                    Write-Host "No item type activity data available." -ForegroundColor Yellow
                }
                
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-ActivityLogReport
            }
            "4"
            {
                # Export full report to CSV
                Export-ActivityReportToCSV -report $activityReport -fileNameSuffix "FullReport"
                
                # Return to this function to see other options
                Start-Sleep -Seconds 2
                Generate-ActivityLogReport
            }
            "5"
            {
                return
            }
            default
            {
                Write-Host "Invalid option. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Generate-ActivityLogReport
            }
        }
    } catch
    {
        Write-Host "An error occurred while generating the report: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Helper function to export activity report to CSV
function Export-ActivityReportToCSV
{
    param (
        [Parameter(Mandatory = $true)]
        [array]$report,
        
        [Parameter(Mandatory = $false)]
        [string]$fileNameSuffix = "Export"
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\SnipeIT_ActivityReport_${fileNameSuffix}_${timestamp}.csv"
    
    $savePath = Read-Host "Enter path to save CSV file (default: $defaultPath)"
    if ([string]::IsNullOrWhiteSpace($savePath))
    {
        $savePath = $defaultPath
    }
    
    try
    {
        # Export to CSV
        $report | Export-Csv -Path $savePath -NoTypeInformation
        
        Write-Host "File successfully exported to: $savePath" -ForegroundColor Green 
    } catch
    {
        Write-Host "Error exporting to CSV: $_" -ForegroundColor Red
    }
}




function Update-SnipeitAssetInteractive
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$AssetTag
    )

    # Retrieve current asset information
    try
    {
        $asset = Get-SnipeitAsset -search $AssetTag | Where-Object { $_.asset_tag -eq $AssetTag }
        if (-not $asset)
        {
            Write-Host "Asset with tag $AssetTag not found." -ForegroundColor Red
            return
        }
    } catch
    {
        Write-Host "Error retrieving asset information: $_" -ForegroundColor Red
        return
    }

    Write-Host "Current Device Information:"
    Write-Host "Asset Tag: $($asset.asset_tag)"
    Write-Host "Name: $($asset.name)"
    Write-Host "Model: $($asset.model.name)"
    Write-Host "Serial: $($asset.serial)"
    Write-Host "Status: $($asset.status_label.name)"
    Write-Host "Location: $($asset.rtd_location.name)"
    Write-Host "Notes: $($asset.notes)"

    Write-Host ""  # Blank line for readability
    Write-Host "Enter new information (leave blank to keep current value):"

    $newName = Read-Host "Name [$($asset.name)]"
    $newSerial = Read-Host "Serial [$($asset.serial)]"
    $newModel = Read-Host "Model [$($asset.model.name)]"
    $newStatus = Read-Host "Status [$($asset.status_label.name)]"
    $newLocation = Read-Host "Location [$($asset.rtd_location.name)]"
    $newNotes = Read-Host "Notes [$($asset.notes)]"

    $params = @{ id = $asset.id }

    if ($newName)
    { $params.name = $newName 
    }
    if ($newSerial)
    { $params.serial = $newSerial 
    }

    # Handle Model ID lookup
    if ($newModel)
    {
        try
        {
            $model = Get-SnipeitModel -search $newModel | Where-Object { $_.name -eq $newModel }
            if ($model)
            { $params.model_id = $model.id 
            } else
            { Write-Host "Model '$newModel' not found." -ForegroundColor Yellow 
            }
        } catch
        {
            Write-Host "Error retrieving model ID: $_" -ForegroundColor Red
        }
    }

    # Handle Status ID lookup
    if ($newStatus)
    {
        try
        {
            $status = Get-SnipeitStatus -search $newStatus | Where-Object { $_.name -eq $newStatus }
            if ($status)
            { $params.status_id = $status.id 
            } else
            { Write-Host "Status '$newStatus' not found." -ForegroundColor Yellow 
            }
        } catch
        {
            Write-Host "Error retrieving status ID: $_" -ForegroundColor Red
        }
    }

    # Handle Location ID lookup
    if ($newLocation)
    {
        try
        {
            $location = Get-SnipeitLocation -search $newLocation | Where-Object { $_.name -eq $newLocation }
            if ($location)
            { $params.rtd_location_id = $location.id 
            } else
            { Write-Host "Location '$newLocation' not found." -ForegroundColor Yellow 
            }
        } catch
        {
            Write-Host "Error retrieving location ID: $_" -ForegroundColor Red
        }
    }

    if ($newNotes)
    { $params.notes = $newNotes 
    }

    try
    {
        Set-SnipeitAsset @params
        Write-Host "Device information updated successfully." -ForegroundColor Green
    } catch
    {
        Write-Host "Error updating device information: $_" -ForegroundColor Red
    }
}

function Get-DepreciationReport
{
    param (
        [int]$ThresholdYears = 3
    )

    try
    {
        # Get all assets
        $assets = Get-SnipeitAsset -all

        # Filter for assets older than the threshold
        $depreciatedAssets = $assets | Where-Object {
            $_.purchase_date -ne $null -and (New-TimeSpan -Start $_.purchase_date -End (Get-Date)).Days -gt ($ThresholdYears * 365)
        }

        # Display the report
        if ($depreciatedAssets.Count -gt 0)
        {
            $depreciatedAssets | Select-Object id, name, asset_tag, purchase_date, model.name, status_label.name | Format-Table -AutoSize
        } else
        {
            Write-Host "No depreciated assets found." -ForegroundColor Yellow
        }

    } catch
    {
        Write-Host "An error occurred while generating the depreciation report: $_" -ForegroundColor Red
    }
}

# Function to call the Custom Report Builder
function Call-CustomReportBuilder
{
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "CustomReportBuilder.ps1"
    
    if (Test-Path $scriptPath)
    {
        # Call the external script with the current API credentials
        & $scriptPath -APIKey $apiKey -SnipeURL $snipeURL
    } else
    {
        Write-Host "Custom Report Builder script not found at: $scriptPath" -ForegroundColor Red
        Write-Host "Please ensure the CustomReportBuilder.ps1 file is in the same directory as this script." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
    }
}

# Function to perform a Snipe-IT system health check
function Check-SystemHealth
{
    Clear-Host
    Write-Host "===== Snipe-IT System Health Check =====" -ForegroundColor Cyan
    
    try
    {
        # Step 1: Verify API Connection
        Write-Host "`nVerifying API connection..." -ForegroundColor Yellow
        try
        {
            $statusCheck = Get-SnipeitStatus -ErrorAction Stop
            if ($statusCheck)
            {
                Write-Host "✓ API connection successful" -ForegroundColor Green
            } else
            {
                Write-Host "✗ API connection returned no data" -ForegroundColor Red
            }
        } catch
        {
            Write-Host "✗ API connection failed: $_" -ForegroundColor Red
            Write-Host "Please check your API key and URL settings." -ForegroundColor Yellow
            Write-Host "Current URL: $snipeURL" -ForegroundColor Yellow
            
            $continueCheck = Read-Host "Would you like to continue with other checks? (Y/N)"
            if ($continueCheck -ne "Y" -and $continueCheck -ne "y")
            {
                return
            }
        }
        
        # Create results collection for final report
        $healthResults = @()
        $healthResults += [PSCustomObject]@{
            'Check' = "API Connection"
            'Status' = if ($statusCheck)
            { "Passed" 
            } else
            { "Failed" 
            }
            'Details' = if ($statusCheck)
            { "Connected to $snipeURL" 
            } else
            { "Could not connect to $snipeURL" 
            }
        }
        
        # Step 2: Check license expirations
        Write-Host "`nChecking license expirations..." -ForegroundColor Yellow
        $licenses = Get-SnipeitLicense -all
        
        if ($licenses)
        {
            $today = Get-Date
            $expiringLicenses = $licenses | Where-Object {
                if ($_.expiration_date)
                {
                    try
                    {
                        $expiryDate = if ($_.expiration_date -is [string])
                        {
                            [DateTime]::Parse($_.expiration_date)
                        } else
                        {
                            $_.expiration_date
                        }
                        
                        # Check if expiring within 90 days
                        ($expiryDate - $today).TotalDays -lt 90 -and $expiryDate -gt $today
                    } catch
                    {
                        $false
                    }
                } else
                {
                    $false
                }
            }
            
            $expiredLicenses = $licenses | Where-Object {
                if ($_.expiration_date)
                {
                    try
                    {
                        $expiryDate = if ($_.expiration_date -is [string])
                        {
                            [DateTime]::Parse($_.expiration_date)
                        } else
                        {
                            $_.expiration_date
                        }
                        
                        # Check if already expired
                        $expiryDate -lt $today
                    } catch
                    {
                        $false
                    }
                } else
                {
                    $false
                }
            }
            
            if ($expiringLicenses.Count -gt 0)
            {
                Write-Host "! $($expiringLicenses.Count) licenses will expire within 90 days:" -ForegroundColor Yellow
                foreach ($license in $expiringLicenses)
                {
                    $expiryDate = if ($license.expiration_date -is [string])
                    {
                        [DateTime]::Parse($license.expiration_date).ToString("yyyy-MM-dd")
                    } else
                    {
                        $license.expiration_date.ToString("yyyy-MM-dd")
                    }
                    
                    Write-Host "  - $($license.name): Expires $expiryDate" -ForegroundColor Yellow
                }
            } else
            {
                Write-Host "✓ No licenses expiring within 90 days" -ForegroundColor Green
            }
            
            if ($expiredLicenses.Count -gt 0)
            {
                Write-Host "✗ $($expiredLicenses.Count) licenses have already expired:" -ForegroundColor Red
                foreach ($license in $expiredLicenses)
                {
                    $expiryDate = if ($license.expiration_date -is [string])
                    {
                        [DateTime]::Parse($license.expiration_date).ToString("yyyy-MM-dd")
                    } else
                    {
                        $license.expiration_date.ToString("yyyy-MM-dd")
                    }
                    
                    Write-Host "  - $($license.name): Expired $expiryDate" -ForegroundColor Red
                }
            } else
            {
                Write-Host "✓ No expired licenses" -ForegroundColor Green
            }
            
            $healthResults += [PSCustomObject]@{
                'Check' = "License Expiration"
                'Status' = if ($expiredLicenses.Count -eq 0)
                { "Passed" 
                } else
                { "Warning" 
                }
                'Details' = "$($expiredLicenses.Count) expired, $($expiringLicenses.Count) expiring soon"
            }
        } else
        {
            Write-Host "✓ No licenses found to check" -ForegroundColor Green
            
            $healthResults += [PSCustomObject]@{
                'Check' = "License Expiration"
                'Status' = "N/A"
                'Details' = "No licenses found in system"
            }
        }
        
        # Step 3: Check for pending asset maintenance
        Write-Host "`nChecking asset maintenance status..." -ForegroundColor Yellow
        $maintenances = Get-SnipeitAssetMaintenance -all
        
        if ($maintenances)
        {
            $today = Get-Date
            $pendingMaintenance = $maintenances | Where-Object {
                if ($_.start_date -and -not $_.completion_date)
                {
                    $true # Maintenance started but not completed
                } else
                {
                    $false
                }
            }
            
            if ($pendingMaintenance.Count -gt 0)
            {
                Write-Host "! $($pendingMaintenance.Count) maintenance tasks are pending completion:" -ForegroundColor Yellow
                foreach ($maintenance in $pendingMaintenance)
                {
                    $asset = Get-SnipeitAsset -id $maintenance.asset_id
                    $startDate = if ($maintenance.start_date -is [string])
                    {
                        try
                        { [DateTime]::Parse($maintenance.start_date).ToString("yyyy-MM-dd") 
                        } catch
                        { $maintenance.start_date 
                        }
                    } else
                    {
                        $maintenance.start_date.ToString("yyyy-MM-dd")
                    }
                    
                    Write-Host "  - $($maintenance.title) for $($asset.name) (started $startDate)" -ForegroundColor Yellow
                }
            } else
            {
                Write-Host "✓ No pending maintenance tasks" -ForegroundColor Green
            }
            
            $healthResults += [PSCustomObject]@{
                'Check' = "Asset Maintenance"
                'Status' = if ($pendingMaintenance.Count -eq 0)
                { "Passed" 
                } else
                { "Warning" 
                }
                'Details' = "$($pendingMaintenance.Count) pending maintenance tasks"
            }
        } else
        {
            Write-Host "✓ No maintenance records found to check" -ForegroundColor Green
            
            $healthResults += [PSCustomObject]@{
                'Check' = "Asset Maintenance"
                'Status' = "N/A"
                'Details' = "No maintenance records found"
            }
        }
        
        # Step 4: Check for assets needing audit
        Write-Host "`nChecking for assets due for audit..." -ForegroundColor Yellow
        
        # Get audit interval setting - this would ideally come from Snipe-IT API
        # For now, assume 12 months as default
        $auditInterval = 12 # months
        $auditIntervalDays = $auditInterval * 30
        
        # Get assets
        $allAssets = Get-SnipeitAsset -all
        
        if ($allAssets)
        {
            $today = Get-Date
            $cutoffDate = $today.AddDays(-$auditIntervalDays)
            
            $assetsNeedingAudit = $allAssets | Where-Object {
                if ($_.last_audit)
                {
                    try
                    {
                        $auditDate = if ($_.last_audit -is [string])
                        {
                            [DateTime]::Parse($_.last_audit)
                        } else
                        {
                            $_.last_audit
                        }
                        
                        $auditDate -lt $cutoffDate
                    } catch
                    {
                        $true # If we can't parse the date, assume audit needed
                    }
                } else
                {
                    $true # No audit date means audit needed
                }
            }
            
            if ($assetsNeedingAudit.Count -gt 0)
            {
                $percentNeedingAudit = [Math]::Round(($assetsNeedingAudit.Count / $allAssets.Count) * 100, 2)
                Write-Host "! $($assetsNeedingAudit.Count) assets ($percentNeedingAudit%) need auditing (not audited in the last $auditInterval months)" -ForegroundColor Yellow
            } else
            {
                Write-Host "✓ All assets have been audited within the last $auditInterval months" -ForegroundColor Green
            }
            
            $healthResults += [PSCustomObject]@{
                'Check' = "Asset Audits"
                'Status' = if ($assetsNeedingAudit.Count -eq 0)
                { "Passed" 
                } else
                { "Warning" 
                }
                'Details' = "$($assetsNeedingAudit.Count) assets need auditing"
            }
        } else
        {
            Write-Host "! Could not retrieve assets to check audit status" -ForegroundColor Yellow
            
            $healthResults += [PSCustomObject]@{
                'Check' = "Asset Audits"
                'Status' = "Unknown"
                'Details' = "Could not retrieve assets"
            }
        }
        
        # Step 5: Check for data completeness
        Write-Host "`nChecking data completeness..." -ForegroundColor Yellow
        
        if ($allAssets)
        {
            $assetsWithoutSerial = $allAssets | Where-Object { [string]::IsNullOrWhiteSpace($_.serial) }
            $assetsWithoutLocation = $allAssets | Where-Object { -not $_.rtd_location -and -not $_.location }
            $assetsWithoutPurchaseDate = $allAssets | Where-Object { -not $_.purchase_date }
            
            $percentWithoutSerial = [Math]::Round(($assetsWithoutSerial.Count / $allAssets.Count) * 100, 2)
            $percentWithoutLocation = [Math]::Round(($assetsWithoutLocation.Count / $allAssets.Count) * 100, 2)
            $percentWithoutPurchaseDate = [Math]::Round(($assetsWithoutPurchaseDate.Count / $allAssets.Count) * 100, 2)
            
            if ($percentWithoutSerial -gt 5)
            {
                Write-Host "! $($assetsWithoutSerial.Count) assets ($percentWithoutSerial%) are missing serial numbers" -ForegroundColor Yellow
            } else
            {
                Write-Host "✓ $percentWithoutSerial% of assets are missing serial numbers (below 5% threshold)" -ForegroundColor Green
            }
            
            if ($percentWithoutLocation -gt 5)
            {
                Write-Host "! $($assetsWithoutLocation.Count) assets ($percentWithoutLocation%) are missing location information" -ForegroundColor Yellow
            } else
            {
                Write-Host "✓ $percentWithoutLocation% of assets are missing location information (below 5% threshold)" -ForegroundColor Green
            }
            
            if ($percentWithoutPurchaseDate -gt 10)
            {
                Write-Host "! $($assetsWithoutPurchaseDate.Count) assets ($percentWithoutPurchaseDate%) are missing purchase dates" -ForegroundColor Yellow
            } else
            {
                Write-Host "✓ $percentWithoutPurchaseDate% of assets are missing purchase dates (below 10% threshold)" -ForegroundColor Green
            }
            
            $overallDataQuality = "Good"
            if ($percentWithoutSerial -gt 20 -or $percentWithoutLocation -gt 20 -or $percentWithoutPurchaseDate -gt 30)
            {
                $overallDataQuality = "Poor"
            } elseif ($percentWithoutSerial -gt 10 -or $percentWithoutLocation -gt 10 -or $percentWithoutPurchaseDate -gt 20)
            {
                $overallDataQuality = "Fair"
            }
            
            $healthResults += [PSCustomObject]@{
                'Check' = "Data Completeness"
                'Status' = $overallDataQuality
                'Details' = "Serial: $percentWithoutSerial% missing, Location: $percentWithoutLocation% missing, Purchase Date: $percentWithoutPurchaseDate% missing"
            }
        } else
        {
            Write-Host "! Could not retrieve assets to check data completeness" -ForegroundColor Yellow
            
            $healthResults += [PSCustomObject]@{
                'Check' = "Data Completeness"
                'Status' = "Unknown"
                'Details' = "Could not retrieve assets"
            }
        }
        
        # Step 6: Check for checkout consistency
        Write-Host "`nChecking checkout consistency..." -ForegroundColor Yellow
        
        if ($allAssets)
        {
            # Find assets with deployed status but no assigned user
            $deployedStatusIds = (Get-SnipeitStatus | Where-Object { $_.name -match "Deployed|Assigned|Checked Out" }).id
            
            $inconsistentAssets = $allAssets | Where-Object { 
                $_.status_id -in $deployedStatusIds -and -not $_.assigned_to 
            }
            
            if ($inconsistentAssets.Count -gt 0)
            {
                Write-Host "✗ $($inconsistentAssets.Count) assets have deployed status but no assigned user" -ForegroundColor Red
                
                # Show a few examples
                $sampleCount = [Math]::Min(3, $inconsistentAssets.Count)
                Write-Host "Examples:" -ForegroundColor Yellow
                for ($i = 0; $i -lt $sampleCount; $i++)
                {
                    Write-Host "  - $($inconsistentAssets[$i].name) (Tag: $($inconsistentAssets[$i].asset_tag))" -ForegroundColor Yellow
                }
                
                if ($inconsistentAssets.Count -gt $sampleCount)
                {
                    Write-Host "  - ... and $($inconsistentAssets.Count - $sampleCount) more" -ForegroundColor Yellow
                }
            } else
            {
                Write-Host "✓ All deployed assets have assigned users" -ForegroundColor Green
            }
            
            $healthResults += [PSCustomObject]@{
                'Check' = "Checkout Consistency"
                'Status' = if ($inconsistentAssets.Count -eq 0)
                { "Passed" 
                } else
                { "Failed" 
                }
                'Details' = "$($inconsistentAssets.Count) assets with inconsistent checkout state"
            }
        } else
        {
            Write-Host "! Could not retrieve assets to check checkout consistency" -ForegroundColor Yellow
            
            $healthResults += [PSCustomObject]@{
                'Check' = "Checkout Consistency"
                'Status' = "Unknown"
                'Details' = "Could not retrieve assets"
            }
        }
        
        # Step 7: Overall system health summary
        Write-Host "`n===== System Health Summary =====" -ForegroundColor Cyan
        
        $passedChecks = ($healthResults | Where-Object { $_.Status -eq "Passed" }).Count
        $warningChecks = ($healthResults | Where-Object { $_.Status -eq "Warning" -or $_.Status -eq "Fair" }).Count
        $failedChecks = ($healthResults | Where-Object { $_.Status -eq "Failed" -or $_.Status -eq "Poor" }).Count
        $unknownChecks = ($healthResults | Where-Object { $_.Status -eq "Unknown" -or $_.Status -eq "N/A" }).Count
        
        $totalChecks = $healthResults.Count
        $healthScore = [Math]::Round((($passedChecks * 100) + ($warningChecks * 50)) / (($totalChecks - $unknownChecks) * 100) * 100)
        
        Write-Host "Health Score: $healthScore%" -ForegroundColor $(
            if ($healthScore -ge 90)
            { "Green" 
            } elseif ($healthScore -ge 70)
            { "Yellow" 
            } else
            { "Red" 
            }
        )
        
        Write-Host "`nDetailed Results:" -ForegroundColor Cyan
        $healthResults | Format-Table -Property Check, Status, Details -AutoSize
        
        # Option to export to CSV
        $exportOption = Read-Host "`nWould you like to export this health report to CSV? (Y/N)"
        if ($exportOption -eq "Y" -or $exportOption -eq "y")
        {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\SnipeIT_HealthCheck_${timestamp}.csv"
            
            $savePath = Read-Host "Enter path to save CSV file (default: $defaultPath)"
            if ([string]::IsNullOrWhiteSpace($savePath))
            {
                $savePath = $defaultPath
            }
            
            try
            {
                $healthResults | Export-Csv -Path $savePath -NoTypeInformation
                Write-Host "Health report exported to: $savePath" -ForegroundColor Green
            } catch
            {
                Write-Host "Error exporting health report: $_" -ForegroundColor Red
            }
        }
    } catch
    {
        Write-Host "An error occurred during the health check: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Main menu logic
function Main-MenuLogic
{
    $menuInput = Read-Host "Select an option"
    
    switch ($menuInput)
    {
        "1"
        { Asset-Management-Menu 
        }
        "2"
        { User-Management-Menu 
        }
        "3"
        { Maintenance-Menu 
        }
        "4"
        { Reporting-Menu 
        }
        "5"
        { Templates-Menu 
        }
        "Q"
        { exit 
        }
        "q"
        { exit 
        }
        default
        { 
            Write-Host "Invalid option, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Main-Menu
        }
    }
}

# Asset Management menu logic
function Asset-Management-Menu
{
    Show-AssetManagementMenu
    $menuInput = Read-Host "Select an option"
    
    switch ($menuInput)
    {
        "1"
        { 
            Write-Host "Add New Asset functionality will be implemented here" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Asset-Management-Menu
        }
        "2"
        { 
            Checkout-Laptop
            Asset-Management-Menu
        }
        "3"
        { 
            Checkin-Laptop
            Asset-Management-Menu
        }
        "4"
        { 
            Update-SnipeitAssetInteractive
            Start-Sleep -Seconds 2
            Asset-Management-Menu
        }
        "5"
        { 
            Write-Host "Bulk Import Assets functionality will be implemented here" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Asset-Management-Menu
        }
        "B"
        { Main-Menu 
        }
        "b"
        { Main-Menu 
        }
        "Q"
        { exit 
        }
        "q"
        { exit 
        }
        default
        { 
            Write-Host "Invalid option, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Asset-Management-Menu
        }
    }
}

# User Management menu logic
function User-Management-Menu
{
    Show-UserManagementMenu
    $menuInput = Read-Host "Select an option"
    
    switch ($menuInput)
    {
        "1"
        { 
            Add-SnipeITUser
            User-Management-Menu
        }
        "2"
        { 
            Update-SnipeITUserInfo
            User-Management-Menu
        }
        "3"
        { 
            View-UserAssets
            User-Management-Menu
        }
        "4"
        { 
            Disable-SnipeITUser
            User-Management-Menu
        }
        "5"
        { 
            Write-Host "Bulk Import Users functionality will be implemented here" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            User-Management-Menu
        }
        "B"
        { Main-Menu 
        }
        "b"
        { Main-Menu 
        }
        "Q"
        { exit 
        }
        "q"
        { exit 
        }
        default
        { 
            Write-Host "Invalid option, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            User-Management-Menu
        }
    }
}

# Maintenance menu logic
function Maintenance-Menu
{
    Show-MaintenanceMenu
    $menuInput = Read-Host "Select an option"
    
    switch ($menuInput)
    {
        "1"
        { 
            Write-Host "Schedule Maintenance functionality will be implemented here" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Maintenance-Menu
        }
        "2"
        { 
            Write-Host "View Overdue Maintenance functionality will be implemented here" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Maintenance-Menu
        }
        "3"
        { 
            Write-Host "View License Expirations functionality will be implemented here" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Maintenance-Menu
        }
        "4"
        { 
            Check-MissingAssets
            Maintenance-Menu
        }
        "5"
        { 
            Check-SystemHealth
            Maintenance-Menu
        }
        "B"
        { Main-Menu 
        }
        "b"
        { Main-Menu 
        }
        "Q"
        { exit 
        }
        "q"
        { exit 
        }
        default
        { 
            Write-Host "Invalid option, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Maintenance-Menu
        }
    }
}

# Reporting menu logic
function Reporting-Menu
{
    Show-ReportingMenu
    $menuInput = Read-Host "Select an option"
    
    switch ($menuInput)
    {
        "1"
        { 
            Generate-AssetAuditReport
            Start-Sleep -Seconds 2
            Reporting-Menu
        }
        "2"
        { 
            Generate-LicenseComplianceReport
            Start-Sleep -Seconds 2
            Reporting-Menu
        }
        "3"
        { 
            Get-DepreciationReport
            Start-Sleep -Seconds 2
            Reporting-Menu
        }
        "4"
        { 
            Generate-ActivityLogReport
            Start-Sleep -Seconds 2
            Reporting-Menu
        }
        "5"
        { 
            Call-CustomReportBuilder
            Start-Sleep -Seconds 2
            Reporting-Menu
        }
        "B"
        { Main-Menu 
        }
        "b"
        { Main-Menu 
        }
        "Q"
        { exit 
        }
        "q"
        { exit 
        }
        default
        { 
            Write-Host "Invalid option, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Reporting-Menu
        }
    }
}

# Templates menu logic
function Templates-Menu
{
    Show-TemplatesMenu
    $menuInput = Read-Host "Select an option"
    
    switch ($menuInput)
    {
        "1"
        { 
            Call-TemplatesScript
            Start-Sleep -Seconds 2
            Templates-Menu
        }
        "2"
        { 
            Call-TemplatesScript
            Start-Sleep -Seconds 2
            Templates-Menu
        }
        "3"
        { 
            Call-TemplatesScript
            Start-Sleep -Seconds 2
            Templates-Menu
        }
        "4"
        { 
            Call-TemplatesScript
            Start-Sleep -Seconds 2
            Templates-Menu
        }
        "5"
        { 
            Call-TemplatesScript
            Start-Sleep -Seconds 2
            Templates-Menu
        }
        "B"
        { Main-Menu 
        }
        "b"
        { Main-Menu 
        }
        "Q"
        { exit 
        }
        "q"
        { exit 
        }
        default
        { 
            Write-Host "Invalid option, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Templates-Menu
        }
    }
}

# Main function to start the menu
function Main-Menu
{
    Show-MainMenu
    Main-MenuLogic
}
# Start the menu Main-Menu
