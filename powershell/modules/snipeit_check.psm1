# SnipeIT Check-In/Check-Out Module
# Version: 1.1
# Description: PowerShell module for managing device check-in/check-out operations in Snipe-IT

# Helper function to handle user input with quit/back options
function Get-UserInputWithOptions {
    param(
        [string]$Prompt,
        [switch]$AllowBack = $true,
        [switch]$AllowQuit = $true
    )
    
    $options = @()
    if ($AllowBack) { $options += "'back' to return to main menu" }
    if ($AllowQuit) { $options += "'quit' to exit" }
    
    if ($options.Count -gt 0) {
        $optionText = " (or " + ($options -join ", ") + ")"
        $fullPrompt = $Prompt + $optionText + ": "
    } else {
        $fullPrompt = $Prompt + ": "
    }
    
    $input = Read-Host $fullPrompt
    
    if ($AllowQuit -and $input.ToLower() -eq "quit") {
        Write-Host "Exiting application..." -ForegroundColor Yellow
        exit
    }
    
    if ($AllowBack -and $input.ToLower() -eq "back") {
        return "BACK"
    }
    
    return $input
}

# Helper function to find location by name or ID
function Get-LocationId {
    param([string]$LocationInput)
    
    try {
        # First try to get all locations
        $locations = Get-SnipeitLocation
        
        # Check if input is a number (ID)
        if ($LocationInput -match '^\d+$') {
            $location = $locations | Where-Object { $_.id -eq [int]$LocationInput }
            if ($location) {
                return $location.id
            }
        }
        
        # Search by name (case-insensitive)
        $location = $locations | Where-Object { $_.name -like "*$LocationInput*" }
        
        if ($location) {
            if ($location.Count -gt 1) {
                Write-Host "Multiple locations found:" -ForegroundColor Yellow
                $location | ForEach-Object { Write-Host "  ID: $($_.id) - Name: $($_.name)" }
                $selectedId = Get-UserInputWithOptions -Prompt "Please enter the specific location ID"
                if ($selectedId -eq "BACK") { return $null }
                return [int]$selectedId
            } else {
                return $location.id
            }
        }
        
        Write-Host "Location not found: $LocationInput" -ForegroundColor Red
        return $null
    }
    catch {
        Write-Host "Error searching for location: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}


# Helper function to find user by username, employee number, or ID
function Get-UserId {
    param([string]$UserInput)
    
    try {
        # First, try searching using the -search parameter (works for username, email, employee number)
        $searchResult = Get-SnipeitUser -search $UserInput -ErrorAction SilentlyContinue
        
        if ($searchResult) {
            # If we get results from search
            if ($searchResult -is [array] -and $searchResult.Count -gt 1) {
                Write-Host "Multiple users found:" -ForegroundColor Yellow
                $searchResult | ForEach-Object { 
                    Write-Host "  ID: $($_.id) - Employee#: $($_.employee_num) - Username: $($_.username) - Name: $($_.name)" 
                }
                $selectedId = Get-UserInputWithOptions -Prompt "Please enter the specific user ID"
                if ($selectedId -eq "BACK") { return $null }
                
                # Get the selected user from the search results
                $selectedUser = $searchResult | Where-Object { $_.id -eq [int]$selectedId }
                if ($selectedUser) {
                    return @{
                        id = $selectedUser.id
                        name = $selectedUser.name
                        username = $selectedUser.username
                    }
                }
            } else {
                # Single user found
                return @{
                    id = $searchResult.id
                    name = $searchResult.name
                    username = $searchResult.username
                }
            }
        }
        
        # If search didn't work and input looks like an ID, try getting user by ID directly
        if ($UserInput -match '^\d+$') {
            try {
                # Try to get all users and filter by ID (this is a fallback)
                $allUsers = Get-SnipeitUser -ErrorAction SilentlyContinue
                $userById = $allUsers | Where-Object { $_.id -eq [int]$UserInput }
                
                if ($userById) {
                    return @{
                        id = $userById.id
                        name = $userById.name
                        username = $userById.username
                    }
                }
            }
            catch {
                # If getting all users fails, continue to not found message
            }
        }
        
        # If we get here, no user was found
        Write-Host "User not found: $UserInput" -ForegroundColor Red
        Write-Host "Try searching by:" -ForegroundColor Yellow
        Write-Host "  - Username (e.g., ldecareaux@nomma.net)" -ForegroundColor Yellow
        Write-Host "  - Employee number (e.g., 210303)" -ForegroundColor Yellow
        Write-Host "  - User ID number" -ForegroundColor Yellow
        return $null
    }
    catch {
        Write-Host "Error searching for user: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Helper function to find asset by tag
function Get-AssetByTag {
    param([string]$AssetTag)
    
    try {
        $asset = Get-SnipeitAsset -asset_tag $AssetTag
        if ($asset) {
            return $asset
        } else {
            Write-Host "Asset not found with tag: $AssetTag" -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Error finding asset: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}
<#
# LAPTOP/CHARGER FUNCTIONS
function CheckoutLaptopCharger {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "     LAPTOP & CHARGER CHECK-OUT        " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    
    # Get user
    $userInput = Get-UserInputWithOptions -Prompt "Enter username, employee number, or user ID"
    if ($userInput -eq "BACK") { return }
    
    $user = Get-UserId -UserInput $userInput
    if (-not $user) {
        Write-Host "Invalid user. Operation cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    Write-Host "Selected user: $($user.name) (ID: $($user.id))" -ForegroundColor Green
    
    # Get laptop asset tag
    $laptopTag = Get-UserInputWithOptions -Prompt "Enter laptop asset tag"
    if ($laptopTag -eq "BACK") { return }
    
    $laptopAsset = Get-AssetByTag -AssetTag $laptopTag
    if (-not $laptopAsset) {
        Write-Host "Operation cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Get charger asset tag
    $chargerTag = Get-UserInputWithOptions -Prompt "Enter charger asset tag"
    if ($chargerTag -eq "BACK") { return }
    
    $chargerAsset = Get-AssetByTag -AssetTag $chargerTag
    if (-not $chargerAsset) {
        Write-Host "Operation cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    try {
        # Checkout laptop to user
        Set-SnipeitAssetOwner -id $laptopAsset.id -assigned_id $user.id -checkout_to_type "user"
        Write-Host "Successfully checked out laptop $laptopTag to $($user.name)" -ForegroundColor Green
        
        # Checkout charger to user
        Set-SnipeitAssetOwner -id $chargerAsset.id -assigned_id $user.id -checkout_to_type "user"
        Write-Host "Successfully checked out charger $chargerTag to $($user.name)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error checking out laptop/charger: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Start-Sleep -Seconds 3
}
#>


# REPLACE the existing Get-UserInputWithOptions function with this one:
function Get-UserInputWithOptions {
    param(
        [string]$Prompt,
        [switch]$AllowBack = $true,
        [switch]$AllowQuit = $true
    )
    
    $options = @()
    if ($AllowBack) { $options += "'back' to return to main menu" }
    if ($AllowQuit) { $options += "'quit' to exit" }
    
    if ($options.Count -gt 0) {
        $optionText = " (or " + ($options -join ", ") + ")"
        $fullPrompt = $Prompt + $optionText + ": "
    } else {
        $fullPrompt = $Prompt + ": "
    }
    
    $input = Read-Host $fullPrompt
    
    if ($AllowQuit -and ($input.ToLower() -eq "quit" -or $input.ToLower() -eq "exit")) {
        Write-Host "Exiting application..." -ForegroundColor Yellow
        exit
    }
    
    if ($AllowBack -and $input.ToLower() -eq "back") {
        return "BACK"
    }
    
    return $input
}
# Helper function to find asset by tag
function Get-AssetByTag {
    param([string]$AssetTag)
    
    try {
        $asset = Get-SnipeitAsset -asset_tag $AssetTag
        if ($asset) {
            return $asset
        } else {
            Write-Host "Asset not found with tag: $AssetTag" -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Error finding asset: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# ADD YOUR NEW FUNCTION HERE
# Helper function to create a charger asset
function New-ChargerAsset {
    param(
        [string]$AssetTag,
        [int]$ChargerType
    )
    
    try {
        # Define charger model and type based on selection
        if ($ChargerType -eq 1) {
            $modelName = "HP 45W Charger"
            $modelId = 33  # REPLACE WITH YOUR ACTUAL HP CHARGER MODEL ID
        } else {
            $modelName = "Dell 65W Charger"
            $modelId = 32  # REPLACE WITH YOUR ACTUAL DELL CHARGER MODEL ID
        }
        
        # Create new asset
        $params = @{
            asset_tag = $AssetTag
            model_id = $modelId
            status_id = 2  # Ready to Deploy status (adjust as needed)
            name = "$modelName - $AssetTag"
        }
        
        $newAsset = New-SnipeitAsset @params
        
        if ($newAsset) {
            Write-Host "Successfully created new $modelName with asset tag $AssetTag" -ForegroundColor Green
            return $newAsset
        } else {
            Write-Host "Failed to create charger asset" -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Error creating charger asset: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}



# REPLACE the existing CheckoutLaptopCharger function with these three functions:
function CheckoutLaptopCharger {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "     LAPTOP & CHARGER CHECK-OUT        " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    
    # Ask if this is normal or summer checkout
    $mode = ""
    while ($mode -ne "normal" -and $mode -ne "summer") {
        $mode = Get-UserInputWithOptions -Prompt "Enter checkout mode (normal or summer)"
        if ($mode -eq "BACK") { return }
        $mode = $mode.ToLower()
        
        if ($mode -ne "normal" -and $mode -ne "summer") {
            Write-Host "Invalid mode. Please enter 'normal' or 'summer'." -ForegroundColor Yellow
        }
    }
    
    if ($mode -eq "normal") {
        # Normal checkout process
        PerformNormalCheckout
    } else {
        # Summer checkout process
        PerformSummerCheckout
    }
}


function PerformNormalCheckout {
    # Get user
    $userInput = Get-UserInputWithOptions -Prompt "Enter username, employee number, or user ID"
    if ($userInput -eq "BACK") { return }
    
    $user = Get-UserId -UserInput $userInput
    if (-not $user) {
        Write-Host "Invalid user. Operation cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    Write-Host "Selected user: $($user.name) (ID: $($user.id))" -ForegroundColor Green
    
    # Ask what to check out
    $checkoutType = ""
    while ($checkoutType -ne "laptop" -and $checkoutType -ne "charger" -and $checkoutType -ne "both") {
        $checkoutType = Get-UserInputWithOptions -Prompt "What do you want to check out? (laptop, charger, or both)"
        if ($checkoutType -eq "BACK") { return }
        $checkoutType = $checkoutType.ToLower()
        
        if ($checkoutType -ne "laptop" -and $checkoutType -ne "charger" -and $checkoutType -ne "both") {
            Write-Host "Invalid option. Please enter 'laptop', 'charger', or 'both'." -ForegroundColor Yellow
        }
    }
    
    # Process laptop checkout if requested
    $laptopAsset = $null
    if ($checkoutType -eq "laptop" -or $checkoutType -eq "both") {
        $laptopTag = Get-UserInputWithOptions -Prompt "Enter laptop asset tag"
        if ($laptopTag -eq "BACK") { return }
        
        $laptopAsset = Get-AssetByTag -AssetTag $laptopTag
        if (-not $laptopAsset) {
            Write-Host "Laptop not found. Operation cancelled." -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
    }
    
    # Process charger checkout if requested
    $chargerAsset = $null
    if ($checkoutType -eq "charger" -or $checkoutType -eq "both") {
        $chargerTag = Get-UserInputWithOptions -Prompt "Enter charger asset tag"
        if ($chargerTag -eq "BACK") { return }
        
        $chargerAsset = Get-AssetByTag -AssetTag $chargerTag
        if (-not $chargerAsset) {
            Write-Host "Charger not found. Operation cancelled." -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
    }
    
    try {
        # Checkout laptop to user if requested
        if ($laptopAsset) {
            Set-SnipeitAssetOwner -id $laptopAsset.id -assigned_id $user.id -checkout_to_type "user"
            Write-Host "Successfully checked out laptop $($laptopAsset.asset_tag) to $($user.name)" -ForegroundColor Green
        }
        
        # Checkout charger to user if requested
        if ($chargerAsset) {
            Set-SnipeitAssetOwner -id $chargerAsset.id -assigned_id $user.id -checkout_to_type "user"
            Write-Host "Successfully checked out charger $($chargerAsset.asset_tag) to $($user.name)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Error checking out asset(s): $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Start-Sleep -Seconds 3
}

function PerformSummerCheckout {
    # Loop until user types quit/exit (handled by Get-UserInputWithOptions)
    while ($true) {
        Write-Host "`n--- NEW CHECKOUT ---" -ForegroundColor Cyan
        
        # Get user for this checkout
        $userInput = Get-UserInputWithOptions -Prompt "Enter username, employee number, or user ID"
        if ($userInput -eq "BACK") { return }
        
        $user = Get-UserId -UserInput $userInput
        if (-not $user) {
            Write-Host "Invalid user. Operation cancelled for this checkout." -ForegroundColor Red
            Start-Sleep -Seconds 2
            continue
        }
        
        Write-Host "Selected user: $($user.name) (ID: $($user.id))" -ForegroundColor Green
        
        # Get laptop asset tag
        $laptopTag = Get-UserInputWithOptions -Prompt "Enter laptop asset tag"
        if ($laptopTag -eq "BACK") { return }
        
        $laptopAsset = Get-AssetByTag -AssetTag $laptopTag
        if (-not $laptopAsset) {
            Write-Host "Operation cancelled for this checkout." -ForegroundColor Red
            Start-Sleep -Seconds 2
            continue
        }
        
        # Get charger asset tag
        $chargerTag = Get-UserInputWithOptions -Prompt "Enter charger asset tag"
        if ($chargerTag -eq "BACK") { return }
        
        $chargerAsset = Get-AssetByTag -AssetTag $chargerTag
        
        # If charger doesn't exist, create it
        if (-not $chargerAsset) {
            Write-Host "Charger with tag $chargerTag not found. Creating new charger asset..." -ForegroundColor Yellow
            
            $chargerType = 0
            while ($chargerType -ne 1 -and $chargerType -ne 2) {
                $chargerTypeInput = Get-UserInputWithOptions -Prompt "Select charger type (1 for 45w HP, 2 for Dell 65w)"
                if ($chargerTypeInput -eq "BACK") { return }
                
                try {
                    $chargerType = [int]$chargerTypeInput
                    if ($chargerType -ne 1 -and $chargerType -ne 2) {
                        Write-Host "Invalid selection. Please enter 1 or 2." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Invalid input. Please enter a number (1 or 2)." -ForegroundColor Yellow
                }
            }
            
            # Create the charger asset
            $chargerAsset = New-ChargerAsset -AssetTag $chargerTag -ChargerType $chargerType
            
            if (-not $chargerAsset) {
                Write-Host "Failed to create charger. Operation cancelled for this checkout." -ForegroundColor Red
                Start-Sleep -Seconds 2
                continue
            }
        }
        
        try {
            # Checkout laptop to user
            Set-SnipeitAssetOwner -id $laptopAsset.id -assigned_id $user.id -checkout_to_type "user"
            Write-Host "Successfully checked out laptop $laptopTag to $($user.name)" -ForegroundColor Green
            
            # Checkout charger to user
            Set-SnipeitAssetOwner -id $chargerAsset.id -assigned_id $user.id -checkout_to_type "user" 
            Write-Host "Successfully checked out charger $chargerTag to $($user.name)" -ForegroundColor Green
            
            Write-Host "`nCheckout complete. Starting next checkout..." -ForegroundColor Cyan
            Write-Host "Type 'quit' or 'exit' at any prompt to finish." -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error checking out laptop/charger: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        Start-Sleep -Seconds 1
    }
}



function CheckinLaptopCharger {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "     LAPTOP & CHARGER CHECK-IN         " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    
    # Get laptop asset tag
    $laptopTag = Get-UserInputWithOptions -Prompt "Enter laptop asset tag"
    if ($laptopTag -eq "BACK") { return }
    
    $laptopAsset = Get-AssetByTag -AssetTag $laptopTag
    if (-not $laptopAsset) {
        Write-Host "Operation cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Get charger asset tag
    $chargerTag = Get-UserInputWithOptions -Prompt "Enter charger asset tag"
    if ($chargerTag -eq "BACK") { return }
    
    $chargerAsset = Get-AssetByTag -AssetTag $chargerTag
    if (-not $chargerAsset) {
        Write-Host "Operation cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    try {
        # Check in laptop
        Reset-SnipeitAssetOwner -id $laptopAsset.id -status_id "2"
        Write-Host "Successfully checked in laptop $laptopTag" -ForegroundColor Green
        
        # Check in charger
        Reset-SnipeitAssetOwner -id $chargerAsset.id -status_id "2"
        Write-Host "Successfully checked in charger $chargerTag" -ForegroundColor Green
    }
    catch {
        Write-Host "Error checking in laptop/charger: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Start-Sleep -Seconds 3
}

# HOTSPOT FUNCTIONS
function CheckoutHotspot {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "         HOTSPOT CHECK-OUT              " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    
    # Get user
    $userInput = Get-UserInputWithOptions -Prompt "Enter username, employee number, or user ID"
    if ($userInput -eq "BACK") { return }
    
    $user = Get-UserId -UserInput $userInput
    if (-not $user) {
        Write-Host "Invalid user. Operation cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    Write-Host "Selected user: $($user.name) (ID: $($user.id))" -ForegroundColor Green
    
    # Get hotspot asset tag
    $hotspotTag = Get-UserInputWithOptions -Prompt "Enter hotspot asset tag"
    if ($hotspotTag -eq "BACK") { return }
    
    $hotspotAsset = Get-AssetByTag -AssetTag $hotspotTag
    if (-not $hotspotAsset) {
        Write-Host "Operation cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    try {
        # Checkout hotspot to user
        Set-SnipeitAssetOwner -id $hotspotAsset.id -assigned_id $user.id 
        Write-Host "Successfully checked out hotspot $hotspotTag to $($user.name)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error checking out hotspot: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Start-Sleep -Seconds 3
}

function CheckinHotspot {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "         HOTSPOT CHECK-IN               " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    
    # Get hotspot asset tag
    $hotspotTag = Get-UserInputWithOptions -Prompt "Enter hotspot asset tag"
    if ($hotspotTag -eq "BACK") { return }
    
    $hotspotAsset = Get-AssetByTag -AssetTag $hotspotTag
    if (-not $hotspotAsset) {
        Write-Host "Operation cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    try {
        # Check in hotspot
        Reset-SnipeitAssetOwner -id $hotspotAsset.id -status_id "2"
        Write-Host "Successfully checked in hotspot $hotspotTag" -ForegroundColor Green
    }
    catch {
        Write-Host "Error checking in hotspot: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Start-Sleep -Seconds 3
}

# Export only laptop/charger and hotspot functions
Export-ModuleMember -Function CheckoutLaptopCharger, CheckinLaptopCharger, CheckoutHotspot, CheckinHotspot
