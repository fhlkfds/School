<#
.SYNOPSIS
    Snipe-IT Management Tool
.DESCRIPTION
    PowerShell script for managing Snipe-IT inventory and assets
.NOTES
    Version:        1.0
    Author:         Your Name
    Creation Date:  May 16, 2025
#>

# Check if SnipeitPS module is installed and import it
function Initialize-SnipeIT {
    if (-not (Get-Module -ListAvailable -Name SnipeitPS)) {
        Write-Host "SnipeitPS module is not installed. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name SnipeitPS -Force -Scope CurrentUser
            Write-Host "SnipeitPS module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install SnipeitPS module. Please install it manually with 'Install-Module -Name SnipeitPS'." -ForegroundColor Red
            return $false
        }
    }
    
    Import-Module SnipeitPS
    
    # Set Snipe-IT API parameters - Replace with your actual values
    $apiUrl = "https://inv.nomma.lan"
    $apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIzIiwianRpIjoiM2MzMGM0NzBmNmRkMGJiODQ0NGI1MDc2NjkwZWY2MzA0MTUzZjQzODhiMDA1MTEzNmM4NzE4NDc2NjQwOTZkYTk2YTQ5MDY4NWRlMzhkY2UiLCJpYXQiOjE3NDc0MDUyODUuNjUzOTY2LCJuYmYiOjE3NDc0MDUyODUuNjUzOTczLCJleHAiOjIyMjA3OTA4ODUuNjMzMTQ5LCJzdWIiOiIxIiwic2NvcGVzIjpbXX0.Bx9nImH2fmqQeWp1KjUi1_hbSpgOF-YJ__BLOthuhct-9OyaJzTgq7wjZ8DvYtEFvVIAr-_wXI37Ihe1PqOPT5SPuYqHl1vES51OQEFHNVHcPSPjQ5gJFraKY4f8Yqs26V5jiEYKo-z7wGfHRpEKAg3MzC8GgfIZUCbh-Xg5OmvdCjtYLQrsFB1G4M2alkGQyBzotI2QV__76JlA1dQIUdAX_6ZNadjxEVG0-GF1CPOO4IrYPZN-YZ6zztCEO8lR0vxSGj-Dtu1WCPqJM4iuE1Jy5TUeyLTCMOtk2Nw_G-LD_w_W6hhEhsxMca8HPwvDnN7V8YHYx1V5uTE5nacHw_gTTpK70kLV-XECljW3rSwfoV0SepHTml3GEECZk4EyNr4vK5DSk5DZwfrjUjzOBGqphyqH1q6mNU296-H5L7OrKfwEIO2HUscuyS6842JDBVZoFH-L2WEYc_PuX2Nndbc0vj0MW8kZUjypVJOn0_biBs2-xEPzgN7mroYGMf5xyaeWgomwEfaA-GX2fYfp7ovWLUhe4KXkkW16kGGgqkKqA62lC8lDYrUbCJuATJGMgBDNGeiSroldB7XlCmmskOwq2AcCUNyKbKMQZWJS89BgYjFmyM__Djv18-3oa0JW6w1norstpOzL8VExCSvM3jE-p2C0r9VMzEYKbaUfrcQ"
    
    try {
        Set-SnipeitInfo -URL $apiUrl -APIKey $apiKey
        Write-Host "Successfully connected to Snipe-IT API." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to connect to Snipe-IT API. Please check your connection settings." -ForegroundColor Red
        return $false
    }
}

# Function to get Snipe-IT user ID from employee ID
function Get-SnipeitUserByEmployeeId {
    param (
        [string]$EmployeeId
    )
    
    try {
        # Search for user by employee ID number
        $user = Get-SnipeitUser -search $EmployeeId
        
        # Filter results to find exact match on employee_num
        $exactMatch = $user | Where-Object { $_.employee_num -eq $EmployeeId }
        
        if ($exactMatch) {
            return $exactMatch.id
        }
        else {
            Write-Host "User with Employee ID $EmployeeId not found in Snipe-IT." -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Error finding user: $_" -ForegroundColor Red
        return $null
    }
}

# Function to get Snipe-IT asset ID from asset tag
function Get-SnipeitAssetByTag {
    param (
        [string]$AssetTag
    )
    
    try {
        # Search for asset by asset tag
        $asset = Get-SnipeitAsset -asset_tag $AssetTag
        
        if ($asset) {
            return $asset.id
        }
        else {
            Write-Host "Asset with tag $AssetTag not found in Snipe-IT." -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Error finding asset: $_" -ForegroundColor Red
        return $null
    }
}


# Function to check out laptops and chargers
function CheckoutLaptopCharger {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "    LAPTOP & CHARGER CHECK-OUT         " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Type 'back' to return to main menu or 'quit' to exit" -ForegroundColor Yellow
    Write-Host ""
    
    while ($true) {
        # Get Employee ID from user
        $employeeId = Read-Host "Enter Employee ID Number"
        
        if ($employeeId -eq "quit") {
            exit
        }
        elseif ($employeeId -eq "back") {
            return
        }
        
        # Get internal Snipe-IT User ID from the employee ID
        try {
            # Search for user by employee ID number
            $user = Get-SnipeitUser -search $employeeId
            
            # Filter results to find exact match on employee_num field
            $exactMatch = $user | Where-Object { $_.employee_num -eq $employeeId }
            
            if ($exactMatch) {
                $internalUserId = $exactMatch.id
                Write-Host "Found user: $($exactMatch.first_name) $($exactMatch.last_name)" -ForegroundColor Green
            }
            else {
                Write-Host "User with Employee ID $employeeId not found in Snipe-IT." -ForegroundColor Red
                Write-Host "Please try again with a valid Employee ID." -ForegroundColor Yellow
                continue
            }
        }
        catch {
            Write-Host "Error finding user: $_" -ForegroundColor Red
            Write-Host "Please try again with a valid Employee ID." -ForegroundColor Yellow
            continue
        }
        
        # Get Laptop Asset Tag
        $laptopTag = Read-Host "Enter Laptop Asset Tag"
        
        if ($laptopTag -eq "quit") {
            exit
        }
        elseif ($laptopTag -eq "back") {
            return
        }
        
        # Get Snipe-IT Asset ID from Asset Tag
        try {
            $laptop = Get-SnipeitAsset -asset_tag $laptopTag
            
            if ($laptop) {
                $laptopId = $laptop.id
                Write-Host "Found laptop: $($laptop.model.name)" -ForegroundColor Green
            }
            else {
                Write-Host "Asset with tag $laptopTag not found in Snipe-IT." -ForegroundColor Red
                Write-Host "Please try again with a valid Laptop Asset Tag." -ForegroundColor Yellow
                continue
            }
        }
        catch {
            Write-Host "Error finding laptop asset: $_" -ForegroundColor Red
            Write-Host "Please try again with a valid Laptop Asset Tag." -ForegroundColor Yellow
            continue
        }
        
        # Get Charger Asset Tag
        $chargerTag = Read-Host "Enter Charger Asset Tag"
        
        if ($chargerTag -eq "quit") {
            exit
        }
        elseif ($chargerTag -eq "back") {
            return
        }
        
        # Get Snipe-IT Asset ID from Asset Tag
        try {
            $charger = Get-SnipeitAsset -asset_tag $chargerTag
            
            if ($charger) {
                $chargerId = $charger.id
                Write-Host "Found charger: $($charger.model.name)" -ForegroundColor Green
            }
            else {
                Write-Host "Asset with tag $chargerTag not found in Snipe-IT." -ForegroundColor Red
                Write-Host "Please try again with a valid Charger Asset Tag." -ForegroundColor Yellow
                continue
            }
        }
        catch {
            Write-Host "Error finding charger asset: $_" -ForegroundColor Red
            Write-Host "Please try again with a valid Charger Asset Tag." -ForegroundColor Yellow
            continue
        }
        
        # Add checkout notes
        $checkoutNotes = Read-Host "Enter any checkout notes (optional)"
        
        # Check out the laptop
        try {
            # Call Set-SnipeitAssetOwner with properly named parameters
            # Based on the error, it seems the cmdlet doesn't accept 'user_id' directly in a splatted parameter
            $laptopResult = Set-SnipeitAssetOwner -id $laptopId -assigned_id $internalUserId -checkout_to_type user -note $checkoutNotes -ErrorAction Stop
            Write-Host "Laptop (Asset Tag: $laptopTag) checked out to Employee ID: $employeeId successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error checking out laptop: $_" -ForegroundColor Red
        }
        
        # Check out the charger
        try {
            # Call Set-SnipeitAssetOwner with properly named parameters
            $chargerResult = Set-SnipeitAssetOwner -id $chargerId -assigned_id $internalUserId -checkout_to_type user -note "Checked out with laptop $laptopTag. $checkoutNotes" -ErrorAction Stop
            Write-Host "Charger (Asset Tag: $chargerTag) checked out to Employee ID: $employeeId successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error checking out charger: $_" -ForegroundColor Red
        }
        
        Write-Host "`nAssets checked out successfully!" -ForegroundColor Green
        Write-Host "User: $($exactMatch.first_name) $($exactMatch.last_name)" -ForegroundColor Green
        Write-Host "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
        Write-Host "`nEnter information for next checkout or type 'back' to return to menu" -ForegroundColor Yellow
        Write-Host "------------------------------------------------" -ForegroundColor Cyan
    }
}


# Function to check in laptops and chargers
function CheckinLaptopCharger {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "    LAPTOP & CHARGER CHECK-IN          " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Type 'back' to return to main menu or 'quit' to exit" -ForegroundColor Yellow
    Write-Host ""
    
    while ($true) {
        # Get Laptop Asset Tag
        $laptopTag = Read-Host "Enter Laptop Asset Tag"
        
        if ($laptopTag -eq "quit") {
            exit
        }
        elseif ($laptopTag -eq "back") {
            return
        }
        
        # Get Snipe-IT Asset ID from Asset Tag
        $laptopId = Get-SnipeitAssetByTag -AssetTag $laptopTag
        
        if (-not $laptopId) {
            Write-Host "Please try again with a valid Laptop Asset Tag." -ForegroundColor Yellow
            continue
        }
        
        # Get Charger Asset Tag
        $chargerTag = Read-Host "Enter Charger Asset Tag"
        
        if ($chargerTag -eq "quit") {
            exit
        }
        elseif ($chargerTag -eq "back") {
            return
        }
        
        # Get Snipe-IT Asset ID from Asset Tag
        $chargerId = Get-SnipeitAssetByTag -AssetTag $chargerTag
        
        if (-not $chargerId) {
            Write-Host "Please try again with a valid Charger Asset Tag." -ForegroundColor Yellow
            continue
        }
        
        # Check in the laptop
        try {
            $laptopResult = Set-SnipeitAssetOwner -id $laptopId -user_id $null -ErrorAction Stop
            Write-Host "Laptop (Asset Tag: $laptopTag) checked in successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error checking in laptop: $_" -ForegroundColor Red
        }
        
        # Check in the charger
        try {
            $chargerResult = Set-SnipeitAssetOwner -id $chargerId -user_id $null -ErrorAction Stop
            Write-Host "Charger (Asset Tag: $chargerTag) checked in successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error checking in charger: $_" -ForegroundColor Red
        }
        
        Write-Host "`nEnter information for next check-in or type 'back' to return to menu" -ForegroundColor Yellow
        Write-Host "------------------------------------------------" -ForegroundColor Cyan
    }
}

# Function to check out hotspots
function CheckoutHotspot {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "        HOTSPOT CHECK-OUT              " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Type 'back' to return to main menu or 'quit' to exit" -ForegroundColor Yellow
    Write-Host ""
    
    while ($true) {
        # Get Employee ID from user
        $employeeId = Read-Host "Enter Employee ID Number"
        
        if ($employeeId -eq "quit") {
            exit
        }
        elseif ($employeeId -eq "back") {
            return
        }
        
        # Get internal Snipe-IT User ID from the employee ID
        try {
            # Search for user by employee ID number
            $user = Get-SnipeitUser -search $employeeId
            
            # Filter results to find exact match on employee_num field
            $exactMatch = $user | Where-Object { $_.employee_num -eq $employeeId }
            
            if ($exactMatch) {
                $internalUserId = $exactMatch.id
                Write-Host "Found user: $($exactMatch.first_name) $($exactMatch.last_name)" -ForegroundColor Green
            }
            else {
                Write-Host "User with Employee ID $employeeId not found in Snipe-IT." -ForegroundColor Red
                Write-Host "Please try again with a valid Employee ID." -ForegroundColor Yellow
                continue
            }
        }
        catch {
            Write-Host "Error finding user: $_" -ForegroundColor Red
            Write-Host "Please try again with a valid Employee ID." -ForegroundColor Yellow
            continue
        }
        
        # Get Hotspot Asset Tag
        $hotspotTag = Read-Host "Enter Hotspot Asset Tag"
        
        if ($hotspotTag -eq "quit") {
            exit
        }
        elseif ($hotspotTag -eq "back") {
            return
        }
        
        # Get Snipe-IT Asset ID from Asset Tag
        try {
            $hotspot = Get-SnipeitAsset -asset_tag $hotspotTag
            
            if ($hotspot) {
                $hotspotId = $hotspot.id
                Write-Host "Found hotspot: $($hotspot.model.name)" -ForegroundColor Green
            }
            else {
                Write-Host "Asset with tag $hotspotTag not found in Snipe-IT." -ForegroundColor Red
                Write-Host "Please try again with a valid Hotspot Asset Tag." -ForegroundColor Yellow
                continue
            }
        }
        catch {
            Write-Host "Error finding hotspot asset: $_" -ForegroundColor Red
            Write-Host "Please try again with a valid Hotspot Asset Tag." -ForegroundColor Yellow
            continue
        }
        
        # Add checkout notes
        $checkoutNotes = Read-Host "Enter any checkout notes (optional)"
        
        # Check out the hotspot
        try {
            $result = Set-SnipeitAssetOwner -id $hotspotId -assigned_id $internalUserId -checkout_to_type user -note $checkoutNotes -ErrorAction Stop
            Write-Host "Hotspot (Asset Tag: $hotspotTag) checked out to Employee ID: $employeeId successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error checking out hotspot: $_" -ForegroundColor Red
        }
        
        Write-Host "`nHotspot checked out successfully!" -ForegroundColor Green
        Write-Host "User: $($exactMatch.first_name) $($exactMatch.last_name)" -ForegroundColor Green
        Write-Host "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
        Write-Host "`nEnter information for next checkout or type 'back' to return to menu" -ForegroundColor Yellow
        Write-Host "------------------------------------------------" -ForegroundColor Cyan
    }
}
# Function to check in hotspots
function CheckinHotspot {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "        HOTSPOT CHECK-IN               " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Type 'back' to return to main menu or 'quit' to exit" -ForegroundColor Yellow
    Write-Host ""
    
    while ($true) {
        # Get Hotspot Asset Tag
        $hotspotTag = Read-Host "Enter Hotspot Asset Tag"
        
        if ($hotspotTag -eq "quit") {
            exit
        }
        elseif ($hotspotTag -eq "back") {
            return
        }
        
        # Get Snipe-IT Asset ID from Asset Tag
        $hotspotId = Get-SnipeitAssetByTag -AssetTag $hotspotTag
        
        if (-not $hotspotId) {
            Write-Host "Please try again with a valid Hotspot Asset Tag." -ForegroundColor Yellow
            continue
        }
        
        # Check in the hotspot
        try {
            $result = Set-SnipeitAssetOwner -id $hotspotId -user_id $null -ErrorAction Stop
            Write-Host "Hotspot (Asset Tag: $hotspotTag) checked in successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error checking in hotspot: $_" -ForegroundColor Red
        }
        
        Write-Host "`nEnter information for next check-in or type 'back' to return to menu" -ForegroundColor Yellow
        Write-Host "------------------------------------------------" -ForegroundColor Cyan
    }
}

# Function to get location ID from name
function Get-SnipeitLocationByName {
    param (
        [string]$LocationName
    )
    
    try {
        # Search for location by name
        $location = Get-SnipeitLocation -search $LocationName
        
        if ($location) {
            # If multiple results, find the closest match
            if ($location -is [array]) {
                $exactMatch = $location | Where-Object { $_.name -eq $LocationName } | Select-Object -First 1
                if ($exactMatch) {
                    return $exactMatch.id
                }
                
                # If no exact match, return the first close match
                return $location[0].id
            }
            
            return $location.id
        }
        else {
            Write-Host "Location '$LocationName' not found in Snipe-IT." -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "Error finding location: $_" -ForegroundColor Red
        return $null
    }
}

# Function to check out printers
function CheckoutPrinter {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "        PRINTER CHECK-OUT              " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Type 'back' to return to main menu or 'quit' to exit" -ForegroundColor Yellow
    Write-Host ""
    
    while ($true) {
        # Get Printer Asset Tag
        $printerTag = Read-Host "Enter Printer Asset Tag"
        
        if ($printerTag -eq "quit") {
            exit
        }
        elseif ($printerTag -eq "back") {
            return
        }
        
        # Get Snipe-IT Asset ID from Asset Tag
        try {
            $printer = Get-SnipeitAsset -asset_tag $printerTag
            
            if ($printer) {
                $printerId = $printer.id
                Write-Host "Found printer: $($printer.model.name)" -ForegroundColor Green
            }
            else {
                Write-Host "Asset with tag $printerTag not found in Snipe-IT." -ForegroundColor Red
                Write-Host "Please try again with a valid Printer Asset Tag." -ForegroundColor Yellow
                continue
            }
        }
        catch {
            Write-Host "Error finding printer asset: $_" -ForegroundColor Red
            Write-Host "Please try again with a valid Printer Asset Tag." -ForegroundColor Yellow
            continue
        }
        
        # Get Room Number/User ID
        $locationInfo = Read-Host "Enter Room Number or Employee ID"
        
        if ($locationInfo -eq "quit") {
            exit
        }
        elseif ($locationInfo -eq "back") {
            return
        }
        
        # Determine if location info is a room (location) or employee ID
        $isNumeric = $locationInfo -match '^\d+$'
        
        # Add checkout notes
        $checkoutNotes = Read-Host "Enter any checkout notes (optional)"
        
        # Check out the printer
        try {
            if ($isNumeric) {
                # Assume it's an employee ID if numeric
                $user = Get-SnipeitUser -search $locationInfo
                
                # Filter results to find exact match on employee_num field
                $exactMatch = $user | Where-Object { $_.employee_num -eq $locationInfo }
                
                if ($exactMatch) {
                    $internalUserId = $exactMatch.id
                    Write-Host "Found user: $($exactMatch.first_name) $($exactMatch.last_name)" -ForegroundColor Green
                    
                    $result = Set-SnipeitAssetOwner -id $printerId -assigned_id $internalUserId -checkout_to_type user -note $checkoutNotes -ErrorAction Stop
                    Write-Host "Printer (Asset Tag: $printerTag) checked out to Employee ID: $locationInfo successfully." -ForegroundColor Green
                    Write-Host "User: $($exactMatch.first_name) $($exactMatch.last_name)" -ForegroundColor Green
                }
                else {
                    Write-Host "User with Employee ID $locationInfo not found in Snipe-IT." -ForegroundColor Red
                    Write-Host "Please try again with a valid Employee ID." -ForegroundColor Yellow
                    continue
                }
            }
            else {
                # Assume it's a location name/room number
                $location = Get-SnipeitLocation -search $locationInfo
                
                # If multiple results, find the closest match
                if ($location -is [array]) {
                    $exactLocationMatch = $location | Where-Object { $_.name -eq $locationInfo } | Select-Object -First 1
                    if ($exactLocationMatch) {
                        $locationId = $exactLocationMatch.id
                        $locationName = $exactLocationMatch.name
                    }
                    else {
                        # If no exact match, use the first close match
                        $locationId = $location[0].id
                        $locationName = $location[0].name
                    }
                }
                else {
                    $locationId = $location.id
                    $locationName = $location.name
                }
                
                if ($locationId) {
                    Write-Host "Found location: $locationName" -ForegroundColor Green
                    $result = Set-SnipeitAssetOwner -id $printerId -assigned_id $locationId -checkout_to_type location -note $checkoutNotes -ErrorAction Stop
                    Write-Host "Printer (Asset Tag: $printerTag) checked out to Location: $locationName successfully." -ForegroundColor Green
                    Write-Host "Location: $locationName" -ForegroundColor Green
                }
                else {
                    Write-Host "Location '$locationInfo' not found in Snipe-IT." -ForegroundColor Red
                    Write-Host "Please try again with a valid Location name." -ForegroundColor Yellow
                    continue
                }
            }
        }
        catch {
            Write-Host "Error checking out printer: $_" -ForegroundColor Red
        }
        
        Write-Host "`nPrinter checked out successfully!" -ForegroundColor Green
        Write-Host "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
        Write-Host "`nEnter information for next checkout or type 'back' to return to menu" -ForegroundColor Yellow
        Write-Host "------------------------------------------------" -ForegroundColor Cyan
    }
}
# Function to check in printers
function CheckinPrinter {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "        PRINTER CHECK-IN               " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Type 'back' to return to main menu or 'quit' to exit" -ForegroundColor Yellow
    Write-Host ""
    
    while ($true) {
        # Get Printer Asset Tag
        $printerTag = Read-Host "Enter Printer Asset Tag"
        
        if ($printerTag -eq "quit") {
            exit
        }
        elseif ($printerTag -eq "back") {
            return
        }
        
        # Get Snipe-IT Asset ID from Asset Tag
        $printerId = Get-SnipeitAssetByTag -AssetTag $printerTag
        
        if (-not $printerId) {
            Write-Host "Please try again with a valid Printer Asset Tag." -ForegroundColor Yellow
            continue
        }
        
        # Check in the printer
        try {
            $result = Set-SnipeitAssetOwner -id $printerId -user_id $null -location_id $null -ErrorAction Stop
            Write-Host "Printer (Asset Tag: $printerTag) checked in successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error checking in printer: $_" -ForegroundColor Red
        }
        
        Write-Host "`nEnter information for next check-in or type 'back' to return to menu" -ForegroundColor Yellow
        Write-Host "------------------------------------------------" -ForegroundColor Cyan
    }
}

function Show-DeviceMenu {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "       DEVICE CHECK-IN/OUT             " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "1. Check-In/Out Laptops and Chargers"
    Write-Host "2. Check-In/Out Hotspots"
    Write-Host "3. Check-In/Out Printers"
    Write-Host "4. Return to Main Menu"
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-4]: " -ForegroundColor Yellow -NoNewline
}

function Show-LaptopChargerOptions {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "   LAPTOP & CHARGER CHECK-IN/OUT       " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "1. Check-Out"
    Write-Host "2. Check-In"
    Write-Host "3. Return to Previous Menu"
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-3]: " -ForegroundColor Yellow -NoNewline
}

function Show-HotspotOptions {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "       HOTSPOT CHECK-IN/OUT            " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "1. Check-Out"
    Write-Host "2. Check-In"
    Write-Host "3. Return to Previous Menu"
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-3]: " -ForegroundColor Yellow -NoNewline
}

function Show-PrinterOptions {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "       PRINTER CHECK-IN/OUT            " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "1. Check-Out"
    Write-Host "2. Check-In"
    Write-Host "3. Return to Previous Menu"
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-3]: " -ForegroundColor Yellow -NoNewline
}

function Process-LaptopChargerOptions {
    $continue = $true
    
    while ($continue) {
        Show-LaptopChargerOptions
        $selection = Read-Host
        
        switch ($selection) {
            "1" { CheckoutLaptopCharger }
            "2" { CheckinLaptopCharger }
            "3" { $continue = $false }
            default { 
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep -Seconds 2 
            }
        }
    }
}

function Process-HotspotOptions {
    $continue = $true
    
    while ($continue) {
        Show-HotspotOptions
        $selection = Read-Host
        
        switch ($selection) {
            "1" { CheckoutHotspot }
            "2" { CheckinHotspot }
            "3" { $continue = $false }
            default { 
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep -Seconds 2 
            }
        }
    }
}

function Process-PrinterOptions {
    $continue = $true
    
    while ($continue) {
        Show-PrinterOptions
        $selection = Read-Host
        
        switch ($selection) {
            "1" { CheckoutPrinter }
            "2" { CheckinPrinter }
            "3" { $continue = $false }
            default { 
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep -Seconds 2 
            }
        }
    }
}

function Process-DeviceMenu {
    $deviceContinue = $true
    
    while ($deviceContinue) {
        Show-DeviceMenu
        $deviceSelection = Read-Host
        
        switch ($deviceSelection) {
            "1" { Process-LaptopChargerOptions }
            "2" { Process-HotspotOptions }
            "3" { Process-PrinterOptions }
            "4" { $deviceContinue = $false }
            default { 
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep -Seconds 2 
            }
        }
    }
}

function Show-Menu {
    Clear-Host
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "       SNIPE-IT ADMIN TOOLKIT          " -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "1. Device Check-In/Out"
    Write-Host "2. User Management"
    Write-Host "3. Temporary and Broken Device Management"
    Write-Host "4. Reporting"
    Write-Host "5. CSV Templates"
    Write-Host "6. Exit"
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-6]: " -ForegroundColor Yellow -NoNewline
}

function Start-SnipeITTool {
    # Initialize Snipe-IT connection
    $initialized = Initialize-SnipeIT
    
    if (-not $initialized) {
        Write-Host "Failed to initialize Snipe-IT connection. Press any key to exit..." -ForegroundColor Red
        [void][System.Console]::ReadKey($true)
        return
    }
    
    $continue = $true
    
    while ($continue) {
        Show-Menu
        $mainSelection = Read-Host
        
        switch ($mainSelection) {
            "1" { Process-DeviceMenu }
            "2" { Write-Host "User Management - Feature not implemented yet" -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
            "3" { Write-Host "Temp & Broken Device Management - Feature not implemented yet" -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
            "4" { Write-Host "Reporting - Feature not implemented yet" -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
            "5" { Write-Host "CSV Templates - Feature not implemented yet" -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
            "6" { $continue = $false }
            default { Write-Host "Invalid selection" -ForegroundColor Red; Start-Sleep -Seconds 2 }
        }
    }
    
    Write-Host "Thank you for using the Snipe-IT Admin Toolkit. Goodbye!" -ForegroundColor Cyan
}

# Start the tool
Start-SnipeITTool
