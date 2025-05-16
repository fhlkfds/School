# Snipe-IT Automation Script
# Author: Claude
# Date: May 16, 2025
# Description: This script provides a menu-based interface for common Snipe-IT asset management tasks

#region Configuration
# Replace with your actual Snipe-IT API URL and token
$snipeITURL = "https://your-snipeit-instance.com/api/v1"
$apiToken = "YourAPITokenHere" 
#endregion

#region API Functions
# Function to make API calls to Snipe-IT
function Invoke-SnipeITAPI {
    param (
        [string]$Endpoint,
        [string]$Method = "GET",
        [object]$Body = $null
    )
    
    $headers = @{
        "Authorization" = "Bearer $apiToken"
        "Accept" = "application/json"
        "Content-Type" = "application/json"
    }
    
    $uri = "$snipeITURL/$Endpoint"
    
    try {
        if ($Method -eq "GET") {
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method $Method
        } else {
            $jsonBody = $Body | ConvertTo-Json -Depth 5
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method $Method -Body $jsonBody
        }
        return $response
    } catch {
        Write-Host "Error calling Snipe-IT API: $_" -ForegroundColor Red
        return $null
    }
}

# Function to get assets from Snipe-IT
function Get-SnipeITAssets {
    param (
        [string]$Search = "",
        [int]$Limit = 50
    )
    
    $endpoint = "hardware?limit=$Limit"
    if ($Search) {
        $endpoint += "&search=$Search"
    }
    
    $response = Invoke-SnipeITAPI -Endpoint $endpoint
    return $response.rows
}

# Function to get users from Snipe-IT
function Get-SnipeITUsers {
    param (
        [string]$Search = "",
        [int]$Limit = 50
    )
    
    $endpoint = "users?limit=$Limit"
    if ($Search) {
        $endpoint += "&search=$Search"
    }
    
    $response = Invoke-SnipeITAPI -Endpoint $endpoint
    return $response.rows
}

# Function to get locations from Snipe-IT
function Get-SnipeITLocations {
    $response = Invoke-SnipeITAPI -Endpoint "locations"
    return $response.rows
}

# Function to get status labels from Snipe-IT
function Get-SnipeITStatusLabels {
    $response = Invoke-SnipeITAPI -Endpoint "statuslabels"
    return $response.rows
}

# Function to get asset models from Snipe-IT
function Get-SnipeITModels {
    $response = Invoke-SnipeITAPI -Endpoint "models"
    return $response.rows
}

# Function to get groups from Snipe-IT
function Get-SnipeITGroups {
    $response = Invoke-SnipeITAPI -Endpoint "groups"
    return $response.rows
}
#endregion

#region Asset Management Functions
# Function to assign a laptop to a user
function Assign-Laptop {
    Clear-Host
    Write-Host "=== Assign Laptop to User ===" -ForegroundColor Cyan
    
    # Search for the asset
    $assetSearch = Read-Host "Enter laptop asset tag or serial number to search"
    $assets = Get-SnipeITAssets -Search $assetSearch
    
    if ($assets.Count -eq 0) {
        Write-Host "No assets found matching that search." -ForegroundColor Yellow
        return
    }
    
    # Display assets and let user choose
    Write-Host "Assets found:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $assets.Count; $i++) {
        Write-Host "[$i] Asset Tag: $($assets[$i].asset_tag) | Name: $($assets[$i].name) | S/N: $($assets[$i].serial)"
    }
    
    $assetIndex = Read-Host "Enter the number of the asset to assign"
    if (-not ($assetIndex -match '^\d+$') -or [int]$assetIndex -ge $assets.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        return
    }
    
    $selectedAsset = $assets[[int]$assetIndex]
    
    # Search for the user
    $userSearch = Read-Host "Enter name or email of user to assign laptop to"
    $users = Get-SnipeITUsers -Search $userSearch
    
    if ($users.Count -eq 0) {
        Write-Host "No users found matching that search." -ForegroundColor Yellow
        return
    }
    
    # Display users and let user choose
    Write-Host "Users found:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $users.Count; $i++) {
        Write-Host "[$i] $($users[$i].name) | $($users[$i].email)"
    }
    
    $userIndex = Read-Host "Enter the number of the user to assign laptop to"
    if (-not ($userIndex -match '^\d+$') -or [int]$userIndex -ge $users.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        return
    }
    
    $selectedUser = $users[[int]$userIndex]
    
    # Get checkout reason
    $checkoutReason = Read-Host "Enter checkout reason (optional)"
    
    # Prepare checkout request
    $body = @{
        assigned_user = $selectedUser.id
        checkout_to_type = "user"
        note = $checkoutReason
    }
    
    # Execute checkout
    $response = Invoke-SnipeITAPI -Endpoint "hardware/$($selectedAsset.id)/checkout" -Method "POST" -Body $body
    
    if ($response) {
        Write-Host "Successfully assigned $($selectedAsset.name) (Asset Tag: $($selectedAsset.asset_tag)) to $($selectedUser.name)" -ForegroundColor Green
    }
    
    Read-Host "Press Enter to continue"
}

# Function to mass checkin/checkout laptops
function Invoke-MassLaptopAction {
    Clear-Host
    Write-Host "=== Mass Laptop Actions ===" -ForegroundColor Cyan
    Write-Host "1. Mass Checkout Laptops"
    Write-Host "2. Mass Checkin Laptops"
    Write-Host "0. Return to Main Menu"
    
    $choice = Read-Host "Select an option"
    
    switch ($choice) {
        "1" { Invoke-MassCheckout }
        "2" { Invoke-MassCheckin }
        "0" { return }
        default { 
            Write-Host "Invalid option selected." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            Invoke-MassLaptopAction
        }
    }
}

# Function for mass checkout of laptops
function Invoke-MassCheckout {
    Clear-Host
    Write-Host "=== Mass Checkout Laptops ===" -ForegroundColor Cyan
    
    # Get list of laptops - various ways to select multiple laptops
    Write-Host "How would you like to select laptops?"
    Write-Host "1. By Location"
    Write-Host "2. By Model"
    Write-Host "3. By Status"
    Write-Host "4. Enter Asset Tags Manually"
    Write-Host "0. Return to Previous Menu"
    
    $selectionMethod = Read-Host "Select an option"
    $selectedAssets = @()
    
    switch ($selectionMethod) {
        "1" {
            # Select by location
            $locations = Get-SnipeITLocations
            Write-Host "Available Locations:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $locations.Count; $i++) {
                Write-Host "[$i] $($locations[$i].name)"
            }
            
            $locationIndex = Read-Host "Enter location number"
            if (-not ($locationIndex -match '^\d+$') -or [int]$locationIndex -ge $locations.Count) {
                Write-Host "Invalid selection." -ForegroundColor Red
                return
            }
            
            $selectedLocation = $locations[[int]$locationIndex]
            $assets = Get-SnipeITAssets -Limit 100 # Increase limit to get more assets
            $selectedAssets = $assets | Where-Object { $_.rtd_location.id -eq $selectedLocation.id -and $_.status_label.status_meta -eq 'deployable' }
        }
        "2" {
            # Select by model
            $models = Get-SnipeITModels
            Write-Host "Available Models:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $models.Count; $i++) {
                Write-Host "[$i] $($models[$i].name)"
            }
            
            $modelIndex = Read-Host "Enter model number"
            if (-not ($modelIndex -match '^\d+$') -or [int]$modelIndex -ge $models.Count) {
                Write-Host "Invalid selection." -ForegroundColor Red
                return
            }
            
            $selectedModel = $models[[int]$modelIndex]
            $assets = Get-SnipeITAssets -Limit 100
            $selectedAssets = $assets | Where-Object { $_.model.id -eq $selectedModel.id -and $_.status_label.status_meta -eq 'deployable' }
        }
        "3" {
            # Select by status
            $statuses = Get-SnipeITStatusLabels
            Write-Host "Available Statuses:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $statuses.Count; $i++) {
                Write-Host "[$i] $($statuses[$i].name) ($($statuses[$i].status_meta))"
            }
            
            $statusIndex = Read-Host "Enter status number"
            if (-not ($statusIndex -match '^\d+$') -or [int]$statusIndex -ge $statuses.Count) {
                Write-Host "Invalid selection." -ForegroundColor Red
                return
            }
            
            $selectedStatus = $statuses[[int]$statusIndex]
            $assets = Get-SnipeITAssets -Limit 100
            $selectedAssets = $assets | Where-Object { $_.status_label.id -eq $selectedStatus.id }
        }
        "4" {
            # Enter asset tags manually
            $assetTags = Read-Host "Enter asset tags separated by commas"
            $tagArray = $assetTags -split ',' | ForEach-Object { $_.Trim() }
            
            foreach ($tag in $tagArray) {
                $asset = Get-SnipeITAssets -Search $tag
                if ($asset) {
                    $selectedAssets += $asset
                }
            }
        }
        "0" { 
            Invoke-MassLaptopAction
            return 
        }
        default { 
            Write-Host "Invalid option selected." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            Invoke-MassCheckout
            return
        }
    }
    
    if ($selectedAssets.Count -eq 0) {
        Write-Host "No deployable assets found with the specified criteria." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Display selected assets
    Write-Host "Selected Assets ($($selectedAssets.Count) total):" -ForegroundColor Cyan
    for ($i = 0; $i -lt [Math]::Min($selectedAssets.Count, 10); $i++) {
        Write-Host "- Asset Tag: $($selectedAssets[$i].asset_tag) | Name: $($selectedAssets[$i].name)"
    }
    
    if ($selectedAssets.Count -gt 10) {
        Write-Host "... and $($selectedAssets.Count - 10) more" -ForegroundColor Yellow
    }
    
    $confirm = Read-Host "Do you want to checkout these assets? (Y/N)"
    if ($confirm -ne "Y" -and $confirm -ne "y") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Search for the user to assign to
    $userSearch = Read-Host "Enter name or email of user to assign laptops to"
    $users = Get-SnipeITUsers -Search $userSearch
    
    if ($users.Count -eq 0) {
        Write-Host "No users found matching that search." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Display users and let user choose
    Write-Host "Users found:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $users.Count; $i++) {
        Write-Host "[$i] $($users[$i].name) | $($users[$i].email)"
    }
    
    $userIndex = Read-Host "Enter the number of the user to assign laptops to"
    if (-not ($userIndex -match '^\d+$') -or [int]$userIndex -ge $users.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $selectedUser = $users[[int]$userIndex]
    
    # Get checkout reason
    $checkoutReason = Read-Host "Enter checkout reason (optional)"
    
    # Checkout each asset
    $successCount = 0
    $failCount = 0
    
    foreach ($asset in $selectedAssets) {
        # Prepare checkout request
        $body = @{
            assigned_user = $selectedUser.id
            checkout_to_type = "user"
            note = $checkoutReason
        }
        
        # Execute checkout
        $response = Invoke-SnipeITAPI -Endpoint "hardware/$($asset.id)/checkout" -Method "POST" -Body $body
        
        if ($response) {
            $successCount++
            Write-Host "Successfully assigned $($asset.name) (Asset Tag: $($asset.asset_tag))" -ForegroundColor Green
        } else {
            $failCount++
            Write-Host "Failed to assign $($asset.name) (Asset Tag: $($asset.asset_tag))" -ForegroundColor Red
        }
    }
    
    Write-Host "Mass checkout complete. Success: $successCount, Failed: $failCount" -ForegroundColor Cyan
    Read-Host "Press Enter to continue"
}

# Function for mass checkin of laptops
function Invoke-MassCheckin {
    Clear-Host
    Write-Host "=== Mass Checkin Laptops ===" -ForegroundColor Cyan
    
    # Get list of checked out laptops
    $assets = Get-SnipeITAssets -Limit 100
    $checkedOutAssets = $assets | Where-Object { $_.assigned_to -ne $null }
    
    if ($checkedOutAssets.Count -eq 0) {
        Write-Host "No checked out assets found." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Display checked out assets
    Write-Host "Checked Out Assets ($($checkedOutAssets.Count) total):" -ForegroundColor Cyan
    for ($i = 0; $i -lt [Math]::Min($checkedOutAssets.Count, 20); $i++) {
        Write-Host "[$i] Asset Tag: $($checkedOutAssets[$i].asset_tag) | Name: $($checkedOutAssets[$i].name) | Assigned To: $($checkedOutAssets[$i].assigned_to.name)"
    }
    
    if ($checkedOutAssets.Count -gt 20) {
        Write-Host "... and $($checkedOutAssets.Count - 20) more" -ForegroundColor Yellow
    }
    
    Write-Host "`nHow would you like to select assets for checkin?"
    Write-Host "1. Select Individual Assets"
    Write-Host "2. Select All Assets From Specific User"
    Write-Host "3. Enter Asset Tags Manually"
    Write-Host "0. Return to Previous Menu"
    
    $selectionMethod = Read-Host "Select an option"
    $selectedAssets = @()
    
    switch ($selectionMethod) {
        "1" {
            # Select individual assets
            $assetIndices = Read-Host "Enter the numbers of the assets to check in (comma separated)"
            $indices = $assetIndices -split ',' | ForEach-Object { $_.Trim() }
            
            foreach ($index in $indices) {
                if ($index -match '^\d+$' -and [int]$index -lt $checkedOutAssets.Count) {
                    $selectedAssets += $checkedOutAssets[[int]$index]
                }
            }
        }
        "2" {
            # Select all assets from specific user
            $userSearch = Read-Host "Enter name or email of user to check in assets from"
            $users = Get-SnipeITUsers -Search $userSearch
            
            if ($users.Count -eq 0) {
                Write-Host "No users found matching that search." -ForegroundColor Yellow
                Read-Host "Press Enter to continue"
                return
            }
            
            # Display users and let user choose
            Write-Host "Users found:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $users.Count; $i++) {
                Write-Host "[$i] $($users[$i].name) | $($users[$i].email)"
            }
            
            $userIndex = Read-Host "Enter the number of the user"
            if (-not ($userIndex -match '^\d+$') -or [int]$userIndex -ge $users.Count) {
                Write-Host "Invalid selection." -ForegroundColor Red
                Read-Host "Press Enter to continue"
                return
            }
            
            $selectedUser = $users[[int]$userIndex]
            $selectedAssets = $checkedOutAssets | Where-Object { $_.assigned_to.id -eq $selectedUser.id }
        }
        "3" {
            # Enter asset tags manually
            $assetTags = Read-Host "Enter asset tags separated by commas"
            $tagArray = $assetTags -split ',' | ForEach-Object { $_.Trim() }
            
            foreach ($tag in $tagArray) {
                $asset = $checkedOutAssets | Where-Object { $_.asset_tag -eq $tag }
                if ($asset) {
                    $selectedAssets += $asset
                }
            }
        }
        "0" { 
            Invoke-MassLaptopAction
            return 
        }
        default { 
            Write-Host "Invalid option selected." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            Invoke-MassCheckin
            return
        }
    }
    
    if ($selectedAssets.Count -eq 0) {
        Write-Host "No assets selected for checkin." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Display selected assets
    Write-Host "Assets Selected for Checkin ($($selectedAssets.Count) total):" -ForegroundColor Cyan
    for ($i = 0; $i -lt [Math]::Min($selectedAssets.Count, 10); $i++) {
        Write-Host "- Asset Tag: $($selectedAssets[$i].asset_tag) | Name: $($selectedAssets[$i].name) | Assigned To: $($selectedAssets[$i].assigned_to.name)"
    }
    
    if ($selectedAssets.Count -gt 10) {
        Write-Host "... and $($selectedAssets.Count - 10) more" -ForegroundColor Yellow
    }
    
    $confirm = Read-Host "Do you want to check in these assets? (Y/N)"
    if ($confirm -ne "Y" -and $confirm -ne "y") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Get checkin reason
    $checkinReason = Read-Host "Enter checkin reason (optional)"
    
    # Checkin each asset
    $successCount = 0
    $failCount = 0
    
    foreach ($asset in $selectedAssets) {
        # Prepare checkin request
        $body = @{
            note = $checkinReason
        }
        
        # Execute checkin
        $response = Invoke-SnipeITAPI -Endpoint "hardware/$($asset.id)/checkin" -Method "POST" -Body $body
        
        if ($response) {
            $successCount++
            Write-Host "Successfully checked in $($asset.name) (Asset Tag: $($asset.asset_tag))" -ForegroundColor Green
        } else {
            $failCount++
            Write-Host "Failed to check in $($asset.name) (Asset Tag: $($asset.asset_tag))" -ForegroundColor Red
        }
    }
    
    Write-Host "Mass checkin complete. Success: $successCount, Failed: $failCount" -ForegroundColor Cyan
    Read-Host "Press Enter to continue"
}

# Function to handle broken laptop workflow
function Handle-BrokenLaptop {
    Clear-Host
    Write-Host "=== Broken Laptop Workflow ===" -ForegroundColor Cyan
    
    # Search for the asset
    $assetSearch = Read-Host "Enter laptop asset tag or serial number"
    $assets = Get-SnipeITAssets -Search $assetSearch
    
    if ($assets.Count -eq 0) {
        Write-Host "No assets found matching that search." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Display assets and let user choose
    Write-Host "Assets found:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $assets.Count; $i++) {
        Write-Host "[$i] Asset Tag: $($assets[$i].asset_tag) | Name: $($assets[$i].name) | S/N: $($assets[$i].serial)"
    }
    
    $assetIndex = Read-Host "Enter the number of the asset to mark as broken"
    if (-not ($assetIndex -match '^\d+$') -or [int]$assetIndex -ge $assets.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $selectedAsset = $assets[[int]$assetIndex]
    
    # Step 1: If checked out, check in the asset
    if ($selectedAsset.assigned_to) {
        Write-Host "Asset is currently assigned to $($selectedAsset.assigned_to.name). Checking in..." -ForegroundColor Yellow
        
        $checkinNote = Read-Host "Enter checkin reason"
        $body = @{
            note = $checkinNote
        }
        
        $checkinResponse = Invoke-SnipeITAPI -Endpoint "hardware/$($selectedAsset.id)/checkin" -Method "POST" -Body $body
        
        if ($checkinResponse) {
            Write-Host "Successfully checked in the asset." -ForegroundColor Green
        } else {
            Write-Host "Failed to check in the asset. Cannot proceed." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    }
    
    # Step 2: Change status to broken
    $statuses = Get-SnipeITStatusLabels
    $brokenStatus = $statuses | Where-Object { $_.name -like "*broken*" -or $_.name -like "*repair*" }
    
    if (-not $brokenStatus) {
        Write-Host "No 'broken' status found. Available statuses:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $statuses.Count; $i++) {
            Write-Host "[$i] $($statuses[$i].name)"
        }
        
        $statusIndex = Read-Host "Enter the number of the status to use"
        if (-not ($statusIndex -match '^\d+$') -or [int]$statusIndex -ge $statuses.Count) {
            Write-Host "Invalid selection." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        $brokenStatus = $statuses[[int]$statusIndex]
    } else {
        Write-Host "Found broken status: $($brokenStatus.name)" -ForegroundColor Green
    }
    
    # Update asset status
    $updateBody = @{
        status_id = $brokenStatus.id
        notes = Read-Host "Enter notes about the broken laptop (issue description)"
    }
    
    $updateResponse = Invoke-SnipeITAPI -Endpoint "hardware/$($selectedAsset.id)" -Method "PATCH" -Body $updateBody
    
    if ($updateResponse) {
        Write-Host "Successfully updated asset status to $($brokenStatus.name)." -ForegroundColor Green
    } else {
        Write-Host "Failed to update asset status." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    # Step 3: Change group to broken/repair group if available
    $groups = Get-SnipeITGroups
    $repairGroup = $groups | Where-Object { $_.name -like "*repair*" -or $_.name -like "*broken*" }
    
    if ($repairGroup) {
        Write-Host "Found repair group: $($repairGroup.name). Updating group..." -ForegroundColor Yellow
        
        $groupUpdateBody = @{
            group_id = $repairGroup.id
        }
        
        $groupUpdateResponse = Invoke-SnipeITAPI -Endpoint "hardware/$($selectedAsset.id)" -Method "PATCH" -Body $groupUpdateBody
        
        if ($groupUpdateResponse) {
            Write-Host "Successfully moved asset to $($repairGroup.name) group." -ForegroundColor Green
        } else {
            Write-Host "Failed to update asset group." -ForegroundColor Yellow
        }
    }
    
    # Step 4: Ask if a replacement laptop should be assigned
    $assignReplacement = Read-Host "Do you want to assign a replacement laptop? (Y/N)"
    
    if ($assignReplacement -eq "Y" -or $assignReplacement -eq "y") {
        # Get available replacement laptops
        $replacements = Get-SnipeITAssets -Limit 100 | Where-Object { 
            $_.model.id -eq $selectedAsset.model.id -and 
            $_.status_label.status_meta -eq 'deployable' -and
            $_.assigned_to -eq $null
        }
        
        if ($replacements.Count -eq 0) {
            Write-Host "No available replacement laptops of the same model found." -ForegroundColor Yellow
            $replacements = Get-SnipeITAssets -Limit 100 | Where-Object { 
                $_.status_label.status_meta -eq 'deployable' -and
                $_.assigned_to -eq $null
            }
            
            if ($replacements.Count -eq 0) {
                Write-Host "No available replacement laptops found." -ForegroundColor Red
                Read-Host "Press Enter to continue"
                return
            }
            
            Write-Host "Found other available laptops instead." -ForegroundColor Yellow
        }
        
        # Display replacement options
        Write-Host "Available Replacement Laptops:" -ForegroundColor Cyan
        for ($i = 0; $i -lt [Math]::Min($replacements.Count, 10); $i++) {
            Write-Host "[$i] Asset Tag: $($replacements[$i].asset_tag) | Model: $($replacements[$i].model.name) | S/N: $($replacements[$i].serial)"
        }
        
        if ($replacements.Count -gt 10) {
            Write-Host "... and $($replacements.Count - 10) more" -ForegroundColor Yellow
        }
        
        $replacementIndex = Read-Host "Enter the number of the replacement laptop to assign"
        if (-not ($replacementIndex -match '^\d+$') -or [int]$replacementIndex -ge $replacements.Count) {
            Write-Host "Invalid selection." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        $selectedReplacement = $replacements[[int]$replacementIndex]
        
        # Get user to assign replacement to
        $userSearch = Read-Host "Enter name or email of user to assign replacement laptop to"
        $users = Get-SnipeITUsers -Search $userSearch
        
        if ($users.Count -eq 0) {
            Write-Host "No users found matching that search." -ForegroundColor Yellow
            Read-Host "Press Enter to continue"
            return
        }
        
        # Display users and let user choose
        Write-Host "Users found:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $users.Count; $i++) {
            Write-Host "[$i] $($users[$i].name) | $($users[$i].email)"
        }
        
        $userIndex = Read-Host "Enter the number of the user to assign replacement laptop to"
        if (-not ($userIndex -match '^\d+$') -or [int]$userIndex -ge $users.Count) {
            Write-Host "Invalid selection." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        $selectedUser = $users[[int]$userIndex]
        
        # Assign replacement laptop
        $checkoutReason = "Replacement for broken laptop (Asset Tag: $($selectedAsset.asset_tag))"
        
        $body = @{
            assigned_user = $selectedUser.id
            checkout_to_type = "user"
            note = $checkoutReason
        }
        
        $response = Invoke-SnipeITAPI -Endpoint "hardware/$($selectedReplacement.id)/checkout" -Method "POST" -Body $body
        
        if ($response) {
            Write-Host "Successfully assigned replacement laptop $($selectedReplacement.name) (Asset Tag: $($selectedReplacement.asset_tag)) to $($selectedUser.name)" -ForegroundColor Green
            
            # Step 5: Change status of replacement to deployed if needed
            $deployedStatus = $statuses | Where-Object { $_.name -like "*deployed*" }
            
            if ($deployedStatus) {
                $deployUpdateBody = @{
                    status_id = $deployedStatus.id
                }
                
                $deployUpdateResponse = Invoke-SnipeITAPI -Endpoint "hardware/$($selectedReplacement.id)" -Method "PATCH" -Body $deployUpdateBody
                
                if ($deployUpdateResponse) {
                    Write-Host "Successfully updated replacement laptop status to $($deployedStatus.name)." -ForegroundColor Green
                }
            }
        } else {
            Write-Host "Failed to assign replacement laptop."
        }
    }
    
    Write-Host "Broken laptop workflow completed successfully." -ForegroundColor Green
    Read-Host "Press Enter to continue"
}


# Function to manage Active User Group
function Manage-ActiveUserGroup {
    Clear-Host
    Write-Host "=== Manage Active User Group ===" -ForegroundColor Cyan
    Write-Host "1. View Active Users"
    Write-Host "2. Add User to Active Group"
    Write-Host "3. Remove User from Active Group"
    Write-Host "0. Return to Main Menu"
    
    $choice = Read-Host "Select an option"
    
    switch ($choice) {
        "1" { View-ActiveUsers }
        "2" { Add-UserToActiveGroup }
        "3" { Remove-UserFromActiveGroup }
        "0" { return }
        default { 
            Write-Host "Invalid option selected." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            Manage-ActiveUserGroup
        }
    }
}

# Function to view active users
function View-ActiveUsers {
    Clear-Host
    Write-Host "=== View Active Users ===" -ForegroundColor Cyan
    
    # Get groups
    $groups = Get-SnipeITGroups
    $activeGroup = $groups | Where-Object { $_.name -like "*active*" }
    
    if (-not $activeGroup) {
        Write-Host "No 'Active' group found. Available groups:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $groups.Count; $i++) {
            Write-Host "[$i] $($groups[$i].name)"
        }
        
        $groupIndex = Read-Host "Enter the number of the group to view"
        if (-not ($groupIndex -match '^\d+) -or [int]$groupIndex -ge $groups.Count) {
            Write-Host "Invalid selection." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        $activeGroup = $groups[[int]$groupIndex]
    }
    
    # Get users in active group
    $users = Get-SnipeITUsers -Limit 200
    $activeUsers = $users | Where-Object { $_.groups -and $_.groups.contains($activeGroup.id) }
    
    if ($activeUsers.Count -eq 0) {
        Write-Host "No users found in the $($activeGroup.name) group." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Display active users
    Write-Host "Users in $($activeGroup.name) Group ($($activeUsers.Count) total):" -ForegroundColor Cyan
    foreach ($user in $activeUsers) {
        Write-Host "- $($user.name) | $($user.email) | $($user.employee_num)"
    }
    
    Read-Host "Press Enter to continue"
}

# Function to add user to active group
function Add-UserToActiveGroup {
    Clear-Host
    Write-Host "=== Add User to Active Group ===" -ForegroundColor Cyan
    
    # Get groups
    $groups = Get-SnipeITGroups
    $activeGroup = $groups | Where-Object { $_.name -like "*active*" }
    
    if (-not $activeGroup) {
        Write-Host "No 'Active' group found. Available groups:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $groups.Count; $i++) {
            Write-Host "[$i] $($groups[$i].name)"
        }
        
        $groupIndex = Read-Host "Enter the number of the group to use"
        if (-not ($groupIndex -match '^\d+) -or [int]$groupIndex -ge $groups.Count) {
            Write-Host "Invalid selection." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        $activeGroup = $groups[[int]$groupIndex]
    }
    
    # Search for the user
    $userSearch = Read-Host "Enter name or email of user to add to $($activeGroup.name) group"
    $users = Get-SnipeITUsers -Search $userSearch
    
    if ($users.Count -eq 0) {
        Write-Host "No users found matching that search." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Display users and let user choose
    Write-Host "Users found:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $users.Count; $i++) {
        Write-Host "[$i] $($users[$i].name) | $($users[$i].email)"
    }
    
    $userIndex = Read-Host "Enter the number of the user to add"
    if (-not ($userIndex -match '^\d+) -or [int]$userIndex -ge $users.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $selectedUser = $users[[int]$userIndex]
    
    # Check if user is already in the group
    if ($selectedUser.groups -and $selectedUser.groups.contains($activeGroup.id)) {
        Write-Host "$($selectedUser.name) is already in the $($activeGroup.name) group." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Add user to group
    $userGroups = $selectedUser.groups -or @()
    $userGroups += $activeGroup.id
    
    $updateBody = @{
        groups = $userGroups
    }
    
    $updateResponse = Invoke-SnipeITAPI -Endpoint "users/$($selectedUser.id)" -Method "PATCH" -Body $updateBody
    
    if ($updateResponse) {
        Write-Host "Successfully added $($selectedUser.name) to $($activeGroup.name) group." -ForegroundColor Green
    } else {
        Write-Host "Failed to add user to group." -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}

# Function to remove user from active group
function Remove-UserFromActiveGroup {
    Clear-Host
    Write-Host "=== Remove User from Active Group ===" -ForegroundColor Cyan
    
    # Get groups
    $groups = Get-SnipeITGroups
    $activeGroup = $groups | Where-Object { $_.name -like "*active*" }
    
    if (-not $activeGroup) {
        Write-Host "No 'Active' group found. Available groups:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $groups.Count; $i++) {
            Write-Host "[$i] $($groups[$i].name)"
        }
        
        $groupIndex = Read-Host "Enter the number of the group to use"
        if (-not ($groupIndex -match '^\d+) -or [int]$groupIndex -ge $groups.Count) {
            Write-Host "Invalid selection." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        $activeGroup = $groups[[int]$groupIndex]
    }
    
    # Get users in active group
    $users = Get-SnipeITUsers -Limit 200
    $activeUsers = $users | Where-Object { $_.groups -and $_.groups.contains($activeGroup.id) }
    
    if ($activeUsers.Count -eq 0) {
        Write-Host "No users found in the $($activeGroup.name) group." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Display active users and let user choose
    Write-Host "Users in $($activeGroup.name) Group:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $activeUsers.Count; $i++) {
        Write-Host "[$i] $($activeUsers[$i].name) | $($activeUsers[$i].email)"
    }
    
    $userIndex = Read-Host "Enter the number of the user to remove"
    if (-not ($userIndex -match '^\d+) -or [int]$userIndex -ge $activeUsers.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    $selectedUser = $activeUsers[[int]$userIndex]
    
    # Remove user from group
    $userGroups = $selectedUser.groups | Where-Object { $_ -ne $activeGroup.id }
    
    $updateBody = @{
        groups = $userGroups
    }
    
    $updateResponse = Invoke-SnipeITAPI -Endpoint "users/$($selectedUser.id)" -Method "PATCH" -Body $updateBody
    
    if ($updateResponse) {
        Write-Host "Successfully removed $($selectedUser.name) from $($activeGroup.name) group." -ForegroundColor Green
    } else {
        Write-Host "Failed to remove user from group." -ForegroundColor Red
    }
    
    Read-Host "Press Enter to continue"
}
#endregion

#region Reporting Functions
# Function for reporting menu
function Show-ReportingMenu {
    Clear-Host
    Write-Host "=== Reporting Menu ===" -ForegroundColor Cyan
    Write-Host "1. Broken/Temporary Assets Report"
    Write-Host "2. Issues Report"
    Write-Host "3. Activity Report"
    Write-Host "4. Asset Audit Report"
    Write-Host "5. Generate Report Templates"
    Write-Host "0. Return to Main Menu"
    
    $choice = Read-Host "Select an option"
    
    switch ($choice) {
        "1" { Generate-BrokenAssetsReport }
        "2" { Generate-IssuesReport }
        "3" { Generate-ActivityReport }
        "4" { Generate-AssetAuditReport }
        "5" { Generate-ReportTemplates }
        "0" { return }
        default { 
            Write-Host "Invalid option selected." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            Show-ReportingMenu
        }
    }
}

# Function to generate broken/temporary assets report
function Generate-BrokenAssetsReport {
    Clear-Host
    Write-Host "=== Broken/Temporary Assets Report ===" -ForegroundColor Cyan
    
    # Get all assets
    Write-Host "Retrieving assets from Snipe-IT..." -ForegroundColor Yellow
    $assets = Get-SnipeITAssets -Limit 500
    
    # Get statuses
    $statuses = Get-SnipeITStatusLabels
    $brokenStatus = $statuses | Where-Object { $_.name -like "*broken*" -or $_.name -like "*repair*" }
    $temporaryStatus = $statuses | Where-Object { $_.name -like "*temp*" -or $_.name -like "*loaner*" }
    
    if (-not $brokenStatus) {
        Write-Host "No 'broken' status found. Available statuses:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $statuses.Count; $i++) {
            Write-Host "[$i] $($statuses[$i].name)"
        }
        
        $statusIndex = Read-Host "Enter the number of the status for broken assets"
        if (-not ($statusIndex -match '^\d+) -or [int]$statusIndex -ge $statuses.Count) {
            Write-Host "Invalid selection." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
        
        $brokenStatus = $statuses[[int]$statusIndex]
    }
    
    # Filter assets
    $brokenAssets = $assets | Where-Object { $_.status_label.id -eq $brokenStatus.id }
    $temporaryAssets = @()
    
    if ($temporaryStatus) {
        $temporaryAssets = $assets | Where-Object { $_.status_label.id -eq $temporaryStatus.id }
    }
    
    # Create report
    $reportPath = "$env:USERPROFILE\Desktop\BrokenAssets_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    $reportData = @()
    
    foreach ($asset in $brokenAssets) {
        $reportData += [PSCustomObject]@{
            AssetTag = $asset.asset_tag
            Name = $asset.name
            Model = $asset.model.name
            SerialNumber = $asset.serial
            Status = $asset.status_label.name
            Location = $asset.rtd_location.name
            LastUpdated = $asset.updated_at
            Notes = $asset.notes
            AssignedTo = if ($asset.assigned_to) { $asset.assigned_to.name } else { "Unassigned" }
        }
    }
    
    foreach ($asset in $temporaryAssets) {
        $reportData += [PSCustomObject]@{
            AssetTag = $asset.asset_tag
            Name = $asset.name
            Model = $asset.model.name
            SerialNumber = $asset.serial
            Status = $asset.status_label.name
            Location = $asset.rtd_location.name
            LastUpdated = $asset.updated_at
            Notes = $asset.notes
            AssignedTo = if ($asset.assigned_to) { $asset.assigned_to.name } else { "Unassigned" }
        }
    }
    
    # Export to CSV
    $reportData | Export-Csv -Path $reportPath -NoTypeInformation
    
    # Display summary
    Write-Host "Report Summary:" -ForegroundColor Cyan
    Write-Host "- Total Broken Assets: $($brokenAssets.Count)" -ForegroundColor Yellow
    Write-Host "- Total Temporary/Loaner Assets: $($temporaryAssets.Count)" -ForegroundColor Yellow
    Write-Host "- Report saved to: $reportPath" -ForegroundColor Green
    
    Read-Host "Press Enter to continue"
}

# Function to generate issues report
function Generate-IssuesReport {
    Clear-Host
    Write-Host "=== Issues Report ===" -ForegroundColor Cyan
    
    # Get all assets
    Write-Host "Retrieving assets from Snipe-IT..." -ForegroundColor Yellow
    $assets = Get-SnipeITAssets -Limit 500
    
    # Filter assets with notes containing issue keywords
    $issueKeywords = @("issue", "problem", "broken", "not working", "repair", "damaged", "faulty")
    $assetsWithIssues = $assets | Where-Object { 
        $asset = $_
        $hasIssue = $false
        
        foreach ($keyword in $issueKeywords) {
            if ($asset.notes -and $asset.notes -match $keyword) {
                $hasIssue = $true
                break
            }
        }
        
        $hasIssue
    }
    
    # Create report
    $reportPath = "$env:USERPROFILE\Desktop\IssuesReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    $reportData = @()
    
    foreach ($asset in $assetsWithIssues) {
        $reportData += [PSCustomObject]@{
            AssetTag = $asset.asset_tag
            Name = $asset.name
            Model = $asset.model.name
            SerialNumber = $asset.serial
            Status = $asset.status_label.name
            Location = $asset.rtd_location.name
            LastUpdated = $asset.updated_at
            Issue = $asset.notes
            AssignedTo = if ($asset.assigned_to) { $asset.assigned_to.name } else { "Unassigned" }
        }
    }
    
    # Export to CSV
    $reportData | Export-Csv -Path $reportPath -NoTypeInformation
    
    # Display summary
    Write-Host "Report Summary:" -ForegroundColor Cyan
    Write-Host "- Total Assets with Issues: $($assetsWithIssues.Count)" -ForegroundColor Yellow
    Write-Host "- Report saved to: $reportPath" -ForegroundColor Green
    
    Read-Host "Press Enter to continue"
}

# Function to generate activity report
function Generate-ActivityReport {
    Clear-Host
    Write-Host "=== Activity Report ===" -ForegroundColor Cyan
    
    # Get date range
    $startDateStr = Read-Host "Enter start date (MM/DD/YYYY) or leave blank for last 30 days"
    
    if ([string]::IsNullOrEmpty($startDateStr)) {
        $startDate = (Get-Date).AddDays(-30)
    } else {
        $startDate = [DateTime]::ParseExact($startDateStr, "MM/dd/yyyy", $null)
    }
    
    $endDateStr = Read-Host "Enter end date (MM/DD/YYYY) or leave blank for today"
    
    if ([string]::IsNullOrEmpty($endDateStr)) {
        $endDate = Get-Date
    } else {
        $endDate = [DateTime]::ParseExact($endDateStr, "MM/dd/yyyy", $null)
    }
    
    # Format dates for API
    $startDateFormatted = $startDate.ToString("yyyy-MM-dd")
    $endDateFormatted = $endDate.ToString("yyyy-MM-dd")
    
    # Get activity logs
    Write-Host "Retrieving activity logs from Snipe-IT..." -ForegroundColor Yellow
    $endpoint = "reports/activity?limit=500&start_date=$startDateFormatted&end_date=$endDateFormatted"
    $response = Invoke-SnipeITAPI -Endpoint $endpoint
    
    if (-not $response -or -not $response.rows) {
        Write-Host "No activity logs found for the specified date range." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    $activities = $response.rows
    
    # Create report
    $reportPath = "$env:USERPROFILE\Desktop\ActivityReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    $reportData = @()
    
    foreach ($activity in $activities) {
        $reportData += [PSCustomObject]@{
            Date = $activity.created_at
            Action = $activity.action_type
            Item = $activity.item.name
            ItemType = $activity.item_type
            User = $activity.admin_name
            Target = $activity.target_name
            Notes = $activity.note
        }
    }
    
    # Export to CSV
    $reportData | Export-Csv -Path $reportPath -NoTypeInformation
    
    # Display summary
    Write-Host "Report Summary:" -ForegroundColor Cyan
    Write-Host "- Date Range: $startDate to $endDate" -ForegroundColor Yellow
    Write-Host "- Total Activities: $($activities.Count)" -ForegroundColor Yellow
    Write-Host "- Report saved to: $reportPath" -ForegroundColor Green
    
    Read-Host "Press Enter to continue"
}

# Function to generate asset audit report
function Generate-AssetAuditReport {
    Clear-Host
    Write-Host "=== Asset Audit Report ===" -ForegroundColor Cyan
    
    # Ask for audit type
    Write-Host "Select Audit Type:"
    Write-Host "1. All Assets"
    Write-Host "2. Assets by Location"
    Write-Host "3. Assets by Model"
    Write-Host "4. Assets by Status"
    
    $auditType = Read-Host "Enter selection"
    
    # Get all assets
    Write-Host "Retrieving assets from Snipe-IT..." -ForegroundColor Yellow
    $assets = Get-SnipeITAssets -Limit 1000
    $filteredAssets = @()
    
    switch ($auditType) {
        "1" {
            $filteredAssets = $assets
            $reportTitle = "All Assets"
        }
        "2" {
            # Filter by location
            $locations = Get-SnipeITLocations
            
            Write-Host "Available Locations:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $locations.Count; $i++) {
                Write-Host "[$i] $($locations[$i].name)"
            }
            
            $locationIndex = Read-Host "Enter the number of the location"
            if (-not ($locationIndex -match '^\d+) -or [int]$locationIndex -ge $locations.Count) {
                Write-Host "Invalid selection." -ForegroundColor Red
                Read-Host "Press Enter to continue"
                return
            }
            
            $selectedLocation = $locations[[int]$locationIndex]
            $filteredAssets = $assets | Where-Object { $_.rtd_location.id -eq $selectedLocation.id }
            $reportTitle = "Assets at $($selectedLocation.name)"
        }
        "3" {
            # Filter by model
            $models = Get-SnipeITModels
            
            Write-Host "Available Models:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $models.Count; $i++) {
                Write-Host "[$i] $($models[$i].name)"
            }
            
            $modelIndex = Read-Host "Enter the number of the model"
            if (-not ($modelIndex -match '^\d+) -or [int]$modelIndex -ge $models.Count) {
                Write-Host "Invalid selection." -ForegroundColor Red
                Read-Host "Press Enter to continue"
                return
            }
            
            $selectedModel = $models[[int]$modelIndex]
            $filteredAssets = $assets | Where-Object { $_.model.id -eq $selectedModel.id }
            $reportTitle = "$($selectedModel.name) Assets"
        }
        "4" {
            # Filter by status
            $statuses = Get-SnipeITStatusLabels
            
            Write-Host "Available Statuses:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $statuses.Count; $i++) {
                Write-Host "[$i] $($statuses[$i].name)"
            }
            
            $statusIndex = Read-Host "Enter the number of the status"
            if (-not ($statusIndex -match '^\d+) -or [int]$statusIndex -ge $statuses.Count) {
                Write-Host "Invalid selection." -ForegroundColor Red
                Read-Host "Press Enter to continue"
                return
            }
            
            $selectedStatus = $statuses[[int]$statusIndex]
            $filteredAssets = $assets | Where-Object { $_.status_label.id -eq $selectedStatus.id }
            $reportTitle = "Assets with $($selectedStatus.name) Status"
        }
        default {
            Write-Host "Invalid selection." -ForegroundColor Red
            Read-Host "Press Enter to continue"
            return
        }
    }
    
    if ($filteredAssets.Count -eq 0) {
        Write-Host "No assets found matching the selected criteria." -ForegroundColor Yellow
        Read-Host "Press Enter to continue"
        return
    }
    
    # Create report
    $reportPath = "$env:USERPROFILE\Desktop\AssetAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    $reportData = @()
    
    foreach ($asset in $filteredAssets) {
        $reportData += [PSCustomObject]@{
            AssetTag = $asset.asset_tag
            Name = $asset.name
            Model = $asset.model.name
            SerialNumber = $asset.serial
            Status = $asset.status_label.name
            Category = $asset.category.name
            Manufacturer = $asset.manufacturer.name
            Location = $asset.rtd_location.name
            PurchaseDate = $asset.purchase_date
            LastCheckin = $asset.last_checkin
            LastCheckout = $asset.last_checkout
            AssignedTo = if ($asset.assigned_to) { $asset.assigned_to.name } else { "Unassigned" }
            Notes = $asset.notes
        }
    }
    
    # Export to CSV
    $reportData | Export-Csv -Path $reportPath -NoTypeInformation
    
    # Display summary
    Write-Host "Report Summary:" -ForegroundColor Cyan
    Write-Host "- Report Type: $reportTitle" -ForegroundColor Yellow
    Write-Host "- Total Assets: $($filteredAssets.Count)" -ForegroundColor Yellow
    Write-Host "- Report saved to: $reportPath" -ForegroundColor Green
    
    Read-Host "Press Enter to continue"
}

# Function to generate report templates
function Generate-ReportTemplates {
    Clear-Host
    Write-Host "=== Generate Report Templates ===" -ForegroundColor Cyan
    
    # Define template folder
    $templateFolder = "$env:USERPROFILE\Documents\SnipeIT_Templates"
    
    # Create folder if it doesn't exist
    if (-not (Test-Path -Path $templateFolder)) {
        New-Item -Path $templateFolder -ItemType Directory | Out-Null
    }
    
    # Create report templates
    
    # 1. Broken Assets Template
    $brokenTemplate = @"
# Broken Assets Report
# Generated: {0}

This report provides details on all assets currently marked as broken or in need of repair.

## Summary Statistics
- Total Broken Assets: {1}
- Estimated Repair Cost: {2}
- Average Age of Broken Assets: {3} days

## Asset Details
{4}

## Recommendations
- Schedule maintenance for assets with similar issues
- Consider replacing assets that have been broken multiple times
- Evaluate vendor quality for repeating failures
"@
    
    $brokenTemplatePath = "$templateFolder\BrokenAssetsTemplate.md"
    $brokenTemplate -f (Get-Date), "[Count]", "[Cost]", "[Age]", "[Table Placeholder]" | Out-File -FilePath $brokenTemplatePath
    
    # 2. Asset Activity Template
    $activityTemplate = @"
# Asset Activity Report
# Period: {0} to {1}

This report summarizes all asset movements and status changes during the specified period.

## Summary Statistics
- Total Activities: {2}
- Checkouts: {3}
- Checkins: {4}
- Status Changes: {5}
- Location Changes: {6}

## Activity Details
{7}

## Trends and Observations
- Assets with highest activity volume
- Unusual patterns or anomalies
- Recommended follow-ups
"@
    
    $activityTemplatePath = "$templateFolder\ActivityReportTemplate.md"
    $activityTemplate -f "[Start Date]", "[End Date]", "[Count]", "[Checkout Count]", "[Checkin Count]", "[Status Count]", "[Location Count]", "[Table Placeholder]" | Out-File -FilePath $activityTemplatePath
    
    # 3. Asset Audit Template
    $auditTemplate = @"
# Asset Audit Report
# Date: {0}
# Scope: {1}

This report provides a comprehensive inventory of assets according to specified criteria.

## Audit Summary
- Total Assets: {2}
- Total Value: {3}
- Assets per Category:
  {4}

## Location Distribution
{5}

## Status Distribution
{6}

## Recommendations
- Areas requiring attention
- Suggested inventory adjustments
- Compliance considerations
"@
    
    $auditTemplatePath = "$templateFolder\AssetAuditTemplate.md"
    $auditTemplate -f (Get-Date), "[Audit Scope]", "[Count]", "[Value]", "[Category List]", "[Location Chart]", "[Status Chart]" | Out-File -FilePath $auditTemplatePath
    
    # 4. Issues Report Template
    $issuesTemplate = @"
# Asset Issues Report
# Generated: {0}

This report compiles all assets with documented issues or problems.

## Summary
- Total Assets with Issues: {1}
- Critical Issues: {2}
- Moderate Issues: {3}
- Minor Issues: {4}

## Issue Categories
{5}

## Detailed Issue Log
{6}

## Action Items
- Critical issues requiring immediate attention
- Scheduled maintenance recommendations
- Replacement considerations
"@
    
    $issuesTemplatePath = "$templateFolder\IssuesReportTemplate.md"
    $issuesTemplate -f (Get-Date), "[Count]", "[Critical]", "[Moderate]", "[Minor]", "[Categories]", "[Issues Log]" | Out-File -FilePath $issuesTemplatePath
    
    # 5. Custom Excel Template
    $excelTemplatePath = "$templateFolder\AssetReportTemplate.xlsx"
    
    try {
        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        
        # Create Sheets
        $summarySheet = $workbook.Worksheets.Item(1)
        $summarySheet.Name = "Summary"
        $detailSheet = $workbook.Worksheets.Add()
        $detailSheet.Name = "Asset Details"
        $maintenanceSheet = $workbook.Worksheets.Add()
        $maintenanceSheet.Name = "Maintenance Log"
        
        # Format Summary Sheet
        $summarySheet.Cells.Item(1, 1).Value = "Snipe-IT Asset Report"
        $summarySheet.Cells.Item(1, 1).Font.Size = 16
        $summarySheet.Cells.Item(1, 1).Font.Bold = $true
        
        $summarySheet.Cells.Item(3, 1).Value = "Report Date:"
        $summarySheet.Cells.Item(3, 2).Value = Get-Date -Format "MM/dd/yyyy"
        
        $summarySheet.Cells.Item(4, 1).Value = "Total Assets:"
        $summarySheet.Cells.Item(4, 2).Value = "[ASSET_COUNT]"
        
        $summarySheet.Cells.Item(5, 1).Value = "Total Value:"
        $summarySheet.Cells.Item(5, 2).Value = "[ASSET_VALUE]"
        
        $summarySheet.Cells.Item(7, 1).Value = "Asset Status Breakdown"
        $summarySheet.Cells.Item(7, 1).Font.Bold = $true
        
        $summarySheet.Cells.Item(8, 1).Value = "Status"
        $summarySheet.Cells.Item(8, 2).Value = "Count"
        $summarySheet.Cells.Item(8, 3).Value = "Percentage"
        
        # Format Detail Sheet
        $detailSheet.Cells.Item(1, 1).Value = "Asset Tag"
        $detailSheet.Cells.Item(1, 2).Value = "Name"
        $detailSheet.Cells.Item(1, 3).Value = "Model"
        $detailSheet.Cells.Item(1, 4).Value = "Serial Number"
        $detailSheet.Cells.Item(1, 5).Value = "Status"
        $detailSheet.Cells.Item(1, 6).Value = "Location"
        $detailSheet.Cells.Item(1, 7).Value = "Assigned To"
        $detailSheet.Cells.Item(1, 8).Value = "Purchase Date"
        $detailSheet.Cells.Item(1, 9).Value = "Purchase Price"
        $detailSheet.Cells.Item(1, 10).Value = "Last Updated"
        
        # Format header row
        $headerRange = $detailSheet.Range("A1:J1")
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        # Format Maintenance Log Sheet
        $maintenanceSheet.Cells.Item(1, 1).Value = "Asset Tag"
        $maintenanceSheet.Cells.Item(1, 2).Value = "Date"
        $maintenanceSheet.Cells.Item(1, 3).Value = "Type"
        $maintenanceSheet.Cells.Item(1, 4).Value = "Notes"
        $maintenanceSheet.Cells.Item(1, 5).Value = "Cost"
        $maintenanceSheet.Cells.Item(1, 6).Value = "Performed By"
        
        # Format header row
        $maintenanceHeaderRange = $maintenanceSheet.Range("A1:F1")
        $maintenanceHeaderRange.Font.Bold = $true
        $maintenanceHeaderRange.Interior.ColorIndex = 15
        
        # Save the workbook
        $workbook.SaveAs($excelTemplatePath)
        $workbook.Close()
        $excel.Quit()
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        Write-Host "Could not create Excel template. Error: $_" -ForegroundColor Red
        # Create a simple CSV template instead
        $csvTemplate = "Asset Tag,Name,Model,Serial Number,Status,Location,Assigned To,Purchase Date,Purchase Price,Last Updated`n"
        $csvTemplate | Out-File -FilePath "$templateFolder\AssetReportTemplate.csv"
    }
    
    # Display summary
    Write-Host "Report Templates Generated:" -ForegroundColor Cyan
    Write-Host "- Broken Assets Template: $brokenTemplatePath" -ForegroundColor Green
    Write-Host "- Activity Report Template: $activityTemplatePath" -ForegroundColor Green
    Write-Host "- Asset Audit Template: $auditTemplatePath" -ForegroundColor Green
    Write-Host "- Issues Report Template: $issuesTemplatePath" -ForegroundColor Green
    
    if (Test-Path -Path $excelTemplatePath) {
        Write-Host "- Excel Asset Report Template: $excelTemplatePath" -ForegroundColor Green
    } else {
        Write-Host "- CSV Asset Report Template: $templateFolder\AssetReportTemplate.csv" -ForegroundColor Green
    }
    
    Read-Host "Press Enter to continue"
}
#endregion

#region Main Menu Function
# Main menu function
function Show-MainMenu {
    $continue = $true
    
    while ($continue) {
        Clear-Host
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host "          SNIPE-IT AUTOMATION TOOL            " -ForegroundColor Cyan
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Asset Management:" -ForegroundColor Yellow
        Write-Host "  1. Assign Laptop to User"
        Write-Host "  2. Mass Checkin/Checkout Laptops"
        Write-Host "  3. Handle Broken Laptop Workflow"
        Write-Host "  4. Move Asset Location"
        Write-Host ""
        Write-Host "User Management:" -ForegroundColor Yellow
        Write-Host "  5. Manage Active User Group"
        Write-Host ""
        Write-Host "Reporting:" -ForegroundColor Yellow
        Write-Host "  6. Reporting Menu (Broken/Temp, Issues, Activity, Audit)"
        Write-Host ""
        Write-Host "System:" -ForegroundColor Yellow
        Write-Host "  7. Settings"
        Write-Host "  0. Exit"
        Write-Host ""
        
        $choice = Read-Host "Select an option"
        
        switch ($choice) {
            "1" { Assign-Laptop }
            "2" { Invoke-MassLaptopAction }
            "3" { Handle-BrokenLaptop }
            "4" { Move-AssetLocation }
            "5" { Manage-ActiveUserGroup }
            "6" { Show-ReportingMenu }
            "7" { Show-SettingsMenu }
            "0" { $continue = $false }
            default { 
                Write-Host "Invalid option selected." -ForegroundColor Red
                Read-Host "Press Enter to continue"
            }
        }
    }
}

# Settings menu function
function Show-SettingsMenu {
    Clear-Host
    Write-Host "=== Settings ===" -ForegroundColor Cyan
    Write-Host "1. Update API URL and Token"
    Write-Host "2. Test API Connection"
    Write-Host "3. Save Configuration"
    Write-Host "4. Load Configuration"
    Write-Host "0. Return to Main Menu"
    
    $choice = Read-Host "Select an option"
    
    switch ($choice) {
        "1" { 
            $script:snipeITURL = Read-Host "Enter Snipe-IT API URL (e.g., https://your-snipeit-instance.com/api/v1)"
            $script:apiToken = Read-Host "Enter API Token"
            Write-Host "API settings updated." -ForegroundColor Green
        }
        "2" { 
            Write-Host "Testing API connection..." -ForegroundColor Yellow
            $response = Invoke-SnipeITAPI -Endpoint "statuslabels"
            if ($response) {
                Write-Host "Connection successful!" -ForegroundColor Green
            } else {
                Write-Host "Connection failed. Please check your API URL and token." -ForegroundColor Red
            }
        }
        "3" {
            $configPath = "$env:USERPROFILE\Documents\SnipeIT_Config.xml"
            $config = @{
                "APIURL" = $script:snipeITURL
                "APIToken" = $script:apiToken
            }
            
            $config | Export-Clixml -Path $configPath
            Write-Host "Configuration saved to $configPath" -ForegroundColor Green
        }
        "4" {
            $configPath = "$env:USERPROFILE\Documents\SnipeIT_Config.xml"
            
            if (Test-Path -Path $configPath) {
                $config = Import-Clixml -Path $configPath
                $script:snipeITURL = $config.APIURL
                $script:apiToken = $config.APIToken
                Write-Host "Configuration loaded successfully." -ForegroundColor Green
            } else {
                Write-Host "No configuration file found at $configPath" -ForegroundColor Yellow
            }
        }
        "0" { return }
        default { 
            Write-Host "Invalid option selected." -ForegroundColor Red
        }
    }
    
    Read-Host "Press Enter to continue"
}
#endregion

#region Script Initialization
# Initialize script
function Initialize-Script {
    # Check for saved configuration
    $configPath = "$env:USERPROFILE\Documents\SnipeIT_Config.xml"
    
    if (Test-Path -Path $configPath) {
        $loadConfig = Read-Host "Found saved configuration. Load it? (Y/N)"
        
        if ($loadConfig -eq "Y" -or $loadConfig -eq "y") {
            $config = Import-Clixml -Path $configPath
            $script:snipeITURL = $config.APIURL
            $script:apiToken = $config.APIToken
            Write-Host "Configuration loaded successfully." -ForegroundColor Green
        } else {
            # Prompt for API settings
            $script:snipeITURL = Read-Host "Enter Snipe-IT API URL (e.g., https://your-snipeit-instance.com/api/v1)"
            $script:apiToken = Read-Host "Enter API Token"
        }
    } else {
        # Prompt for API settings
        $script:snipeITURL = Read-Host "Enter Snipe-IT API URL (e.g., https://your-snipeit-instance.com/api/v1)"
        $script:apiToken = Read-Host "Enter API Token"
    }
    
    # Test API connection
    Write-Host "Testing API connection..." -ForegroundColor Yellow
    $response = Invoke-SnipeITAPI -Endpoint "statuslabels"
    
    if ($response) {
        Write-Host "Connection successful!" -ForegroundColor Green
        Read-Host "Press Enter to continue to the main menu"
    } else {
        Write-Host "Connection failed. Please check your API URL and token." -ForegroundColor Red
        $retry = Read-Host "Would you like to retry? (Y/N)"
        
        if ($retry -eq "Y" -or $retry -eq "y") {
            Initialize-Script
        } else {
            Write-Host "You can update API settings from the Settings menu." -ForegroundColor Yellow
            Read-Host "Press Enter to continue to the main menu"
        }
    }
}
#endregion

# Start script
Write-Host "Starting Snipe-IT Automation Tool..." -ForegroundColor Cyan
Initialize-Script
Show-MainMenu
