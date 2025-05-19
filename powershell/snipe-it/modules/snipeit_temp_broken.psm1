<#
.SYNOPSIS
    Snipe-IT Temporary and Broken Device Management Module
.DESCRIPTION
    PowerShell module for managing temporary laptop assignments and broken device tracking
.NOTES
    Version:        1.1
    Author:         Your Name
    Creation Date:  May 19, 2025
    Requires:       SnipeitPS Module
#>

# Temporary Laptop Management Functions

function Show-TempBrokenMenu {
    Clear-Host
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "   TEMP & BROKEN DEVICE MANAGEMENT        " -ForegroundColor Cyan
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "1. Assign Temporary Laptop"
    Write-Host "2. Report Broken Laptop (with replacement)"
    Write-Host "3. View Broken Laptop Statistics"
    Write-Host "4. View Today's Temp Laptop Assignments"
    Write-Host "5. Return Temporary Laptop"
    Write-Host "6. Return to Main Menu"
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-6]: " -ForegroundColor Yellow -NoNewline
}

function Get-UserByEmailOrEmployeeNum {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserIdentifier
    )
    
    try {
        # Search by email or employee number
        $user = Get-SnipeitUser -search $UserIdentifier
        
        if ($user) {
            # Verify the search result matches what we're looking for
            # Check if it matches email
            if ($user.email -eq $UserIdentifier) {
                return $user
            }
            # Check if it matches employee number
            if ($user.employee_num -eq $UserIdentifier) {
                return $user
            }
            # If search returned a result but doesn't exactly match, it might be a partial match
            # Let's still return it but inform the user
            Write-Host "Found user: $($user.name) - please verify this is correct" -ForegroundColor Yellow
            return $user
        }
        
        Write-Host "User not found with email or ID number: $UserIdentifier" -ForegroundColor Red
        return $null
    }
    catch {
        Write-Host "Error retrieving user with identifier $UserIdentifier : $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Get-AssetByTag {
    param(
        [Parameter(Mandatory=$true)]
        [string]$AssetTag
    )
    
    try {
        $asset = Get-SnipeitAsset -asset_tag $AssetTag
        return $asset
    }
    catch {
        Write-Host "Error retrieving asset with tag $AssetTag : $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function AssignTempLaptop {
    Write-Host "`n===========================================" -ForegroundColor Green
    Write-Host "        ASSIGN TEMPORARY LAPTOP            " -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    
    # Get User identifier (email or employee ID number)
    $userIdentifier = Read-Host "`nEnter User Email or ID Number (Employee Number)"
    if ([string]::IsNullOrWhiteSpace($userIdentifier)) {
        Write-Host "User email or ID number cannot be empty." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Validate user exists using enhanced lookup
    $user = Get-UserByEmailOrEmployeeNum -UserIdentifier $userIdentifier
    if (-not $user) {
        Start-Sleep -Seconds 3
        return
    }
    
    Write-Host "User found: $($user.name) ($($user.email)) [Employee #: $($user.employee_num)]" -ForegroundColor Green
    
    # Get Asset Tag
    $assetTag = Read-Host "Enter Temporary Laptop Asset Tag"
    if ([string]::IsNullOrWhiteSpace($assetTag)) {
        Write-Host "Asset tag cannot be empty." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Validate asset exists and is available
    $asset = Get-AssetByTag -AssetTag $assetTag
    if (-not $asset) {
        Write-Host "Asset with tag $assetTag not found." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Check if asset is available
    if ($asset.assigned_to) {
        Write-Host "Asset $assetTag is already assigned to someone else." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    Write-Host "Asset found: $($asset.name) - $($asset.model)" -ForegroundColor Green
    
    # Get expected return date
    $dueDate = Read-Host "Enter expected return date (YYYY-MM-DD) or press Enter for no due date"
    
    try {
        # Set checkout date to current date and time
        $checkoutDate = Get-Date
        
        # Checkout asset to user using the correct syntax
        $checkoutParams = @{
            id = $asset.id
            assigned_id = $user.id
            checkout_to_type = "user"
            note = "TEMPORARY ASSIGNMENT"
            checkout_at = $checkoutDate
        }
        
        if (-not [string]::IsNullOrWhiteSpace($dueDate)) {
            $checkoutParams['expected_checkin'] = [DateTime]::Parse($dueDate)
        }
        
        $result = Set-SnipeitAssetOwner @checkoutParams
        
        if ($result) {
            Write-Host "`nSuccess! Temporary laptop $assetTag assigned to user $($user.name)" -ForegroundColor Green
            Write-Host "Checkout Date: $($checkoutDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Green
            if ($dueDate) {
                Write-Host "Expected Return: $dueDate" -ForegroundColor Green
            }
        }
        else {
            Write-Host "Failed to assign temporary laptop." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error assigning temporary laptop: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}

function ReportBrokenLaptop {
    Write-Host "`n===========================================" -ForegroundColor Yellow
    Write-Host "      REPORT BROKEN LAPTOP                " -ForegroundColor Yellow
    Write-Host "===========================================" -ForegroundColor Yellow
    
    # Get User identifier (email or employee ID number)
    $userIdentifier = Read-Host "`nEnter User Email or ID Number (Employee Number)"
    if ([string]::IsNullOrWhiteSpace($userIdentifier)) {
        Write-Host "User email or ID number cannot be empty." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Validate user exists using enhanced lookup
    $user = Get-UserByEmailOrEmployeeNum -UserIdentifier $userIdentifier
    if (-not $user) {
        Start-Sleep -Seconds 3
        return
    }
    
    Write-Host "User found: $($user.name) ($($user.email)) [Employee #: $($user.employee_num)]" -ForegroundColor Green
    
    # Get broken laptop asset tag
    $brokenAssetTag = Read-Host "Enter Broken Laptop Asset Tag"
    if ([string]::IsNullOrWhiteSpace($brokenAssetTag)) {
        Write-Host "Broken laptop asset tag cannot be empty." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Validate broken asset exists
    $brokenAsset = Get-AssetByTag -AssetTag $brokenAssetTag
    if (-not $brokenAsset) {
        Write-Host "Broken asset with tag $brokenAssetTag not found." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    Write-Host "Broken Asset: $($brokenAsset.name) - $($brokenAsset.model)" -ForegroundColor Yellow
    
    # Get replacement laptop asset tag
    $newAssetTag = Read-Host "Enter Replacement Laptop Asset Tag"
    if ([string]::IsNullOrWhiteSpace($newAssetTag)) {
        Write-Host "Replacement laptop asset tag cannot be empty." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Validate replacement asset exists and is available
    $newAsset = Get-AssetByTag -AssetTag $newAssetTag
    if (-not $newAsset) {
        Write-Host "Replacement asset with tag $newAssetTag not found." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    if ($newAsset.assigned_to) {
        Write-Host "Replacement asset $newAssetTag is already assigned to someone else." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    Write-Host "Replacement Asset: $($newAsset.name) - $($newAsset.model)" -ForegroundColor Green
    
    # Get issue description
    $issueDescription = Read-Host "Describe what is broken"
    if ([string]::IsNullOrWhiteSpace($issueDescription)) {
        Write-Host "Issue description cannot be empty." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Get supplier ID for maintenance record
    $supplierIdInput = Read-Host "Enter Supplier ID for maintenance record (or press Enter for default: 1)"
    $supplierId = if ([string]::IsNullOrWhiteSpace($supplierIdInput)) { 1 } else { [int]$supplierIdInput }
    
    try {
        # Step 1: Check in broken laptop using Reset-SnipeitAssetOwner
        Write-Host "Checking in broken laptop..." -ForegroundColor Yellow
        $checkinResult = Reset-SnipeitAssetOwner -id $brokenAsset.id -status_id 6 -location_id 24 -note "BROKEN - $issueDescription - Replaced with $newAssetTag"
        
        if ($checkinResult) {
            Write-Host "✓ Broken laptop $brokenAssetTag checked in successfully" -ForegroundColor Green
            Write-Host "✓ Status set to broken (ID: 6)" -ForegroundColor Green
            Write-Host "✓ Location set to 24" -ForegroundColor Green
            
            # Step 2: Create maintenance record
            Write-Host "Creating maintenance record..." -ForegroundColor Yellow
            try {
                $maintenanceParams = @{
                    asset_id = $brokenAsset.id
                    supplier_id = $supplierId
                    asset_maintenance_type = "Repair"
                    title = "Laptop Repair - $issueDescription"
                    start_date = Get-Date
                    notes = "Broken laptop reported by $($user.name). Issue: $issueDescription. Replaced with asset $newAssetTag."
                }
                
                $maintenanceResult = New-SnipeitAssetMaintenance @maintenanceParams
                
                if ($maintenanceResult) {
                    Write-Host "✓ Maintenance record created successfully" -ForegroundColor Green
                }
                else {
                    Write-Host "⚠ Warning: Failed to create maintenance record" -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "⚠ Warning: Error creating maintenance record: $($_.Exception.Message)" -ForegroundColor Yellow
            }
            
            # Step 3: Checkout replacement laptop
            Write-Host "Assigning replacement laptop..." -ForegroundColor Yellow
            $checkoutDate = Get-Date
            
            $checkoutParams = @{
                id = $newAsset.id
                assigned_id = $user.id
                checkout_to_type = "user"
                note = "Replacement for broken laptop $brokenAssetTag"
                checkout_at = $checkoutDate
            }
            
            $checkoutResult = Set-SnipeitAssetOwner @checkoutParams
            
            if ($checkoutResult) {
                Write-Host "`n🎉 SUCCESS! Replacement process completed:" -ForegroundColor Green
                Write-Host "├─ Broken laptop $brokenAssetTag: Checked in to location 24, status set to broken" -ForegroundColor Green
                Write-Host "├─ Maintenance record: Created for repair tracking" -ForegroundColor Green
                Write-Host "├─ Replacement laptop $newAssetTag: Assigned to $($user.name)" -ForegroundColor Green
                Write-Host "├─ Issue description: $issueDescription" -ForegroundColor Green
                Write-Host "└─ Checkout date: $($checkoutDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Green
            }
            else {
                Write-Host "⚠ Warning: Broken laptop processed, but failed to assign replacement" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "❌ Failed to check in broken laptop" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "❌ Error processing broken laptop: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Full error details: $($_.Exception.ToString())" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}

function ViewBrokenLaptopStats {
    Write-Host "`n===========================================" -ForegroundColor Magenta
    Write-Host "      BROKEN LAPTOP STATISTICS            " -ForegroundColor Magenta
    Write-Host "===========================================" -ForegroundColor Magenta
    
    try {
        # Get all assets with status indicating they're broken (status ID 6)
        $brokenAssets = Get-SnipeitAsset -status_id 6
        
        if ($brokenAssets -and $brokenAssets.Count -gt 0) {
            Write-Host "`nTotal Broken Laptops: $($brokenAssets.Count)" -ForegroundColor Red
            Write-Host "`nBroken Assets Details:" -ForegroundColor White
            Write-Host "=" * 80 -ForegroundColor Gray
            
            foreach ($asset in $brokenAssets) {
                Write-Host "Asset Tag: $($asset.asset_tag)" -ForegroundColor Yellow
                Write-Host "Name: $($asset.name)" -ForegroundColor White
                Write-Host "Model: $($asset.model.name)" -ForegroundColor White
                Write-Host "Serial: $($asset.serial)" -ForegroundColor White
                Write-Host "Status: $($asset.status_label.name)" -ForegroundColor Red
                
                # Get asset history to find latest notes
                try {
                    $assetDetails = Get-SnipeitAsset -id $asset.id
                    if ($assetDetails.notes) {
                        Write-Host "Notes: $($assetDetails.notes)" -ForegroundColor Cyan
                    }
                }
                catch {
                    Write-Host "Could not retrieve detailed notes for this asset." -ForegroundColor Gray
                }
                
                Write-Host "-" * 80 -ForegroundColor Gray
            }
        }
        else {
            Write-Host "`nNo broken laptops found in the system." -ForegroundColor Green
        }
        
        # Additional statistics
        Write-Host "`nSummary Statistics:" -ForegroundColor White
        Write-Host "Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Gray
    }
    catch {
        Write-Host "Error retrieving broken laptop statistics: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}

function ViewTodayTempAssignments {
    Write-Host "`n===========================================" -ForegroundColor Blue
    Write-Host "    TODAY'S TEMP LAPTOP ASSIGNMENTS       " -ForegroundColor Blue
    Write-Host "===========================================" -ForegroundColor Blue
    
    try {
        $today = Get-Date -Format "yyyy-MM-dd"
        
        # Get all checked out assets
        $allAssets = Get-SnipeitAsset -status "Deployed"
        
        # Filter for temp assignments made today
        $tempAssignments = @()
        
        foreach ($asset in $allAssets) {
            # Get asset history to check for temp assignments
            try {
                $assetDetails = Get-SnipeitAsset -id $asset.id
                
                # Check if the asset has notes indicating it's a temp assignment
                if ($assetDetails.notes -like "*TEMPORARY ASSIGNMENT*") {
                    # Check if checkout date is today
                    if ($assetDetails.last_checkout -like "$today*") {
                        $tempAssignments += $assetDetails
                    }
                }
            }
            catch {
                # Continue with next asset if there's an error
                continue
            }
        }
        
        if ($tempAssignments.Count -gt 0) {
            Write-Host "`nTemp Laptop Assignments Today: $($tempAssignments.Count)" -ForegroundColor Green
            Write-Host "`nToday's Temporary Assignments:" -ForegroundColor White
            Write-Host "=" * 100 -ForegroundColor Gray
            
            foreach ($assignment in $tempAssignments) {
                Write-Host "Asset Tag: $($assignment.asset_tag)" -ForegroundColor Yellow
                Write-Host "Laptop: $($assignment.name)" -ForegroundColor White
                Write-Host "Model: $($assignment.model.name)" -ForegroundColor White
                Write-Host "Assigned To: $($assignment.assigned_to.name)" -ForegroundColor Cyan
                Write-Host "Checkout Time: $($assignment.last_checkout)" -ForegroundColor Green
                
                if ($assignment.expected_checkin) {
                    Write-Host "Expected Return: $($assignment.expected_checkin)" -ForegroundColor Magenta
                }
                
                if ($assignment.notes) {
                    Write-Host "Notes: $($assignment.notes)" -ForegroundColor Gray
                }
                
                Write-Host "-" * 100 -ForegroundColor Gray
            }
        }
        else {
            Write-Host "`nNo temporary laptop assignments found for today ($today)." -ForegroundColor Yellow
        }
        
        Write-Host "`nReport Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Gray
    }
    catch {
        Write-Host "Error retrieving today's temp assignments: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}

function ReturnTempLaptop {
    Write-Host "`n===========================================" -ForegroundColor Green
    Write-Host "       RETURN TEMPORARY LAPTOP            " -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    
    # Get Asset Tag
    $assetTag = Read-Host "`nEnter Temporary Laptop Asset Tag to return"
    if ([string]::IsNullOrWhiteSpace($assetTag)) {
        Write-Host "Asset tag cannot be empty." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Validate asset exists
    $asset = Get-AssetByTag -AssetTag $assetTag
    if (-not $asset) {
        Write-Host "Asset with tag $assetTag not found." -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
    
    # Check if asset is currently assigned
    if (-not $asset.assigned_to) {
        Write-Host "Asset $assetTag is not currently assigned to anyone." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        return
    }
    
    Write-Host "Asset: $($asset.name) - $($asset.model.name)" -ForegroundColor White
    Write-Host "Currently assigned to: $($asset.assigned_to.name)" -ForegroundColor Cyan
    
    # Confirm it's a temp assignment
    if ($asset.notes -notlike "*TEMPORARY ASSIGNMENT*") {
        Write-Host "Warning: This asset is not marked as a temporary assignment." -ForegroundColor Yellow
        $confirm = Read-Host "Continue with check-in anyway? (y/N)"
        if ($confirm -ne 'y' -and $confirm -ne 'Y') {
            return
        }
    }
    
    # Get return notes
    $returnNotes = Read-Host "Enter return condition/notes (optional)"
    
    try {
        # Check in the asset using Reset-SnipeitAssetOwner (to location 24)
        $noteText = if ($returnNotes) { "TEMP RETURN: $returnNotes" } else { "TEMP RETURN" }
        $checkinResult = Reset-SnipeitAssetOwner -id $asset.id -location_id 24 -note $noteText
        
        if ($checkinResult) {
            Write-Host "`nSuccess! Temporary laptop $assetTag returned successfully." -ForegroundColor Green
            Write-Host "Return Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
            Write-Host "Returned by: $($asset.assigned_to.name)" -ForegroundColor Green
            Write-Host "Checked into location 24" -ForegroundColor Green
            if ($returnNotes) {
                Write-Host "Return Notes: $returnNotes" -ForegroundColor Green
            }
        }
        else {
            Write-Host "Failed to return temporary laptop." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error returning temporary laptop: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Full error details: $($_.Exception.ToString())" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}

function Process-TempBrokenMenu {
    $continue = $true
    
    while ($continue) {
        Show-TempBrokenMenu
        $selection = Read-Host
        
        switch ($selection) {
            "1" { AssignTempLaptop }
            "2" { ReportBrokenLaptop }
            "3" { ViewBrokenLaptopStats }
            "4" { ViewTodayTempAssignments }
            "5" { ReturnTempLaptop }
            "6" { $continue = $false }
            default { 
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep -Seconds 2 
            }
        }
    }
}

# Export the main function that will be called from the menu script
Export-ModuleMember -Function Process-TempBrokenMenu
