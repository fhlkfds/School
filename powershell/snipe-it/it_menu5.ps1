# Check if SnipeitPS module is installed and import it
function Initialize-SnipeIT
{
    if (-not (Get-Module -ListAvailable -Name SnipeitPS))
    {
        Write-Host "SnipeitPS module is not installed. Installing..." -ForegroundColor Yellow
        try
        {
            Install-Module -Name SnipeitPS -Force -Scope CurrentUser
            Write-Host "SnipeitPS module installed successfully." -ForegroundColor Green
        } catch
        {
            Write-Host "Failed to install SnipeitPS module. Please install it manually with 'Install-Module -Name SnipeitPS'." -ForegroundColor Red
            return $false
        }
    }
    
    Import-Module SnipeitPS
    
    # Use $PSScriptRoot for better reliability
    $scriptDir = $PSScriptRoot
    $moduleDir = Join-Path $scriptDir "modules"
    
    # Import the check-in/check-out module
    $checkModulePath = Join-Path $moduleDir "snipeit_check.psm1"
    if (Test-Path $checkModulePath)
    {
        try
        {
            Import-Module $checkModulePath -Force -ErrorAction Stop
            Write-Host "Successfully imported snipeit_check module." -ForegroundColor Green
        } catch
        {
            Write-Host "Error importing snipeit_check module: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Please check the module file for syntax errors." -ForegroundColor Yellow
            return $false
        }
    } else
    {
        Write-Host "Warning: snipeit_check.psm1 module not found in modules directory." -ForegroundColor Yellow
        Write-Host "Expected path: $checkModulePath" -ForegroundColor Yellow
        Write-Host "Some functions will not be available." -ForegroundColor Yellow
    }
    
    # Import the temporary and broken device management module
    $tempBrokenModulePath = Join-Path $moduleDir "snipeit_temp_broken.psm1"
    if (Test-Path $tempBrokenModulePath)
    {
        try
        {
            Import-Module $tempBrokenModulePath -Force -ErrorAction Stop
            Write-Host "Successfully imported snipeit_temp_broken module." -ForegroundColor Green
        } catch
        {
            Write-Host "Error importing snipeit_temp_broken module: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Please check the module file for syntax errors." -ForegroundColor Yellow
            return $false
        }
    } else
    {
        Write-Host "Warning: snipeit_temp_broken.psm1 module not found in modules directory." -ForegroundColor Yellow
        Write-Host "Expected path: $tempBrokenModulePath" -ForegroundColor Yellow
        Write-Host "Temp/Broken Device Management will not be available." -ForegroundColor Yellow
    }
    
    # Set Snipe-IT API parameters - Replace with your actual values
    $apiUrl = "http://192.168.122.217"
    $apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiZTIzMTYxMzZmODdlMGRmNWVkYjA2YTRmMTAxYWU1ZjBlMzFmYTY1Y2ExYTU1MGZiYWFhM2RiOGE4ZTZmMTMxYTFlMzk3ZThiNmY3ODVjZDYiLCJpYXQiOjE3NDc2ODg3MzkuMTIxMjk2LCJuYmYiOjE3NDc2ODg3MzkuMTIxMjk4LCJleHAiOjIyMjEwNzQzMzkuMTA4NTE3LCJzdWIiOiIxIiwic2NvcGVzIjpbXX0.rDF7Sj38TULBJY_NRnQ5ODE_65qb2FzvhmEu_8FMHakKYdjKSZjqxq8Qkrz6gN3o96MOfAPBldUWQIs4-efqdhOxzxEI5T1dq1h35s7enTHgz7UzFQGJALl_VLXUB6KuBQrpUJT1f_y9BeEt0NG3ZFMGa8G9v7a7tzJv1h5xrHFEZsx_3bjI_wv6ZrXPn9YJedMkSeAHO8IfEW3PvzOCzAOBvKrdtuXIPTv-EL6F4e9CfuK5BVhzW94JEmDqaLbsQwDFHVJcau-Ij9v1nFM9-ek7Lzfrn8PaFcalpmyND-q3IgiSV3yiYTJzvlsNCi_KYavFLFWnzflxtqMEVvqHDBQGr9Qk5JsNXpMAPZVgmIlRXCzqdDd6E00RVZJ79w44zAWg8Y3eoEMqTKVJFAz3ZYY5D4gD73yb3lP7V4sJ_bGZbV7DkKdWOZdiDNM7LcWbcZQw4W6CjAUt3Eo0yepwZlNrqY9af-CK4uyNkYvxPK5TtexN_2vfMKbqCC94cuj-t3pZ4V5IJS4ZXVKkDBsMJ6wRHdIWAB37z3JZ8RDvRFw-RXVsT9aSpI0F7N3xwm4Iw-al6dZQCVmsbyaQRBJahVVde0Yu50gBWtiaKPepBW0iBh89Ry6KaL82TUq4p72WDl41PTaSKMV7j-E6_mu5QJkaRUC4HV6j6lbtZpYax4U"
    
    try
    {
        Set-SnipeitInfo -URL $apiUrl -APIKey $apiKey
        Write-Host "Successfully connected to Snipe-IT API." -ForegroundColor Green
        return $true
    } catch
    {
        Write-Host "Failed to connect to Snipe-IT API. Please check your connection settings." -ForegroundColor Red
        return $false
    }
}

# Main Menu Display Function
function Show-MainMenu {
    Clear-Host
    Write-Host "===========================================" -ForegroundColor Blue
    Write-Host "         SNIPE-IT MANAGEMENT SYSTEM       " -ForegroundColor Blue
    Write-Host "===========================================" -ForegroundColor Blue
    Write-Host "1. 💻 Standard Checkout/Checkin Operations" -ForegroundColor Cyan
    Write-Host "2. 🔧 Temp & Broken Device Management" -ForegroundColor Yellow
    Write-Host "3. 📊 Bulk Status Updates" -ForegroundColor Magenta
    Write-Host "4. 🎫 Submit Repair Ticket" -ForegroundColor Green
    Write-Host "5. 📋 Reports & Statistics" -ForegroundColor White
    Write-Host "6. 🚪 Exit" -ForegroundColor Red
    Write-Host "===========================================" -ForegroundColor Blue
    Write-Host "Enter your choice [1-6]: " -ForegroundColor Yellow -NoNewline
}

# Standard Checkout/Checkin Menu
function Show-CheckoutMenu {
    Clear-Host
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "    STANDARD CHECKOUT/CHECKIN OPERATIONS   " -ForegroundColor Cyan
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "1. Check Out Laptop & Charger"
    Write-Host "2. Check In Laptop & Charger"
    Write-Host "3. Check Out Hotspot"
    Write-Host "4. Check In Hotspot"
    Write-Host "5. Back to Main Menu"
    Write-Host "===========================================" -ForegroundColor Cyan
    Write-Host "Enter your choice [1-5]: " -ForegroundColor Yellow -NoNewline
}

# Process Standard Checkout/Checkin Menu
function Process-CheckoutMenu {
    $continue = $true
    
    while ($continue) {
        Show-CheckoutMenu
        $selection = Read-Host
        
        # Check if functions are available before calling them
        switch ($selection) {
            "1" { 
                if (Get-Command -Name "CheckoutLaptopCharger" -ErrorAction SilentlyContinue) {
                    CheckoutLaptopCharger 
                } else {
                    Write-Host "CheckoutLaptopCharger function not available. Please check snipeit_check.psm1 module." -ForegroundColor Red
                    Start-Sleep -Seconds 3
                }
            }
            "2" { 
                if (Get-Command -Name "CheckinLaptopCharger" -ErrorAction SilentlyContinue) {
                    CheckinLaptopCharger 
                } else {
                    Write-Host "CheckinLaptopCharger function not available. Please check snipeit_check.psm1 module." -ForegroundColor Red
                    Start-Sleep -Seconds 3
                }
            }
            "3" { 
                if (Get-Command -Name "CheckoutHotspot" -ErrorAction SilentlyContinue) {
                    CheckoutHotspot 
                } else {
                    Write-Host "CheckoutHotspot function not available. Please check snipeit_check.psm1 module." -ForegroundColor Red
                    Start-Sleep -Seconds 3
                }
            }
            "4" { 
                if (Get-Command -Name "CheckinHotspot" -ErrorAction SilentlyContinue) {
                    CheckinHotspot 
                } else {
                    Write-Host "CheckinHotspot function not available. Please check snipeit_check.psm1 module." -ForegroundColor Red
                    Start-Sleep -Seconds 3
                }
            }
            "5" { $continue = $false }
            default { 
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep -Seconds 2 
            }
        }
    }
}

# Bulk Status Update Function
function Process-BulkStatusUpdate {
    Clear-Host
    Write-Host "===========================================" -ForegroundColor Magenta
    Write-Host "         BULK STATUS UPDATE               " -ForegroundColor Magenta
    Write-Host "===========================================" -ForegroundColor Magenta
    
    Write-Host "Status Update Options:" -ForegroundColor White
    Write-Host "1 = Ready to Deploy" -ForegroundColor Green
    Write-Host "2 = Deployed" -ForegroundColor Cyan
    Write-Host "3 = Undeployable" -ForegroundColor Yellow
    Write-Host "4 = Pending" -ForegroundColor Blue
    Write-Host "5 = Out for Repair" -ForegroundColor Orange
    Write-Host "6 = Broken" -ForegroundColor Red
    Write-Host "7 = Lost/Stolen" -ForegroundColor Magenta
    Write-Host "8 = Archived" -ForegroundColor Gray
    
    $statusId = Read-Host "`nEnter new status ID (1-8)"
    
    if ($statusId -match '^[1-8]$') {
        $assetTags = Read-Host "Enter asset tags separated by commas (e.g., LAPTOP001,LAPTOP002)"
        
        if (-not [string]::IsNullOrWhiteSpace($assetTags)) {
            $tags = $assetTags.Split(',') | ForEach-Object { $_.Trim() }
            
            Write-Host "`nProcessing bulk status update..." -ForegroundColor Yellow
            $successCount = 0
            $failCount = 0
            
            foreach ($tag in $tags) {
                try {
                    $asset = Get-SnipeitAsset -asset_tag $tag
                    if ($asset) {
                        Set-SnipeitAsset -id $asset.id -status_id $statusId
                        Write-Host "✓ Updated $tag" -ForegroundColor Green
                        $successCount++
                    } else {
                        Write-Host "✗ Asset $tag not found" -ForegroundColor Red
                        $failCount++
                    }
                } catch {
                    Write-Host "✗ Failed to update $tag : $($_.Exception.Message)" -ForegroundColor Red
                    $failCount++
                }
            }
            
            Write-Host "`nBulk update completed:" -ForegroundColor White
            Write-Host "✓ Success: $successCount" -ForegroundColor Green
            Write-Host "✗ Failed: $failCount" -ForegroundColor Red
        } else {
            Write-Host "No asset tags provided." -ForegroundColor Red
        }
    } else {
        Write-Host "Invalid status ID. Please enter a number between 1-8." -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}

# Submit Repair Ticket Function
function Submit-RepairTicket {
    Clear-Host
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host "         SUBMIT REPAIR TICKET             " -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    
    $assetTag = Read-Host "`nEnter asset tag for repair"
    
    if (-not [string]::IsNullOrWhiteSpace($assetTag)) {
        try {
            $asset = Get-SnipeitAsset -asset_tag $assetTag
            if ($asset) {
                Write-Host "Asset: $($asset.name) - $($asset.model.name)" -ForegroundColor White
                
                $issue = Read-Host "Describe the issue"
                $priority = Read-Host "Priority (Low/Medium/High)"
                $supplierIdInput = Read-Host "Supplier ID (or press Enter for default: 1)"
                $supplierId = if ([string]::IsNullOrWhiteSpace($supplierIdInput)) { 1 } else { [int]$supplierIdInput }
                
                # Create maintenance record
                $maintenanceParams = @{
                    asset_id = $asset.id
                    supplier_id = $supplierId
                    asset_maintenance_type = "Repair"
                    title = "Repair Ticket - $priority Priority"
                    start_date = Get-Date
                    notes = "Priority: $priority`nIssue Description: $issue`nReported: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
                }
                
                $result = New-SnipeitAssetMaintenance @maintenanceParams
                
                if ($result) {
                    # Update asset status to "Out for Repair" (status ID 5)
                    Set-SnipeitAsset -id $asset.id -status_id 5
                    
                    Write-Host "`n✓ Repair ticket submitted successfully!" -ForegroundColor Green
                    Write-Host "✓ Asset status updated to 'Out for Repair'" -ForegroundColor Green
                    Write-Host "Ticket Details:" -ForegroundColor White
                    Write-Host "- Asset: $assetTag" -ForegroundColor Gray
                    Write-Host "- Priority: $priority" -ForegroundColor Gray
                    Write-Host "- Issue: $issue" -ForegroundColor Gray
                } else {
                    Write-Host "✗ Failed to submit repair ticket" -ForegroundColor Red
                }
            } else {
                Write-Host "Asset not found: $assetTag" -ForegroundColor Red
            }
        } catch {
            Write-Host "Error submitting repair ticket: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "Asset tag cannot be empty." -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}

# Reports and Statistics Function
function Show-ReportsMenu {
    Clear-Host
    Write-Host "===========================================" -ForegroundColor White
    Write-Host "         REPORTS & STATISTICS             " -ForegroundColor White
    Write-Host "===========================================" -ForegroundColor White
    
    Write-Host "1. Asset Status Summary"
    Write-Host "2. Today's Checkouts"
    Write-Host "3. Overdue Returns"
    Write-Host "4. Maintenance Records"
    Write-Host "5. Back to Main Menu"
    Write-Host "===========================================" -ForegroundColor White
    Write-Host "Enter your choice [1-5]: " -ForegroundColor Yellow -NoNewline
    
    $selection = Read-Host
    
    switch ($selection) {
        "1" { Show-AssetStatusSummary }
        "2" { Show-TodaysCheckouts }
        "3" { Show-OverdueReturns }
        "4" { Show-MaintenanceRecords }
        "5" { return }
        default { 
            Write-Host "Invalid selection" -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-ReportsMenu
        }
    }
}

# Asset Status Summary
function Show-AssetStatusSummary {
    Write-Host "`nGenerating asset status summary..." -ForegroundColor Yellow
    
    try {
        $allAssets = Get-SnipeitAsset
        $statusCounts = @{}
        
        foreach ($asset in $allAssets) {
            $status = $asset.status_label.name
            if ($statusCounts.ContainsKey($status)) {
                $statusCounts[$status]++
            } else {
                $statusCounts[$status] = 1
            }
        }
        
        Write-Host "`nAsset Status Summary:" -ForegroundColor White
        Write-Host "=" * 50 -ForegroundColor Gray
        foreach ($status in $statusCounts.Keys | Sort-Object) {
            Write-Host "$status : $($statusCounts[$status])" -ForegroundColor Cyan
        }
        Write-Host "=" * 50 -ForegroundColor Gray
        Write-Host "Total Assets: $($allAssets.Count)" -ForegroundColor Green
    } catch {
        Write-Host "Error generating status summary: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

# Today's Checkouts
function Show-TodaysCheckouts {
    Write-Host "`nRetrieving today's checkouts..." -ForegroundColor Yellow
    
    try {
        $today = Get-Date -Format "yyyy-MM-dd"
        $allAssets = Get-SnipeitAsset -status "Deployed"
        $todaysCheckouts = @()
        
        foreach ($asset in $allAssets) {
            if ($asset.last_checkout -like "$today*") {
                $todaysCheckouts += $asset
            }
        }
        
        Write-Host "`nToday's Checkouts ($today):" -ForegroundColor White
        Write-Host "=" * 80 -ForegroundColor Gray
        
        if ($todaysCheckouts.Count -gt 0) {
            foreach ($checkout in $todaysCheckouts) {
                Write-Host "Asset: $($checkout.asset_tag) | User: $($checkout.assigned_to.name) | Time: $($checkout.last_checkout)" -ForegroundColor Cyan
            }
            Write-Host "=" * 80 -ForegroundColor Gray
            Write-Host "Total checkouts today: $($todaysCheckouts.Count)" -ForegroundColor Green
        } else {
            Write-Host "No checkouts found for today." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error retrieving today's checkouts: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

# Overdue Returns
function Show-OverdueReturns {
    Write-Host "`nChecking for overdue returns..." -ForegroundColor Yellow
    
    try {
        $today = Get-Date
        $allAssets = Get-SnipeitAsset -status "Deployed"
        $overdueAssets = @()
        
        foreach ($asset in $allAssets) {
            if ($asset.expected_checkin) {
                $expectedDate = [DateTime]::Parse($asset.expected_checkin)
                if ($expectedDate -lt $today) {
                    $overdueAssets += $asset
                }
            }
        }
        
        Write-Host "`nOverdue Returns:" -ForegroundColor White
        Write-Host "=" * 100 -ForegroundColor Gray
        
        if ($overdueAssets.Count -gt 0) {
            foreach ($overdue in $overdueAssets) {
                $daysOverdue = ([DateTime]::Now - [DateTime]::Parse($overdue.expected_checkin)).Days
                Write-Host "Asset: $($overdue.asset_tag) | User: $($overdue.assigned_to.name) | Expected: $($overdue.expected_checkin) | Days Overdue: $daysOverdue" -ForegroundColor Red
            }
            Write-Host "=" * 100 -ForegroundColor Gray
            Write-Host "Total overdue assets: $($overdueAssets.Count)" -ForegroundColor Red
        } else {
            Write-Host "No overdue returns found." -ForegroundColor Green
        }
    } catch {
        Write-Host "Error checking overdue returns: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

# Maintenance Records
function Show-MaintenanceRecords {
    Write-Host "`nRetrieving maintenance records..." -ForegroundColor Yellow
    
    try {
        $maintenanceRecords = Get-SnipeitAssetMaintenance
        
        Write-Host "`nRecent Maintenance Records:" -ForegroundColor White
        Write-Host "=" * 120 -ForegroundColor Gray
        
        if ($maintenanceRecords -and $maintenanceRecords.Count -gt 0) {
            $recentRecords = $maintenanceRecords | Sort-Object start_date -Descending | Select-Object -First 10
            
            foreach ($record in $recentRecords) {
                Write-Host "Asset: $($record.asset.asset_tag) | Type: $($record.asset_maintenance_type) | Start: $($record.start_date) | Title: $($record.title)" -ForegroundColor Cyan
            }
            Write-Host "=" * 120 -ForegroundColor Gray
            Write-Host "Showing 10 most recent records. Total records: $($maintenanceRecords.Count)" -ForegroundColor Green
        } else {
            Write-Host "No maintenance records found." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error retrieving maintenance records: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

# Main Application Loop
function Start-MainApplication {
    $continue = $true
    
    while ($continue) {
        Show-MainMenu
        $selection = Read-Host
        
        switch ($selection) {
            "1" { Process-CheckoutMenu }
            "2" { 
                if (Get-Command -Name "Process-TempBrokenMenu" -ErrorAction SilentlyContinue) {
                    Process-TempBrokenMenu 
                } else {
                    Write-Host "Process-TempBrokenMenu function not available. Please check snipeit_temp_broken.psm1 module." -ForegroundColor Red
                    Write-Host "Available functions from imported modules:" -ForegroundColor Yellow
                    Get-Command -Module snipeit_temp_broken -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Cyan }
                    Start-Sleep -Seconds 5
                }
            }
            "3" { Process-BulkStatusUpdate }
            "4" { Submit-RepairTicket }
            "5" { Show-ReportsMenu }
            "6" { 
                Write-Host "Exiting Snipe-IT Management System..." -ForegroundColor Yellow
                $continue = $false 
            }
            default { 
                Write-Host "Invalid selection. Please enter a number between 1-6." -ForegroundColor Red
                Start-Sleep -Seconds 2 
            }
        }
    }
}

# Main script execution starts here
Write-Host "===========================================" -ForegroundColor Blue
Write-Host "         SNIPE-IT MANAGEMENT SYSTEM       " -ForegroundColor Blue
Write-Host "===========================================" -ForegroundColor Blue

# Initialize the Snipe-IT connection and modules
Write-Host "Initializing Snipe-IT system..." -ForegroundColor Yellow
$initResult = Initialize-SnipeIT

if ($initResult)
{
    Write-Host "`nInitialization successful! Starting main application..." -ForegroundColor Green
    Start-Sleep -Seconds 2
    
    # Start the main application
    Start-MainApplication
} else
{
    Write-Host "`nInitialization failed. Please check the errors above and try again." -ForegroundColor Red
    Write-Host "Make sure the modules directory exists with the required .psm1 files." -ForegroundColor Yellow
}

Write-Host "`nThank you for using Snipe-IT Management System!" -ForegroundColor Green
Write-Host "Press any key to exit..." -ForegroundColor Gray
[void][System.Console]::ReadKey($true)
