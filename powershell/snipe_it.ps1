# Function to handle secure credentials storage and retrieval
function Initialize-SnipeITCredentials
{
    # Define the path for storing encrypted credentials
    $credentialPath = Join-Path $PSScriptRoot "secure\snipeit_credentials.xml"
    $credentialDir = Split-Path $credentialPath -Parent

    # Create the secure directory if it doesn't exist
    if (-not (Test-Path $credentialDir))
    {
        try
        {
            New-Item -Path $credentialDir -ItemType Directory -Force | Out-Null
            Write-Host "Created secure directory for credentials." -ForegroundColor Green
        } catch
        {
            Write-Host "Error creating secure directory: $($_.Exception.Message)" -ForegroundColor Red
            return $null, $null
        }
    }

    # Check if credentials file exists
    if (Test-Path $credentialPath)
    {
        try
        {
            # Import existing credentials
            $credentialObject = Import-Clixml -Path $credentialPath
            $apiUrl = $credentialObject.ApiUrl
            $apiKeySecure = $credentialObject.ApiKey | ConvertTo-SecureString
            $apiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKeySecure))
            
            Write-Host "Loaded existing Snipe-IT credentials." -ForegroundColor Green
            return $apiUrl, $apiKey
        } catch
        {
            Write-Host "Error loading credentials: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Will prompt for new credentials." -ForegroundColor Yellow
        }
    }

    # If we get here, we need to prompt for credentials
    Write-Host "No saved Snipe-IT credentials found. Please enter your API details:" -ForegroundColor Yellow
    $apiUrl = Read-Host "Enter Snipe-IT URL (e.g., https://inventory.company.com)"
    $apiKeySecure = Read-Host "Enter your Snipe-IT API Key" -AsSecureString
    
    # Convert SecureString to plain text for immediate use
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKeySecure)
    $apiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    
    # Save the credentials
    try
    {
        $credentialObject = New-Object PSObject -Property @{
            ApiUrl = $apiUrl
            ApiKey = $apiKeySecure | ConvertFrom-SecureString
        }
        
        $credentialObject | Export-Clixml -Path $credentialPath
        Write-Host "Credentials saved securely." -ForegroundColor Green
    } catch
    {
        Write-Host "Failed to save credentials: $($_.Exception.Message)" -ForegroundColor Red
    }

    return $apiUrl, $apiKey
}

# Function to initialize the comprehensive reporting module
function Initialize-ComprehensiveReporting
{
    if ($global:ReportingModuleAvailable)
    {
        try
        {
            # Get credentials from secure storage
            $apiUrl, $apiKey = Initialize-SnipeITCredentials
            
            if ([string]::IsNullOrEmpty($apiUrl) -or [string]::IsNullOrEmpty($apiKey))
            {
                Write-Host "Failed to get secure credentials for reporting module." -ForegroundColor Red
                return $false
            }
            
            Initialize-SnipeITConnection -BaseUrl $apiUrl -ApiToken $apiKey
            Write-Host "Comprehensive reporting module initialized successfully!" -ForegroundColor Green
            return $true
        } catch
        {
            Write-Host "Failed to initialize comprehensive reporting: $($Error[0])" -ForegroundColor Red
            return $false
        }
    }
    return $false
}


# Main Menu Display Function
function Show-MainMenu
{
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
function Show-CheckoutMenu
{
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
function Process-CheckoutMenu
{
    $continue = $true
    
    while ($continue)
    {
        Show-CheckoutMenu
        $selection = Read-Host
        
        # Check if functions are available before calling them
        switch ($selection)
        {
            "1"
            { 
                if (Get-Command -Name "CheckoutLaptopCharger" -ErrorAction SilentlyContinue)
                {
                    CheckoutLaptopCharger 
                } else
                {
                    Write-Host "CheckoutLaptopCharger function not available. Please check snipeit_check.psm1 module." -ForegroundColor Red
                    Start-Sleep -Seconds 3
                }
            }
            "2"
            { 
                if (Get-Command -Name "CheckinLaptopCharger" -ErrorAction SilentlyContinue)
                {
                    CheckinLaptopCharger 
                } else
                {
                    Write-Host "CheckinLaptopCharger function not available. Please check snipeit_check.psm1 module." -ForegroundColor Red
                    Start-Sleep -Seconds 3
                }
            }
            "3"
            { 
                if (Get-Command -Name "CheckoutHotspot" -ErrorAction SilentlyContinue)
                {
                    CheckoutHotspot 
                } else
                {
                    Write-Host "CheckoutHotspot function not available. Please check snipeit_check.psm1 module." -ForegroundColor Red
                    Start-Sleep -Seconds 3
                }
            }
            "4"
            { 
                if (Get-Command -Name "CheckinHotspot" -ErrorAction SilentlyContinue)
                {
                    CheckinHotspot 
                } else
                {
                    Write-Host "CheckinHotspot function not available. Please check snipeit_check.psm1 module." -ForegroundColor Red
                    Start-Sleep -Seconds 3
                }
            }
            "5"
            { $continue = $false 
            }
            default
            { 
                Write-Host "Invalid selection" -ForegroundColor Red
                Start-Sleep -Seconds 2 
            }
        }
    }
}

# Bulk Status Update Function
function Process-BulkStatusUpdate
{
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
    
    if ($statusId -match '^[1-8]$')
    {
        $assetTags = Read-Host "Enter asset tags separated by commas (e.g., LAPTOP001,LAPTOP002)"
        
        if (-not [string]::IsNullOrWhiteSpace($assetTags))
        {
            $tags = $assetTags.Split(',') | ForEach-Object { $_.Trim() }
            
            Write-Host "`nProcessing bulk status update..." -ForegroundColor Yellow
            $successCount = 0
            $failCount = 0
            
            foreach ($tag in $tags)
            {
                try
                {
                    $asset = Get-SnipeitAsset -asset_tag $tag
                    if ($asset)
                    {
                        Set-SnipeitAsset -id $asset.id -status_id $statusId
                        Write-Host "✓ Updated $tag" -ForegroundColor Green
                        $successCount++
                    } else
                    {
                        Write-Host "✗ Asset $tag not found" -ForegroundColor Red
                        $failCount++
                    }
                } catch
                {
                    Write-Host "✗ Failed to update $tag : $($_.Exception.Message)" -ForegroundColor Red
                    $failCount++
                }
            }
            
            Write-Host "`nBulk update completed:" -ForegroundColor White
            Write-Host "✓ Success: $successCount" -ForegroundColor Green
            Write-Host "✗ Failed: $failCount" -ForegroundColor Red
        } else
        {
            Write-Host "No asset tags provided." -ForegroundColor Red
        }
    } else
    {
        Write-Host "Invalid status ID. Please enter a number between 1-8." -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}

# Submit Repair Ticket Function
function Submit-RepairTicket
{
    Clear-Host
    Write-Host "===========================================" -ForegroundColor Green
    Write-Host "         SUBMIT REPAIR TICKET             " -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    
    $assetTag = Read-Host "`nEnter asset tag for repair"
    
    if (-not [string]::IsNullOrWhiteSpace($assetTag))
    {
        try
        {
            $asset = Get-SnipeitAsset -asset_tag $assetTag
            if ($asset)
            {
                Write-Host "Asset: $($asset.name) - $($asset.model.name)" -ForegroundColor White
                
                $issue = Read-Host "Describe the issue"
                $priority = Read-Host "Priority (Low/Medium/High)"
                $supplierIdInput = Read-Host "Supplier ID (or press Enter for default: 1)"
                $supplierId = if ([string]::IsNullOrWhiteSpace($supplierIdInput))
                { 1 
                } else
                { [int]$supplierIdInput 
                }
                
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
                
                if ($result)
                {
                    # Update asset status to "Out for Repair" (status ID 5)
                    Set-SnipeitAsset -id $asset.id -status_id 5
                    
                    Write-Host "`n✓ Repair ticket submitted successfully!" -ForegroundColor Green
                    Write-Host "✓ Asset status updated to 'Out for Repair'" -ForegroundColor Green
                    Write-Host "Ticket Details:" -ForegroundColor White
                    Write-Host "- Asset: $assetTag" -ForegroundColor Gray
                    Write-Host "- Priority: $priority" -ForegroundColor Gray
                    Write-Host "- Issue: $issue" -ForegroundColor Gray
                } else
                {
                    Write-Host "✗ Failed to submit repair ticket" -ForegroundColor Red
                }
            } else
            {
                Write-Host "Asset not found: $assetTag" -ForegroundColor Red
            }
        } catch
        {
            Write-Host "Error submitting repair ticket: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else
    {
        Write-Host "Asset tag cannot be empty." -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
}

# Enhanced Reports and Statistics Function
function Show-ReportsMenu
{
    Clear-Host
    Write-Host "===========================================" -ForegroundColor White
    Write-Host "         REPORTS & STATISTICS             " -ForegroundColor White
    Write-Host "===========================================" -ForegroundColor White
    
    Write-Host "📊 BASIC REPORTS:" -ForegroundColor Cyan
    Write-Host "1. Asset Status Summary"
    Write-Host "2. Today's Checkouts"
    Write-Host "3. Overdue Returns"
    Write-Host "4. Maintenance Records"
    Write-Host ""
    
    if ($global:ReportingModuleAvailable)
    {
        Write-Host "📈 COMPREHENSIVE REPORTING:" -ForegroundColor Green
        Write-Host "5. Advanced Asset Reports Dashboard"
        Write-Host "6. Quick Asset Overview"
        Write-Host "7. Warranty Status Report"
        Write-Host "8. User Activity Report"
        Write-Host "9. Location Distribution Report"
        Write-Host ""
    }
    
    Write-Host "0. Back to Main Menu" -ForegroundColor Gray
    Write-Host "===========================================" -ForegroundColor White
    Write-Host "Enter your choice: " -ForegroundColor Yellow -NoNewline
    
    $selection = Read-Host
    
    switch ($selection)
    {
        "1"
        { Show-AssetStatusSummary 
        }
        "2"
        { Show-TodaysCheckouts 
        }
        "3"
        { Show-OverdueReturns 
        }
        "4"
        { Show-MaintenanceRecords 
        }
        "5"
        { 
            if ($global:ReportingModuleAvailable)
            {
                Launch-ComprehensiveReporting
            } else
            {
                Write-Host "Comprehensive reporting not available." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Show-ReportsMenu
            }
        }
        "6"
        { 
            if ($global:ReportingModuleAvailable)
            {
                Run-QuickAssetOverview
            } else
            {
                Write-Host "Advanced reporting not available." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Show-ReportsMenu
            }
        }
        "7"
        { 
            if ($global:ReportingModuleAvailable)
            {
                Run-WarrantyReport
            } else
            {
                Write-Host "Advanced reporting not available." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Show-ReportsMenu
            }
        }
        "8"
        { 
            if ($global:ReportingModuleAvailable)
            {
                Run-UserActivityReport
            } else
            {
                Write-Host "Advanced reporting not available." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Show-ReportsMenu
            }
        }
        "9"
        { 
            if ($global:ReportingModuleAvailable)
            {
                Run-LocationReport
            } else
            {
                Write-Host "Advanced reporting not available." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Show-ReportsMenu
            }
        }
        "0"
        { return 
        }
        default
        { 
            Write-Host "Invalid selection" -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-ReportsMenu
        }
    }
}

# Comprehensive reporting integration functions

function Launch-ComprehensiveReporting
{
    Clear-Host
    Write-Host "Launching Comprehensive Reporting Dashboard..." -ForegroundColor Yellow
    
    if (Initialize-ComprehensiveReporting)
    {
        try
        {
            Show-AssetManagementMenu
        } catch
        {
            Write-Host "Error launching comprehensive reporting: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else
    {
        Write-Host "Failed to initialize comprehensive reporting system." -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

function Run-QuickAssetOverview
{
    Clear-Host
    Write-Host "Generating Quick Asset Overview..." -ForegroundColor Yellow
    
    if (Initialize-ComprehensiveReporting)
    {
        try
        {
            $exportChoice = Read-Host "Export to CSV? (y/n)"
            if ($exportChoice -eq 'y' -or $exportChoice -eq 'Y')
            {
                $exportPath = Read-Host "Enter export path (or press Enter for default)"
                if ([string]::IsNullOrEmpty($exportPath))
                {
                    $exportPath = ".\Reports\AssetOverview_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
                }
                Get-AssetOverview -ExportPath $exportPath
            } else
            {
                Get-AssetOverview
            }
        } catch
        {
            Write-Host "Error generating asset overview: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else
    {
        Write-Host "Failed to initialize reporting system." -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

function Run-WarrantyReport
{
    Clear-Host
    Write-Host "Generating Warranty Status Report..." -ForegroundColor Yellow
    
    if (Initialize-ComprehensiveReporting)
    {
        try
        {
            $daysAhead = Read-Host "Days ahead to check for expiring warranties (default: 90)"
            if ([string]::IsNullOrEmpty($daysAhead))
            { $daysAhead = 90 
            }
            
            $exportChoice = Read-Host "Export to CSV? (y/n)"
            if ($exportChoice -eq 'y' -or $exportChoice -eq 'Y')
            {
                $exportPath = Read-Host "Enter export path (or press Enter for default)"
                if ([string]::IsNullOrEmpty($exportPath))
                {
                    $exportPath = ".\Reports\WarrantyStatus_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
                }
                Get-WarrantyStatus -DaysAhead $daysAhead -ExportPath $exportPath
            } else
            {
                Get-WarrantyStatus -DaysAhead $daysAhead
            }
        } catch
        {
            Write-Host "Error generating warranty report: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else
    {
        Write-Host "Failed to initialize reporting system." -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

function Run-UserActivityReport
{
    Clear-Host
    Write-Host "Generating User Activity Report..." -ForegroundColor Yellow
    
    if (Initialize-ComprehensiveReporting)
    {
        try
        {
            $exportChoice = Read-Host "Export to CSV? (y/n)"
            if ($exportChoice -eq 'y' -or $exportChoice -eq 'Y')
            {
                $exportPath = Read-Host "Enter export path (or press Enter for default)"
                if ([string]::IsNullOrEmpty($exportPath))
                {
                    $exportPath = ".\Reports\UserActivity_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
                }
                Get-UserActivity -ExportPath $exportPath
            } else
            {
                Get-UserActivity
            }
        } catch
        {
            Write-Host "Error generating user activity report: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else
    {
        Write-Host "Failed to initialize reporting system." -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

function Run-LocationReport
{
    Clear-Host
    Write-Host "Generating Location Distribution Report..." -ForegroundColor Yellow
    
    if (Initialize-ComprehensiveReporting)
    {
        try
        {
            $exportChoice = Read-Host "Export to CSV? (y/n)"
            if ($exportChoice -eq 'y' -or $exportChoice -eq 'Y')
            {
                $exportPath = Read-Host "Enter export path (or press Enter for default)"
                if ([string]::IsNullOrEmpty($exportPath))
                {
                    $exportPath = ".\Reports\LocationDistribution_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
                }
                Get-LocationDistribution -ExportPath $exportPath
            } else
            {
                Get-LocationDistribution
            }
        } catch
        {
            Write-Host "Error generating location report: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else
    {
        Write-Host "Failed to initialize reporting system." -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

# Basic reporting functions

# Asset Status Summary
function Show-AssetStatusSummary
{
    Write-Host "`nGenerating asset status summary..." -ForegroundColor Yellow
    
    try
    {
        $allAssets = Get-SnipeitAsset
        $statusCounts = @{}
        
        foreach ($asset in $allAssets)
        {
            $status = $asset.status_label.name
            if ($statusCounts.ContainsKey($status))
            {
                $statusCounts[$status]++
            } else
            {
                $statusCounts[$status] = 1
            }
        }
        
        Write-Host "`nAsset Status Summary:" -ForegroundColor White
        Write-Host "=" * 50 -ForegroundColor Gray
        foreach ($status in $statusCounts.Keys | Sort-Object)
        {
            Write-Host "$status : $($statusCounts[$status])" -ForegroundColor Cyan
        }
        Write-Host "=" * 50 -ForegroundColor Gray
        Write-Host "Total Assets: $($allAssets.Count)" -ForegroundColor Green
    } catch
    {
        Write-Host "Error generating status summary: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

# Today's Checkouts
function Show-TodaysCheckouts
{
    Write-Host "`nRetrieving today's checkouts..." -ForegroundColor Yellow
    
    try
    {
        $today = Get-Date -Format "yyyy-MM-dd"
        $allAssets = Get-SnipeitAsset -status "Deployed"
        $todaysCheckouts = @()
        
        foreach ($asset in $allAssets)
        {
            if ($asset.last_checkout -like "$today*")
            {
                $todaysCheckouts += $asset
            }
        }
        
        Write-Host "`nToday's Checkouts ($today):" -ForegroundColor White
        Write-Host "=" * 80 -ForegroundColor Gray
        
        if ($todaysCheckouts.Count -gt 0)
        {
            foreach ($checkout in $todaysCheckouts)
            {
                Write-Host "Asset: $($checkout.asset_tag) | User: $($checkout.assigned_to.name) | Time: $($checkout.last_checkout)" -ForegroundColor Cyan
            }
            Write-Host "=" * 80 -ForegroundColor Gray
            Write-Host "Total checkouts today: $($todaysCheckouts.Count)" -ForegroundColor Green
        } else
        {
            Write-Host "No checkouts found for today." -ForegroundColor Yellow
        }
    } catch
    {
        Write-Host "Error retrieving today's checkouts: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

# Overdue Returns
function Show-OverdueReturns
{
    Write-Host "`nChecking for overdue returns..." -ForegroundColor Yellow
    
    try
    {
        $today = Get-Date
        $allAssets = Get-SnipeitAsset -status "Deployed"
        $overdueAssets = @()
        
        foreach ($asset in $allAssets)
        {
            if ($asset.expected_checkin)
            {
                $expectedDate = [DateTime]::Parse($asset.expected_checkin)
                if ($expectedDate -lt $today)
                {
                    $overdueAssets += $asset
                }
            }
        }
        
        Write-Host "`nOverdue Returns:" -ForegroundColor White
        Write-Host "=" * 100 -ForegroundColor Gray
        
        if ($overdueAssets.Count -gt 0)
        {
            foreach ($overdue in $overdueAssets)
            {
                $daysOverdue = ([DateTime]::Now - [DateTime]::Parse($overdue.expected_checkin)).Days
                Write-Host "Asset: $($overdue.asset_tag) | User: $($overdue.assigned_to.name) | Expected: $($overdue.expected_checkin) | Days Overdue: $daysOverdue" -ForegroundColor Red
            }
            Write-Host "=" * 100 -ForegroundColor Gray
            Write-Host "Total overdue assets: $($overdueAssets.Count)" -ForegroundColor Red
        } else
        {
            Write-Host "No overdue returns found." -ForegroundColor Green
        }
    } catch
    {
        Write-Host "Error checking overdue returns: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

# Maintenance Records
function Show-MaintenanceRecords
{
    Write-Host "`nRetrieving maintenance records..." -ForegroundColor Yellow
    
    try
    {
        $maintenanceRecords = Get-SnipeitAssetMaintenance
        
        Write-Host "`nRecent Maintenance Records:" -ForegroundColor White
        Write-Host "=" * 120 -ForegroundColor Gray
        
        if ($maintenanceRecords -and $maintenanceRecords.Count -gt 0)
        {
            $recentRecords = $maintenanceRecords | Sort-Object start_date -Descending | Select-Object -First 10
            
            foreach ($record in $recentRecords)
            {
                Write-Host "Asset: $($record.asset.asset_tag) | Type: $($record.asset_maintenance_type) | Start: $($record.start_date) | Title: $($record.title)" -ForegroundColor Cyan
            }
            Write-Host "=" * 120 -ForegroundColor Gray
            Write-Host "Showing 10 most recent records. Total records: $($maintenanceRecords.Count)" -ForegroundColor Green
        } else
        {
            Write-Host "No maintenance records found." -ForegroundColor Yellow
        }
    } catch
    {
        Write-Host "Error retrieving maintenance records: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to reports menu..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    Show-ReportsMenu
}

# Main Application Loop
function Start-MainApplication
{
    $continue = $true
    
    while ($continue)
    {
        Show-MainMenu
        $selection = Read-Host
        
        switch ($selection)
        {
            "1"
            { Process-CheckoutMenu 
            }
            "2"
            { 
                if (Get-Command -Name "Process-TempBrokenMenu" -ErrorAction SilentlyContinue)
                {
                    Process-TempBrokenMenu 
                } else
                {
                    Write-Host "Process-TempBrokenMenu function not available. Please check snipeit_temp_broken.psm1 module." -ForegroundColor Red
                    Write-Host "Available functions from imported modules:" -ForegroundColor Yellow
                    Get-Command -Module snipeit_temp_broken -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Cyan }
                    Start-Sleep -Seconds 5
                }
            }
            "3"
            { Process-BulkStatusUpdate 
            }
            "4"
            { Submit-RepairTicket 
            }
            "5"
            { Show-ReportsMenu 
            }
            "6"
            { 
                Write-Host "Exiting Snipe-IT Management System..." -ForegroundColor Yellow
                $continue = $false 
            }
            default
            { 
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
$initResult = Initialize-SnipeITCredentials

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
