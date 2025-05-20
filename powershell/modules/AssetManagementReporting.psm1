#Requires -Version 5.1

<#
.SYNOPSIS
    PowerShell module for Snipe-IT and SNITEPS asset management reporting and statistics.

.DESCRIPTION
    This module provides comprehensive reporting capabilities for IT asset management systems.
    It includes functions for connecting to Snipe-IT and SNITEPS APIs and generating detailed
    reports across 10 key areas of asset management.

.AUTHOR
    Asset Management Team

.VERSION
    1.0.0

.NOTES
    Requires PowerShell 5.1 or higher
    Requires valid API credentials for Snipe-IT and SNITEPS
#>

# Module Variables
$Script:SnipeITBaseUrl = $null
$Script:SnipeITHeaders = $null
$Script:SNITEPSBaseUrl = $null
$Script:SNITEPSHeaders = $null
$Script:DefaultTimeout = 30
$Script:RateLimitDelay = 100 # milliseconds

#region Helper Functions

<#
.SYNOPSIS
    Initializes connection to Snipe-IT API.

.DESCRIPTION
    Sets up the base URL and authentication headers for Snipe-IT API calls.

.PARAMETER BaseUrl
    The base URL of your Snipe-IT instance (e.g., "https://your-domain.snipe-it.io")

.PARAMETER ApiToken
    Your Snipe-IT API token

.EXAMPLE
    Initialize-SnipeITConnection -BaseUrl "https://company.snipe-it.io" -ApiToken "your-api-token"
#>
function Initialize-SnipeITConnection
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$ApiToken
    )
    
    $Script:SnipeITBaseUrl = $BaseUrl.TrimEnd('/')
    $Script:SnipeITHeaders = @{
        'Authorization' = "Bearer $ApiToken"
        'Accept' = 'application/json'
        'Content-Type' = 'application/json'
    }
    
    Write-Verbose "Snipe-IT connection initialized for $BaseUrl"
}

<#
.SYNOPSIS
    Initializes connection to SNITEPS API.

.DESCRIPTION
    Sets up the base URL and authentication headers for SNITEPS API calls.

.PARAMETER BaseUrl
    The base URL of your SNITEPS instance

.PARAMETER Username
    SNITEPS username

.PARAMETER Password
    SNITEPS password

.EXAMPLE
    Initialize-SNITEPSConnection -BaseUrl "https://sniteps.company.com" -Username "admin" -Password "password"
#>
function Initialize-SNITEPSConnection
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$Username,
        
        [Parameter(Mandatory = $true)]
        [string]$Password
    )
    
    $Script:SNITEPSBaseUrl = $BaseUrl.TrimEnd('/')
    
    # Create credentials for SNITEPS (assuming basic auth)
    $credentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${Username}:${Password}"))
    $Script:SNITEPSHeaders = @{
        'Authorization' = "Basic $credentials"
        'Accept' = 'application/json'
        'Content-Type' = 'application/json'
    }
    
    Write-Verbose "SNITEPS connection initialized for $BaseUrl"
}

<#
.SYNOPSIS
    Makes an API call with error handling and rate limiting.

.DESCRIPTION
    Wrapper function for Invoke-RestMethod with built-in error handling,
    rate limiting, and retry logic.

.PARAMETER Url
    The API endpoint URL

.PARAMETER Headers
    HTTP headers for the request

.PARAMETER Method
    HTTP method (default: GET)

.PARAMETER Body
    Request body (for POST/PUT requests)

.EXAMPLE
    Invoke-ApiCall -Url "https://api.example.com/assets" -Headers $headers
#>
function Invoke-ApiCall
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,
        
        [string]$Method = 'GET',
        
        [object]$Body = $null,
        
        [int]$MaxRetries = 3
    )
    
    $attempt = 0
    
    do
    {
        try
        {
            $attempt++
            
            $params = @{
                Uri = $Url
                Headers = $Headers
                Method = $Method
                TimeoutSec = $Script:DefaultTimeout
            }
            
            if ($Body)
            {
                $params.Body = $Body | ConvertTo-Json -Depth 10
            }
            
            $response = Invoke-RestMethod @params
            
            # Rate limiting
            Start-Sleep -Milliseconds $Script:RateLimitDelay
            
            return $response
        } catch
        {
            if ($_.Exception.Response.StatusCode -eq 429)
            {
                # Rate limited, wait longer
                Write-Warning "Rate limited. Waiting before retry..."
                Start-Sleep -Seconds (2 * $attempt)
            } elseif ($attempt -ge $MaxRetries)
            {
                throw "API call failed after $MaxRetries attempts: $($_.Exception.Message)"
            } else
            {
                Write-Warning "API call failed (attempt $attempt/$MaxRetries): $($_.Exception.Message)"
                Start-Sleep -Seconds $attempt
            }
        }
    } while ($attempt -lt $MaxRetries)
}

<#
.SYNOPSIS
    Exports data to CSV file.

.DESCRIPTION
    Helper function to export report data to CSV with error handling.

.PARAMETER Data
    The data to export

.PARAMETER FilePath
    Path for the output CSV file

.EXAMPLE
    Export-ToCsv -Data $assetData -FilePath "C:\Reports\assets.csv"
#>
function Export-ToCsv
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try
    {
        $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        Write-Host "Data exported to: $FilePath" -ForegroundColor Green
        $errorMessage = $_.Exception.Message
    } catch
    {
        Write-Error "Failed to export data to $FilePath $(errorMessage)"
    }
}

#endregion

#region 1. Asset Overview

<#
.SYNOPSIS
    Retrieves comprehensive asset overview statistics.

.DESCRIPTION
    Generates an overview of all assets including total counts by status,
    new assets added this month, and asset type distribution.

.PARAMETER ExportPath
    Optional path to export results to CSV

.EXAMPLE
    Get-AssetOverview
    Get-AssetOverview -ExportPath "C:\Reports\asset_overview.csv"
#>
function Get-AssetOverview
{
    [CmdletBinding()]
    param(
        [string]$ExportPath
    )
    
    if (-not $Script:SnipeITBaseUrl)
    {
        throw "Snipe-IT connection not initialized. Run Initialize-SnipeITConnection first."
    }
    
    try
    {
        Write-Host "Gathering asset overview..." -ForegroundColor Yellow
        
        # Get asset counts by status
        $allAssets = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/hardware" -Headers $Script:SnipeITHeaders
        
        # Calculate status distribution
        $statusCounts = $allAssets.rows | Group-Object status_label.name | 
            Select-Object @{Name='Status'; Expression={$_.Name}}, 
            @{Name='Count'; Expression={$_.Count}}
        
        # Get assets added this month
        $thisMonth = (Get-Date).ToString("yyyy-MM")
        $newAssetsThisMonth = $allAssets.rows | Where-Object { 
            $_.created_at.date -match "^$thisMonth" 
        } | Measure-Object | Select-Object -ExpandProperty Count
        
        # Get asset types
        $assetTypes = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/categories" -Headers $Script:SnipeITHeaders
        $typeDistribution = $assetTypes.rows | Select-Object name, assets_count
        
        # Create summary object
        $overview = [PSCustomObject]@{
            TotalAssets = $allAssets.total
            ActiveAssets = ($statusCounts | Where-Object Status -eq "Ready to Deploy").Count
            InactiveAssets = ($statusCounts | Where-Object Status -eq "Undeployable").Count
            ArchivedAssets = ($statusCounts | Where-Object Status -eq "Archived").Count
            NewAssetsThisMonth = $newAssetsThisMonth
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Display results
        Write-Host "`n=== ASSET OVERVIEW ===" -ForegroundColor Cyan
        $overview | Format-List
        
        Write-Host "`n=== ASSET TYPE DISTRIBUTION ===" -ForegroundColor Cyan
        $typeDistribution | Format-Table -AutoSize
        
        # Export if requested
        if ($ExportPath)
        {
            $exportData = @()
            $exportData += $overview
            $exportData += $typeDistribution
            Export-ToCsv -Data $exportData -FilePath $ExportPath
        }
        
        return @{
            Overview = $overview
            TypeDistribution = $typeDistribution
            StatusCounts = $statusCounts
        }
    } catch
    {
        Write-Error "Failed to retrieve asset overview: $($Error[0])"
    }
}

#endregion

#region 2. Audit Status

<#
.SYNOPSIS
    Retrieves asset audit status and compliance metrics.

.DESCRIPTION
    Analyzes audit completion rates, overdue audits, and compliance statistics
    for the current quarter.

.PARAMETER ExportPath
    Optional path to export results to CSV

.EXAMPLE
    Get-AuditStatus
    Get-AuditStatus -ExportPath "C:\Reports\audit_status.csv"
#>
function Get-AuditStatus
{
    [CmdletBinding()]
    param(
        [string]$ExportPath
    )
    
    if (-not $Script:SnipeITBaseUrl)
    {
        throw "Snipe-IT connection not initialized. Run Initialize-SnipeITConnection first."
    }
    
    try
    {
        Write-Host "Analyzing audit status..." -ForegroundColor Yellow
        
        # Get current quarter dates
        $currentDate = Get-Date
        $quarter = [math]::Ceiling($currentDate.Month / 3)
        $quarterStart = Get-Date -Year $currentDate.Year -Month (($quarter - 1) * 3 + 1) -Day 1
        $quarterEnd = $quarterStart.AddMonths(3).AddDays(-1)
        
        # Get audit data (assuming audit log endpoint exists)
        $audits = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/reports/activity" -Headers $Script:SnipeITHeaders
        
        # Filter audits for current quarter
        $quarterAudits = $audits.rows | Where-Object {
            $auditDate = [DateTime]::Parse($_.created_at.date)
            $auditDate -ge $quarterStart -and $auditDate -le $quarterEnd -and
            $_.action_type -eq "audit"
        }
        
        # Calculate metrics
        $completedAudits = $quarterAudits.Count
        $totalAssets = (Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/hardware" -Headers $Script:SnipeITHeaders).total
        
        # Get overdue audits (assets not audited in last 6 months)
        $sixMonthsAgo = (Get-Date).AddMonths(-6)
        $allAssets = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/hardware?limit=500" -Headers $Script:SnipeITHeaders
        
        $overdueAudits = $allAssets.rows | Where-Object {
            if ($_.last_audit_date)
            {
                [DateTime]::Parse($_.last_audit_date.date) -lt $sixMonthsAgo
            } else
            {
                $true # Never audited
            }
        }
        
        $complianceRate = if ($totalAssets -gt 0)
        { 
            [math]::Round((($totalAssets - $overdueAudits.Count) / $totalAssets) * 100, 2) 
        } else
        { 0 
        }
        
        # Create summary
        $auditSummary = [PSCustomObject]@{
            CompletedAuditsThisQuarter = $completedAudits
            OverdueAudits = $overdueAudits.Count
            ComplianceRate = "$complianceRate%"
            TotalAssets = $totalAssets
            QuarterStart = $quarterStart.ToString("yyyy-MM-dd")
            QuarterEnd = $quarterEnd.ToString("yyyy-MM-dd")
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Display results
        Write-Host "`n=== AUDIT STATUS ===" -ForegroundColor Cyan
        $auditSummary | Format-List
        
        if ($overdueAudits.Count -gt 0)
        {
            Write-Host "`n=== OVERDUE AUDITS (Sample) ===" -ForegroundColor Red
            $overdueAudits | Select-Object -First 10 name, asset_tag, location.name, last_audit_date | 
                Format-Table -AutoSize
        }
        
        # Export if requested
        if ($ExportPath)
        {
            $exportData = @($auditSummary) + $overdueAudits
            Export-ToCsv -Data $exportData -FilePath $ExportPath
        }
        
        return @{
            Summary = $auditSummary
            OverdueAudits = $overdueAudits
        }
    } catch
    {
        Write-Error "Failed to retrieve audit status: $($_.Exception.Message)"
    }
}

#endregion

#region 3. Cart Utilization

<#
.SYNOPSIS
    Analyzes cart utilization and checkout patterns.

.DESCRIPTION
    Retrieves statistics on active carts, most frequent users,
    and average checkout duration from SNITEPS.

.PARAMETER ExportPath
    Optional path to export results to CSV

.EXAMPLE
    Get-CartUtilization
    Get-CartUtilization -ExportPath "C:\Reports\cart_utilization.csv"
#>
function Get-CartUtilization
{
    [CmdletBinding()]
    param(
        [string]$ExportPath
    )
    
    if (-not $Script:SNITEPSBaseUrl)
    {
        throw "SNITEPS connection not initialized. Run Initialize-SNITEPSConnection first."
    }
    
    try
    {
        Write-Host "Analyzing cart utilization..." -ForegroundColor Yellow
        
        # Get active carts from SNITEPS
        $activeCarts = Invoke-ApiCall -Url "$Script:SNITEPSBaseUrl/api/carts/active" -Headers $Script:SNITEPSHeaders
        
        # Get checkout history
        $checkoutHistory = Invoke-ApiCall -Url "$Script:SNITEPSBaseUrl/api/checkouts/history" -Headers $Script:SNITEPSHeaders
        
        # Calculate top 5 most frequent cart users
        $topUsers = $checkoutHistory.data | Group-Object user_id | 
            Sort-Object Count -Descending | Select-Object -First 5 |
            ForEach-Object {
                [PSCustomObject]@{
                    UserId = $_.Name
                    CheckoutCount = $_.Count
                    UserName = ($_.Group | Select-Object -First 1).user_name
                }
            }
        
        # Calculate average checkout duration
        $checkoutsWithDuration = $checkoutHistory.data | Where-Object { 
            $_.checked_out_at -and $_.checked_in_at 
        } | ForEach-Object {
            $checkoutTime = [DateTime]::Parse($_.checked_out_at)
            $checkinTime = [DateTime]::Parse($_.checked_in_at)
            ($checkinTime - $checkoutTime).TotalHours
        }
        
        $avgDuration = if ($checkoutsWithDuration.Count -gt 0)
        {
            [math]::Round(($checkoutsWithDuration | Measure-Object -Average).Average, 2)
        } else
        { 0 
        }
        
        # Create summary
        $cartSummary = [PSCustomObject]@{
            ActiveCarts = $activeCarts.total
            ItemsCheckedOut = ($activeCarts.data | Measure-Object items_count -Sum).Sum
            AverageCheckoutDuration = "$avgDuration hours"
            TotalCheckouts = $checkoutHistory.total
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Display results
        Write-Host "`n=== CART UTILIZATION ===" -ForegroundColor Cyan
        $cartSummary | Format-List
        
        Write-Host "`n=== TOP 5 CART USERS ===" -ForegroundColor Cyan
        $topUsers | Format-Table -AutoSize
        
        # Export if requested
        if ($ExportPath)
        {
            $exportData = @($cartSummary) + $topUsers
            Export-ToCsv -Data $exportData -FilePath $ExportPath
        }
        
        return @{
            Summary = $cartSummary
            TopUsers = $topUsers
            ActiveCarts = $activeCarts.data
        }
    } catch
    {
        Write-Error "Failed to retrieve cart utilization: $($_.Exception.Message)"
    }
}

#endregion

#region 4. User Activity

<#
.SYNOPSIS
    Analyzes user activity patterns and asset handling.

.DESCRIPTION
    Retrieves statistics on most active users, recent logins,
    and high-volume asset handlers.

.PARAMETER ExportPath
    Optional path to export results to CSV

.EXAMPLE
    Get-UserActivity
    Get-UserActivity -ExportPath "C:\Reports\user_activity.csv"
#>

function Get-UserActivity
{
    [CmdletBinding()]
    param(
        [string]$ExportPath
    )
    
    if (-not $Script:SnipeITBaseUrl)
    {
        throw "Snipe-IT connection not initialized. Run Initialize-SnipeITConnection first."
    }
    
    try
    {
        Write-Host "Analyzing user activity..." -ForegroundColor Yellow
        
        # Get user activity data
        $activityLog = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/reports/activity?limit=1000" -Headers $Script:SnipeITHeaders
        
        # Get most active users (top 10)
        $lastMonth = (Get-Date).AddDays(-30)
        $recentActivity = $activityLog.rows | Where-Object {
            # Add null check and empty string check
            $_.created_at -and 
            $_.created_at.date -and 
            ![string]::IsNullOrEmpty($_.created_at.date) -and
            [DateTime]::Parse($_.created_at.date) -ge $lastMonth
        }
        
        $topActiveUsers = $recentActivity | Group-Object admin.name | 
            Sort-Object Count -Descending | Select-Object -First 10 |
            ForEach-Object {
                [PSCustomObject]@{
                    UserName = $_.Name
                    ActivityCount = $_.Count
                    LastActivity = ($_.Group | Sort-Object created_at.date -Descending | 
                            Select-Object -First 1).created_at.date
                    }
                }
        
        # Get users with most asset checkouts/checkins
        $assetHandlers = $recentActivity | Where-Object { 
            $_.action_type -eq "checkout" -or $_.action_type -eq "checkin" 
        } | Group-Object admin.name | 
        Sort-Object Count -Descending | Select-Object -First 10 |
        ForEach-Object {
            [PSCustomObject]@{
                UserName = $_.Name
                AssetTransactions = $_.Count
                CheckOuts = ($_.Group | Where-Object action_type -eq "checkout").Count
                CheckIns = ($_.Group | Where-Object action_type -eq "checkin").Count
            }
        }
        
        # Get recent logins (from users endpoint)
        $users = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/users?limit=100" -Headers $Script:SnipeITHeaders
        $recentLogins = $users.rows | Where-Object { 
            # Enhanced null checking before parsing date
            $_.last_login -and 
            $_.last_login.date -and 
            ![string]::IsNullOrEmpty($_.last_login.date) -and
            [DateTime]::Parse($_.last_login.date) -ge $lastMonth 
        } | Sort-Object last_login.date -Descending | Select-Object -First 10 |
        Select-Object @{Name='UserName'; Expression={$_.name}}, 
        @{Name='LastLogin'; Expression={$_.last_login.date}},
        @{Name='AssetsCount'; Expression={$_.assets_count}}
        
        # Create summary
        $activitySummary = [PSCustomObject]@{
            TotalUsers = $users.total
            ActiveUsersLastMonth = $topActiveUsers.Count
            TotalActivitiesLastMonth = $recentActivity.Count
            HighVolumeHandlers = $assetHandlers.Count
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Display results
        Write-Host "`n=== USER ACTIVITY SUMMARY ===" -ForegroundColor Cyan
        $activitySummary | Format-List
        
        Write-Host "`n=== TOP 10 MOST ACTIVE USERS ===" -ForegroundColor Cyan
        $topActiveUsers | Format-Table -AutoSize
        
        Write-Host "`n=== HIGH-VOLUME ASSET HANDLERS ===" -ForegroundColor Cyan
        $assetHandlers | Format-Table -AutoSize
        
        # Export if requested
        if ($ExportPath)
        {
            $exportData = @($activitySummary) + $topActiveUsers + $assetHandlers + $recentLogins
            Export-ToCsv -Data $exportData -FilePath $ExportPath
        }
        
        return @{
            Summary = $activitySummary
            TopActiveUsers = $topActiveUsers
            AssetHandlers = $assetHandlers
            RecentLogins = $recentLogins
        }
    } catch
    {
        Write-Error "Failed to retrieve user activity: $($_.Exception.Message)"
    }
}



#endregion

#region 5. Warranty Status

<#
.SYNOPSIS
    Analyzes warranty status across all assets.

.DESCRIPTION
    Retrieves statistics on expiring warranties, expired warranties,
    and identifies high-risk warranty gaps.

.PARAMETER ExportPath
    Optional path to export results to CSV

.PARAMETER DaysAhead
    Number of days ahead to look for expiring warranties (default: 90)

.EXAMPLE
    Get-WarrantyStatus
    Get-WarrantyStatus -DaysAhead 60 -ExportPath "C:\Reports\warranty_status.csv"
#>

# Modified Get-WarrantyStatus function with comprehensive fixes for the "N/A" issue

function Get-WarrantyStatus
{
    [CmdletBinding()]
    param(
        [string]$ExportPath,
        [int]$DaysAhead = 90
    )
    
    if (-not $Script:SnipeITBaseUrl)
    {
        throw "Snipe-IT connection not initialized. Run Initialize-SnipeITConnection first."
    }
    
    try
    {
        Write-Host "Analyzing warranty status..." -ForegroundColor Yellow
        
        # Get all assets with warranty information
        $allAssets = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/hardware?limit=1000" -Headers $Script:SnipeITHeaders
        
        $currentDate = Get-Date
        $futureDate = $currentDate.AddDays($DaysAhead)
        
        # Analyze warranty status
        $warrantyAnalysis = $allAssets.rows | ForEach-Object {
            $asset = $_
            $warrantyEnd = $null
            $status = "Unknown"
            
            if ($asset.warranty_expires)
            {
                $warrantyEnd = [DateTime]::Parse($asset.warranty_expires.date)
                
                if ($warrantyEnd -lt $currentDate)
                {
                    $status = "Expired"
                } elseif ($warrantyEnd -le $futureDate)
                {
                    $status = "Expiring Soon"
                } else
                {
                    $status = "Active"
                }
            } else
            {
                $status = "No Warranty Data"
            }
            
            # Create a properly formatted purchase cost 
            $purchaseCost = $null
            if ($asset.purchase_cost -and ![string]::IsNullOrEmpty($asset.purchase_cost))
            {
                # Try to parse the purchase cost - if successful use the value, otherwise use null
                $tempCost = 0
                if ([double]::TryParse($asset.purchase_cost, [ref]$tempCost)) {
                    $purchaseCost = $tempCost
                }
            }
            
            [PSCustomObject]@{
                AssetTag = $asset.asset_tag
                Name = $asset.name
                Model = $asset.model.name
                WarrantyExpires = if ($asset.warranty_expires)
                { $asset.warranty_expires.date 
                } else
                { "N/A" 
                }
                Status = $status
                DaysRemaining = if ($warrantyEnd)
                { [math]::Max(0, ($warrantyEnd - $currentDate).Days) 
                } else
                { $null 
                }
                PurchaseDate = if ($asset.purchase_date)
                { $asset.purchase_date.date 
                } else
                { "N/A" 
                }
                PurchaseCost = $purchaseCost
                PurchaseCostDisplay = if ($purchaseCost -ne $null) 
                { $purchaseCost.ToString("0.00") 
                } else 
                { "N/A" 
                }
            }
        }
        
        # Calculate summary statistics
        $expired = $warrantyAnalysis | Where-Object Status -eq "Expired"
        $expiringSoon = $warrantyAnalysis | Where-Object Status -eq "Expiring Soon"
        $active = $warrantyAnalysis | Where-Object Status -eq "Active"
        $noData = $warrantyAnalysis | Where-Object Status -eq "No Warranty Data"
        
        # Identify high-risk gaps (expensive assets with expired/no warranty)
        # Only include assets where PurchaseCost is a number and >= 1000
        $highRiskAssets = $warrantyAnalysis | Where-Object {
            ($_.Status -eq "Expired" -or $_.Status -eq "No Warranty Data") -and
            $_.PurchaseCost -ne $null -and $_.PurchaseCost -ge 1000
        } | Sort-Object -Property PurchaseCost -Descending
        
        # Create summary
        $warrantySummary = [PSCustomObject]@{
            TotalAssets = $warrantyAnalysis.Count
            ActiveWarranties = $active.Count
            ExpiredWarranties = $expired.Count
            ExpiringSoon = $expiringSoon.Count
            NoWarrantyData = $noData.Count
            HighRiskGaps = $highRiskAssets.Count
            DaysAheadChecked = $DaysAhead
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Display results
        Write-Host "`n=== WARRANTY STATUS SUMMARY ===" -ForegroundColor Cyan
        $warrantySummary | Format-List
        
        if ($expiringSoon.Count -gt 0)
        {
            Write-Host "`n=== WARRANTIES EXPIRING SOON ===" -ForegroundColor Yellow
            $expiringSoon | Sort-Object DaysRemaining | Select-Object -First 10 |
                Format-Table AssetTag, Name, WarrantyExpires, DaysRemaining -AutoSize
        }
        
        if ($highRiskAssets.Count -gt 0)
        {
            Write-Host "`n=== HIGH-RISK WARRANTY GAPS ===" -ForegroundColor Red
            $highRiskAssets | Select-Object -First 10 |
                Format-Table AssetTag, Name, Status, PurchaseCostDisplay -AutoSize
        }
        
        # Export if requested - use PurchaseCostDisplay for exporting
        if ($ExportPath)
        {
            # Create a proper export object that uses the display value instead of the numeric value
            $exportData = $warrantyAnalysis | Select-Object AssetTag, Name, Model, WarrantyExpires, 
                Status, DaysRemaining, PurchaseDate, 
                @{Name='PurchaseCost'; Expression={$_.PurchaseCostDisplay}}
                
            Export-ToCsv -Data $exportData -FilePath $ExportPath
        }
        
        return @{
            Summary = $warrantySummary
            Expired = $expired
            ExpiringSoon = $expiringSoon
            HighRiskGaps = $highRiskAssets
            AllWarrantyData = $warrantyAnalysis
        }
    } catch
    {
        Write-Error "Failed to retrieve warranty status: $($_.Exception.Message)"
    }
}


#endregion

#region 6. License Management

<#
.SYNOPSIS
    Analyzes software license status and utilization.

.DESCRIPTION
    Retrieves statistics on active, expired, and soon-to-expire licenses,
    along with utilization rates and top software by license count.

.PARAMETER ExportPath
    Optional path to export results to CSV

.PARAMETER DaysAhead
    Number of days ahead to look for expiring licenses (default: 60)

.EXAMPLE
    Get-LicenseManagement
    Get-LicenseManagement -DaysAhead 30 -ExportPath "C:\Reports\license_status.csv"
#>
function Get-LicenseManagement
{
    [CmdletBinding()]
    param(
        [string]$ExportPath,
        [int]$DaysAhead = 60
    )
    
    if (-not $Script:SnipeITBaseUrl)
    {
        throw "Snipe-IT connection not initialized. Run Initialize-SnipeITConnection first."
    }
    
    try
    {
        Write-Host "Analyzing license management..." -ForegroundColor Yellow
        
        # Get all licenses
        $allLicenses = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/licenses?limit=500" -Headers $Script:SnipeITHeaders
        
        $currentDate = Get-Date
        $futureDate = $currentDate.AddDays($DaysAhead)
        
        # Analyze license status
        $licenseAnalysis = $allLicenses.rows | ForEach-Object {
            $license = $_
            $expirationDate = $null
            $status = "Unknown"
            
            if ($license.expiration_date)
            {
                $expirationDate = [DateTime]::Parse($license.expiration_date.date)
                
                if ($expirationDate -lt $currentDate)
                {
                    $status = "Expired"
                } elseif ($expirationDate -le $futureDate)
                {
                    $status = "Expiring Soon"
                } else
                {
                    $status = "Active"
                }
            } else
            {
                $status = "No Expiration"
            }
            
            # Calculate utilization rate
            $utilizationRate = if ($license.seats -gt 0)
            {
                [math]::Round(($license.free_seats_count / $license.seats) * 100, 2)
            } else
            { 0 
            }
            
            [PSCustomObject]@{
                Name = $license.name
                Category = if ($license.category)
                { $license.category.name 
                } else
                { "N/A" 
                }
                TotalSeats = $license.seats
                UsedSeats = $license.seats - $license.free_seats_count
                FreeSeats = $license.free_seats_count
                UtilizationRate = "$utilizationRate%"
                ExpirationDate = if ($license.expiration_date)
                { $license.expiration_date.date 
                } else
                { "No Expiration" 
                }
                Status = $status
                DaysRemaining = if ($expirationDate)
                { [math]::Max(0, ($expirationDate - $currentDate).Days) 
                } else
                { $null 
                }
                PurchaseCost = if ($license.purchase_cost)
                { $license.purchase_cost 
                } else
                { "N/A" 
                }
                Supplier = if ($license.supplier)
                { $license.supplier.name 
                } else
                { "N/A" 
                }
            }
        }
        
        # Calculate summary statistics
        $active = $licenseAnalysis | Where-Object Status -eq "Active"
        $expired = $licenseAnalysis | Where-Object Status -eq "Expired"
        $expiringSoon = $licenseAnalysis | Where-Object Status -eq "Expiring Soon"
        $noExpiration = $licenseAnalysis | Where-Object Status -eq "No Expiration"
        
        # Top software by license count
        $topSoftware = $licenseAnalysis | Group-Object Category | 
            Sort-Object Count -Descending | Select-Object -First 10 |
            ForEach-Object {
                [PSCustomObject]@{
                    Category = $_.Name
                    LicenseCount = $_.Count
                    TotalSeats = ($_.Group | Measure-Object TotalSeats -Sum).Sum
                    TotalCost = ($_.Group | Where-Object { $_.PurchaseCost -ne "N/A" } | 
                            Measure-Object @{Expression={[double]$_.PurchaseCost}} -Sum).Sum
                    }
                }
        
        # License utilization statistics
        $utilizationStats = $licenseAnalysis | Where-Object { $_.TotalSeats -gt 0 } |
            ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    UtilizationRate = [double]($_.UtilizationRate -replace '%', '')
                    TotalSeats = $_.TotalSeats
                    UsedSeats = $_.UsedSeats
                }
            } | Sort-Object UtilizationRate -Descending
        
        # Create summary
        $licenseSummary = [PSCustomObject]@{
            TotalLicenses = $licenseAnalysis.Count
            ActiveLicenses = $active.Count
            ExpiredLicenses = $expired.Count
            ExpiringSoon = $expiringSoon.Count
            NoExpirationSet = $noExpiration.Count
            AverageUtilization = [math]::Round(($utilizationStats | Measure-Object UtilizationRate -Average).Average, 2)
            DaysAheadChecked = $DaysAhead
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Display results
        Write-Host "`n=== LICENSE MANAGEMENT SUMMARY ===" -ForegroundColor Cyan
        $licenseSummary | Format-List
        
        Write-Host "`n=== TOP SOFTWARE BY LICENSE COUNT ===" -ForegroundColor Cyan
        $topSoftware | Format-Table -AutoSize
        
        if ($expiringSoon.Count -gt 0)
        {
            Write-Host "`n=== LICENSES EXPIRING SOON ===" -ForegroundColor Yellow
            $expiringSoon | Sort-Object DaysRemaining | Select-Object -First 10 |
                Format-Table Name, ExpirationDate, DaysRemaining, TotalSeats -AutoSize
        }
        
        Write-Host "`n=== TOP LICENSE UTILIZATION ===" -ForegroundColor Cyan
        $utilizationStats | Select-Object -First 10 |
            Format-Table Name, UtilizationRate, UsedSeats, TotalSeats -AutoSize
        
        # Export if requested
        if ($ExportPath)
        {
            Export-ToCsv -Data $licenseAnalysis -FilePath $ExportPath
        }
        
        return @{
            Summary = $licenseSummary
            Active = $active
            Expired = $expired
            ExpiringSoon = $expiringSoon
            TopSoftware = $topSoftware
            UtilizationStats = $utilizationStats
            AllLicenseData = $licenseAnalysis
        }
    } catch
    {
        Write-Error "Failed to retrieve license management data: $($_.Exception.Message)"
    }
}

#endregion

#region 7. Location Distribution

<#
.SYNOPSIS
    Analyzes asset distribution across physical locations.

.DESCRIPTION
    Retrieves statistics on asset counts by location, identifies
    high-turnover locations, and provides distribution analytics.

.PARAMETER ExportPath
    Optional path to export results to CSV

.EXAMPLE
    Get-LocationDistribution
    Get-LocationDistribution -ExportPath "C:\Reports\location_distribution.csv"
#>


function Get-LocationDistribution {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$ExportPath
    )

    try {
        Write-Host "Analyzing location distribution..." -ForegroundColor Yellow
        
        # Get all assets
        $allAssets = Get-SnipeitAsset -all
        
        # Get all locations
        $locations = Get-SnipeitLocation -all
        
        # Create a hashtable to store the location distribution
        $locationStats = @{}
        
        # Initialize the hashtable with all locations
        foreach ($location in $locations) {
            $locationStats[$location.id] = @{
                'Name' = $location.name
                'Address' = $location.address
                'Count' = 0
                'Assets' = @()
            }
        }
        
        # Count assets per location
        foreach ($asset in $allAssets) {
            # Safe check for location_id
            if ($asset.location -and $asset.location.id) {
                $locationId = $asset.location.id
                $locationStats[$locationId]['Count']++
                
                # Create asset info object with safe date parsing
                $assetInfo = @{
                    'Tag' = $asset.asset_tag
                    'Name' = $asset.name
                    'Serial' = $asset.serial
                    'Status' = $asset.status_label.name
                }
                
                # Safely parse dates - handle empty values
                if (-not [string]::IsNullOrEmpty($asset.created_at)) {
                    $success = [DateTime]::TryParse($asset.created_at, [ref]$createdDate)
                    $assetInfo['Created'] = if ($success) { $createdDate } else { $null }
                } else {
                    $assetInfo['Created'] = $null
                }
                
                if (-not [string]::IsNullOrEmpty($asset.updated_at)) {
                    $success = [DateTime]::TryParse($asset.updated_at, [ref]$updatedDate)
                    $assetInfo['Updated'] = if ($success) { $updatedDate } else { $null }
                } else {
                    $assetInfo['Updated'] = $null
                }
                
                if (-not [string]::IsNullOrEmpty($asset.last_checkout)) {
                    $success = [DateTime]::TryParse($asset.last_checkout, [ref]$checkoutDate)
                    $assetInfo['LastCheckout'] = if ($success) { $checkoutDate } else { $null }
                } else {
                    $assetInfo['LastCheckout'] = $null
                }
                
                if (-not [string]::IsNullOrEmpty($asset.expected_checkin)) {
                    $success = [DateTime]::TryParse($asset.expected_checkin, [ref]$expectedDate)
                    $assetInfo['ExpectedCheckin'] = if ($success) { $expectedDate } else { $null }
                } else {
                    $assetInfo['ExpectedCheckin'] = $null
                }
                
                # Add asset to location's asset list
                $locationStats[$locationId]['Assets'] += $assetInfo
            }
        }
        
        # Display the results
        $sortedLocations = $locationStats.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending
        
        Write-Host "`nLocation Distribution:" -ForegroundColor Cyan
        Write-Host "======================" -ForegroundColor Cyan
        
        foreach ($location in $sortedLocations) {
            $locationData = $location.Value
            Write-Host "$($locationData.Name): $($locationData.Count) assets" -ForegroundColor Green
            
            if ($locationData.Count -gt 0) {
                $statusBreakdown = $locationData.Assets | Group-Object Status | 
                                  Select-Object @{N='Status';E={$_.Name}}, @{N='Count';E={$_.Count}}
                
                Write-Host "  Status Breakdown:" -ForegroundColor Yellow
                foreach ($status in $statusBreakdown) {
                    Write-Host "    $($status.Status): $($status.Count)" -ForegroundColor Gray
                }
            }
            
            Write-Host ""
        }
        
        # Export to CSV if requested
        if ($ExportPath) {
            $exportData = foreach ($location in $sortedLocations) {
                $locationData = $location.Value
                
                if ($locationData.Count -gt 0) {
                    foreach ($asset in $locationData.Assets) {
                        [PSCustomObject]@{
                            LocationName = $locationData.Name
                            LocationAddress = $locationData.Address
                            AssetTag = $asset.Tag
                            AssetName = $asset.Name
                            SerialNumber = $asset.Serial
                            Status = $asset.Status
                            Created = $asset.Created
                            Updated = $asset.Updated
                            LastCheckout = $asset.LastCheckout
                            ExpectedCheckin = $asset.ExpectedCheckin
                        }
                    }
                } else {
                    [PSCustomObject]@{
                        LocationName = $locationData.Name
                        LocationAddress = $locationData.Address
                        AssetTag = ""
                        AssetName = ""
                        SerialNumber = ""
                        Status = ""
                        Created = $null
                        Updated = $null
                        LastCheckout = $null
                        ExpectedCheckin = $null
                    }
                }
            }
            
            try {
                $exportData | Export-Csv -Path $ExportPath -NoTypeInformation
                Write-Host "Location distribution exported to: $ExportPath" -ForegroundColor Green
            } catch {
                Write-Host "Error exporting to CSV: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        return $true
    } catch {
        Write-Host "Failed to retrieve location distribution: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}





#endregion

#region 8. Depreciation Overview

<#
.SYNOPSIS
    Analyzes asset depreciation and current values.

.DESCRIPTION
    Compares current asset values against purchase costs,
    identifies most depreciated assets, and projects value loss.

.PARAMETER ExportPath
    Optional path to export results to CSV

.EXAMPLE
    Get-DepreciationOverview
    Get-DepreciationOverview -ExportPath "C:\Reports\depreciation.csv"
#>
function Get-DepreciationOverview
{
    [CmdletBinding()]
    param(
        [string]$ExportPath
    )
    
    if (-not $Script:SnipeITBaseUrl)
    {
        throw "Snipe-IT connection not initialized. Run Initialize-SnipeITConnection first."
    }
    
    try
    {
        Write-Host "Analyzing asset depreciation..." -ForegroundColor Yellow
        
        # Get all assets with depreciation information
        $allAssets = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/hardware?limit=1000" -Headers $Script:SnipeITHeaders
        
        # Filter assets with purchase cost data
        $assetsWithCost = $allAssets.rows | Where-Object { 
            $_.purchase_cost -and $_.purchase_cost -ne "0.00" 
        }
        
        # Analyze depreciation
        $depreciationAnalysis = $assetsWithCost | ForEach-Object {
            $asset = $_
            $purchaseCost = [double]$asset.purchase_cost
            $currentValue = if ($asset.book_value)
            { [double]$asset.book_value 
            } else
            { 0 
            }
            
            $depreciationAmount = $purchaseCost - $currentValue
            $depreciationPercent = if ($purchaseCost -gt 0)
            { 
                [math]::Round(($depreciationAmount / $purchaseCost) * 100, 2) 
            } else
            { 0 
            }
            
            # Calculate age in years
            $ageInYears = if ($asset.purchase_date)
            {
                $purchaseDate = [DateTime]::Parse($asset.purchase_date.date)
                [math]::Round(((Get-Date) - $purchaseDate).Days / 365.25, 2)
            } else
            { $null 
            }
            
            [PSCustomObject]@{
                AssetTag = $asset.asset_tag
                Name = $asset.name
                Model = if ($asset.model)
                { $asset.model.name 
                } else
                { "N/A" 
                }
                Category = if ($asset.category)
                { $asset.category.name 
                } else
                { "N/A" 
                }
                PurchaseCost = $purchaseCost
                CurrentValue = $currentValue
                DepreciationAmount = $depreciationAmount
                DepreciationPercent = $depreciationPercent
                AgeInYears = $ageInYears
                PurchaseDate = if ($asset.purchase_date)
                { $asset.purchase_date.date 
                } else
                { "N/A" 
                }
                Status = if ($asset.status_label)
                { $asset.status_label.name 
                } else
                { "N/A" 
                }
            }
        }
        
        # Calculate summary statistics
        $totalPurchaseCost = ($depreciationAnalysis | Measure-Object PurchaseCost -Sum).Sum
        $totalCurrentValue = ($depreciationAnalysis | Measure-Object CurrentValue -Sum).Sum
        $totalDepreciation = $totalPurchaseCost - $totalCurrentValue
        $overallDepreciationPercent = if ($totalPurchaseCost -gt 0)
        { 
            [math]::Round(($totalDepreciation / $totalPurchaseCost) * 100, 2) 
        } else
        { 0 
        }
        
        # Find most depreciated assets
        $mostDepreciated = $depreciationAnalysis | 
            Sort-Object DepreciationPercent -Descending | Select-Object -First 10
        
        # Find assets with highest depreciation amounts
        $highestDepreciationAmount = $depreciationAnalysis | 
            Sort-Object DepreciationAmount -Descending | Select-Object -First 10
        
        # Analyze by category
        $categoryAnalysis = $depreciationAnalysis | Group-Object Category |
            ForEach-Object {
                $categoryAssets = $_.Group
                $categoryPurchaseCost = ($categoryAssets | Measure-Object PurchaseCost -Sum).Sum
                $categoryCurrentValue = ($categoryAssets | Measure-Object CurrentValue -Sum).Sum
                $categoryDepreciation = $categoryPurchaseCost - $categoryCurrentValue
                
                [PSCustomObject]@{
                    Category = $_.Name
                    AssetCount = $categoryAssets.Count
                    TotalPurchaseCost = $categoryPurchaseCost
                    TotalCurrentValue = $categoryCurrentValue
                    TotalDepreciation = $categoryDepreciation
                    DepreciationPercent = if ($categoryPurchaseCost -gt 0)
                    { 
                        [math]::Round(($categoryDepreciation / $categoryPurchaseCost) * 100, 2) 
                    } else
                    { 0 
                    }
                    AverageAge = if ($categoryAssets.AgeInYears)
                    { 
                        [math]::Round(($categoryAssets | Where-Object AgeInYears | 
                                    Measure-Object AgeInYears -Average).Average, 2) 
                        } else
                        { $null 
                        }
                    }
                } | Sort-Object DepreciationPercent -Descending
        
        # Create summary
        $depreciationSummary = [PSCustomObject]@{
            TotalAssetsAnalyzed = $depreciationAnalysis.Count
            TotalPurchaseCost = $totalPurchaseCost
            TotalCurrentValue = $totalCurrentValue
            TotalDepreciation = $totalDepreciation
            OverallDepreciationPercent = "$overallDepreciationPercent%"
            AverageAssetAge = if ($depreciationAnalysis.AgeInYears)
            { 
                [math]::Round(($depreciationAnalysis | Where-Object AgeInYears | 
                            Measure-Object AgeInYears -Average).Average, 2) 
            } else
            { $null 
            }
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Display results
        Write-Host "`n=== DEPRECIATION OVERVIEW ===" -ForegroundColor Cyan
        $depreciationSummary | Format-List
        
        Write-Host "`n=== DEPRECIATION BY CATEGORY ===" -ForegroundColor Cyan
        $categoryAnalysis | Format-Table Category, AssetCount, DepreciationPercent, AverageAge -AutoSize
        
        Write-Host "`n=== MOST DEPRECIATED ASSETS (%) ===" -ForegroundColor Red
        $mostDepreciated | Format-Table AssetTag, Name, DepreciationPercent, AgeInYears -AutoSize
        
        Write-Host "`n=== HIGHEST DEPRECIATION AMOUNTS ===" -ForegroundColor Red
        $highestDepreciationAmount | Format-Table AssetTag, Name, DepreciationAmount, PurchaseCost -AutoSize
        
        # Export if requested
        if ($ExportPath)
        {
            Export-ToCsv -Data $depreciationAnalysis -FilePath $ExportPath
        }
        
        return @{
            Summary = $depreciationSummary
            DepreciationAnalysis = $depreciationAnalysis
            CategoryAnalysis = $categoryAnalysis
            MostDepreciated = $mostDepreciated
            HighestAmounts = $highestDepreciationAmount
        }
    } catch
    {
        Write-Error "Failed to retrieve depreciation overview: $($_.Exception.Message)"
    }
}

#endregion

#region 9. Repair and Maintenance

<#
.SYNOPSIS
    Analyzes repair and maintenance patterns for assets.

.DESCRIPTION
    Retrieves statistics on assets under repair, average repair times,
    and identifies high-maintenance assets.

.PARAMETER ExportPath
    Optional path to export results to CSV

.EXAMPLE
    Get-RepairMaintenance
    Get-RepairMaintenance -ExportPath "C:\Reports\repair_maintenance.csv"
#>
function Get-RepairMaintenance
{
    [CmdletBinding()]
    param(
        [string]$ExportPath
    )
    
    if (-not $Script:SnipeITBaseUrl)
    {
        throw "Snipe-IT connection not initialized. Run Initialize-SnipeITConnection first."
    }
    
    try
    {
        Write-Host "Analyzing repair and maintenance..." -ForegroundColor Yellow
        
        # Get assets currently under repair (assuming "Out for Repair" status)
        $allAssets = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/hardware?limit=1000" -Headers $Script:SnipeITHeaders
        $assetsUnderRepair = $allAssets.rows | Where-Object { 
            $_.status_label -and $_.status_label.name -like "*repair*" 
        }
        
        # Get activity log for repair-related activities
        $activityLog = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/reports/activity?limit=2000" -Headers $Script:SnipeITHeaders
        
        # Analyze repair activities (looking for checkout to repair and checkin from repair)
        $repairActivities = $activityLog.rows | Where-Object {
            $_.action_type -eq "checkout" -and $_.note -like "*repair*" -or
            $_.action_type -eq "checkin" -and $_.note -like "*repair*"
        }
        
        # Calculate repair times for completed repairs
        $completedRepairs = @()
        $repairCheckouts = $repairActivities | Where-Object { $_.action_type -eq "checkout" }
        
        foreach ($checkout in $repairCheckouts)
        {
            $checkin = $repairActivities | Where-Object {
                $_.action_type -eq "checkin" -and
                $_.item.id -eq $checkout.item.id -and
                [DateTime]::Parse($_.created_at.date) -gt [DateTime]::Parse($checkout.created_at.date)
            } | Sort-Object created_at.date | Select-Object -First 1
            
            if ($checkin)
            {
                $checkoutDate = [DateTime]::Parse($checkout.created_at.date)
                $checkinDate = [DateTime]::Parse($checkin.created_at.date)
                $repairDays = ($checkinDate - $checkoutDate).Days
                
                $completedRepairs += [PSCustomObject]@{
                    AssetTag = $checkout.item.asset_tag
                    AssetName = $checkout.item.name
                    Model = if ($checkout.item.model)
                    { $checkout.item.model.name 
                    } else
                    { "N/A" 
                    }
                    Category = if ($checkout.item.category)
                    { $checkout.item.category.name 
                    } else
                    { "N/A" 
                    }
                    RepairStartDate = $checkout.created_at.date
                    RepairEndDate = $checkin.created_at.date
                    RepairDays = $repairDays
                    RepairNote = $checkout.note
                }
            }
        }
        
        # Calculate average repair times by category
        $avgRepairByCategory = $completedRepairs | Group-Object Category |
            ForEach-Object {
                [PSCustomObject]@{
                    Category = $_.Name
                    RepairCount = $_.Count
                    AverageRepairDays = [math]::Round(($_.Group | Measure-Object RepairDays -Average).Average, 1)
                    MinRepairDays = ($_.Group | Measure-Object RepairDays -Minimum).Minimum
                    MaxRepairDays = ($_.Group | Measure-Object RepairDays -Maximum).Maximum
                }
            } | Sort-Object AverageRepairDays -Descending
        
        # Identify high-maintenance assets (multiple repairs)
        $highMaintenanceAssets = $completedRepairs | Group-Object AssetTag |
            Where-Object Count -gt 1 | ForEach-Object {
                $assetRepairs = $_.Group
                [PSCustomObject]@{
                    AssetTag = $_.Name
                    AssetName = ($assetRepairs | Select-Object -First 1).AssetName
                    Model = ($assetRepairs | Select-Object -First 1).Model
                    Category = ($assetRepairs | Select-Object -First 1).Category
                    RepairCount = $_.Count
                    TotalRepairDays = ($assetRepairs | Measure-Object RepairDays -Sum).Sum
                    AverageRepairDays = [math]::Round(($assetRepairs | Measure-Object RepairDays -Average).Average, 1)
                    LastRepairDate = ($assetRepairs | Sort-Object RepairEndDate -Descending | 
                            Select-Object -First 1).RepairEndDate
                    }
                } | Sort-Object RepairCount -Descending
        
        # Get current assets under repair with time analysis
        $currentRepairs = $assetsUnderRepair | ForEach-Object {
            $asset = $_
            # Find when this asset was sent for repair
            $lastRepairCheckout = $repairActivities | Where-Object {
                $_.action_type -eq "checkout" -and $_.item.id -eq $asset.id
            } | Sort-Object created_at.date -Descending | Select-Object -First 1
            
            $daysInRepair = if ($lastRepairCheckout)
            {
                ((Get-Date) - [DateTime]::Parse($lastRepairCheckout.created_at.date)).Days
            } else
            { $null 
            }
            
            [PSCustomObject]@{
                AssetTag = $asset.asset_tag
                AssetName = $asset.name
                Model = if ($asset.model)
                { $asset.model.name 
                } else
                { "N/A" 
                }
                Category = if ($asset.category)
                { $asset.category.name 
                } else
                { "N/A" 
                }
                Status = $asset.status_label.name
                DaysInRepair = $daysInRepair
                SentForRepair = if ($lastRepairCheckout)
                { $lastRepairCheckout.created_at.date 
                } else
                { "Unknown" 
                }
                Location = if ($asset.location)
                { $asset.location.name 
                } else
                { "N/A" 
                }
            }
        }
        
        # Calculate summary statistics
        $totalRepairs = $completedRepairs.Count
        $averageRepairTime = if ($totalRepairs -gt 0)
        {
            [math]::Round(($completedRepairs | Measure-Object RepairDays -Average).Average, 1)
        } else
        { 0 
        }
        
        # Create summary
        $repairSummary = [PSCustomObject]@{
            AssetsCurrentlyUnderRepair = $currentRepairs.Count
            CompletedRepairsAnalyzed = $totalRepairs
            AverageRepairTimeDays = $averageRepairTime
            HighMaintenanceAssets = $highMaintenanceAssets.Count
            CategoriesWithRepairs = $avgRepairByCategory.Count
            LongestCurrentRepair = if ($currentRepairs.DaysInRepair)
            { 
                ($currentRepairs | Where-Object DaysInRepair | 
                    Measure-Object DaysInRepair -Maximum).Maximum 
            } else
            { 0 
            }
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Display results
        Write-Host "`n=== REPAIR AND MAINTENANCE SUMMARY ===" -ForegroundColor Cyan
        $repairSummary | Format-List
        
        if ($currentRepairs.Count -gt 0)
        {
            Write-Host "`n=== ASSETS CURRENTLY UNDER REPAIR ===" -ForegroundColor Yellow
            $currentRepairs | Sort-Object DaysInRepair -Descending |
                Format-Table AssetTag, AssetName, DaysInRepair, SentForRepair -AutoSize
        }
        
        Write-Host "`n=== AVERAGE REPAIR TIME BY CATEGORY ===" -ForegroundColor Cyan
        $avgRepairByCategory | Format-Table -AutoSize
        
        if ($highMaintenanceAssets.Count -gt 0)
        {
            Write-Host "`n=== HIGH-MAINTENANCE ASSETS ===" -ForegroundColor Red
            $highMaintenanceAssets | Select-Object -First 10 |
                Format-Table AssetTag, AssetName, RepairCount, TotalRepairDays -AutoSize
        }
        
        # Export if requested
        if ($ExportPath)
        {
            $exportData = $currentRepairs + $completedRepairs + $highMaintenanceAssets
            Export-ToCsv -Data $exportData -FilePath $ExportPath
        }
        
        return @{
            Summary = $repairSummary
            CurrentRepairs = $currentRepairs
            CompletedRepairs = $completedRepairs
            CategoryAnalysis = $avgRepairByCategory
            HighMaintenanceAssets = $highMaintenanceAssets
        }
    } catch
    {
        Write-Error "Failed to retrieve repair and maintenance data: $($_.Exception.Message)"
    }
}

#endregion

#region 10. Vendor Performance

<#
.SYNOPSIS
    Analyzes vendor performance metrics and support statistics.

.DESCRIPTION
    Retrieves statistics on asset counts by vendor, support resolution times,
    and identifies high-support volume vendors.

.PARAMETER ExportPath
    Optional path to export results to CSV

.EXAMPLE
    Get-VendorPerformance
    Get-VendorPerformance -ExportPath "C:\Reports\vendor_performance.csv"
#>
function Get-VendorPerformance
{
    [CmdletBinding()]
    param(
        [string]$ExportPath
    )
    
    if (-not $Script:SnipeITBaseUrl)
    {
        throw "Snipe-IT connection not initialized. Run Initialize-SnipeITConnection first."
    }
    
    try
    {
        Write-Host "Analyzing vendor performance..." -ForegroundColor Yellow
        
        # Get all suppliers/vendors
        $allSuppliers = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/suppliers?limit=500" -Headers $Script:SnipeITHeaders
        
        # Get all assets to analyze by supplier
        $allAssets = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/hardware?limit=1000" -Headers $Script:SnipeITHeaders
        
        # Analyze assets by vendor
        $vendorAnalysis = $allSuppliers.rows | ForEach-Object {
            $supplier = $_
            $supplierAssets = $allAssets.rows | Where-Object { 
                $_.supplier -and $_.supplier.id -eq $supplier.id 
            }
            
            # Calculate total purchase value
            $totalValue = ($supplierAssets | Where-Object purchase_cost | 
                    Measure-Object @{Expression={[double]$_.purchase_cost}} -Sum).Sum
            
                # Analyze asset status distribution
                $statusDistribution = $supplierAssets | Group-Object { $_.status_label.name } |
                    ForEach-Object {
                        @{ $_.Name = $_.Count }
                    }
            
                    [PSCustomObject]@{
                        VendorName = $supplier.name
                        VendorId = $supplier.id
                        AssetCount = $supplierAssets.Count
                        TotalPurchaseValue = $totalValue
                        AveragePurchaseValue = if ($supplierAssets.Count -gt 0)
                        { 
                            [math]::Round($totalValue / $supplierAssets.Count, 2) 
                        } else
                        { 0 
                        }
                        ActiveAssets = ($supplierAssets | Where-Object { $_.status_label.name -eq "Ready to Deploy" }).Count
                        InRepairAssets = ($supplierAssets | Where-Object { $_.status_label.name -like "*repair*" }).Count
                        RetiredAssets = ($supplierAssets | Where-Object { $_.status_label.name -eq "Archived" }).Count
                        ContactEmail = $supplier.email
                        ContactPhone = $supplier.phone
                        Website = $supplier.url
                    }
                } | Sort-Object AssetCount -Descending
        
        # Get activity log to analyze support interactions
        $activityLog = Invoke-ApiCall -Url "$Script:SnipeITBaseUrl/api/v1/reports/activity?limit=1500" -Headers $Script:SnipeITHeaders
        
        # Analyze vendor-related activities (repairs, warranties, etc.)
        $vendorActivities = $activityLog.rows | Where-Object {
            $_.note -like "*support*" -or $_.note -like "*warranty*" -or 
            $_.note -like "*vendor*" -or $_.note -like "*supplier*"
        } | ForEach-Object {
            $activity = $_
            $supplierName = "Unknown"
            
            # Try to identify supplier from asset
            if ($activity.item -and $activity.item.supplier)
            {
                $supplierName = $activity.item.supplier.name
            }
            
            [PSCustomObject]@{
                VendorName = $supplierName
                AssetTag = if ($activity.item)
                { $activity.item.asset_tag 
                } else
                { "N/A" 
                }
                ActivityType = $activity.action_type
                ActivityDate = $activity.created_at.date
                Note = $activity.note
                AdminUser = if ($activity.admin)
                { $activity.admin.name 
                } else
                { "System" 
                }
            }
        }
        
        # Analyze support volume by vendor
        $supportVolume = $vendorActivities | Group-Object VendorName |
            ForEach-Object {
                [PSCustomObject]@{
                    VendorName = $_.Name
                    SupportTickets = $_.Count
                    WarrantyIssues = ($_.Group | Where-Object Note -like "*warranty*").Count
                    RepairIssues = ($_.Group | Where-Object Note -like "*repair*").Count
                    LastSupportDate = ($_.Group | Sort-Object ActivityDate -Descending | 
                            Select-Object -First 1).ActivityDate
                    }
                } | Sort-Object SupportTickets -Descending
        
        # Calculate vendor reliability score (fewer repairs/issues = higher score)
        $vendorReliability = $vendorAnalysis | ForEach-Object {
            $vendor = $_
            $supportData = $supportVolume | Where-Object VendorName -eq $vendor.VendorName
            
            $supportTicketsPerAsset = if ($vendor.AssetCount -gt 0 -and $supportData)
            { 
                [math]::Round($supportData.SupportTickets / $vendor.AssetCount, 2) 
            } else
            { 0 
            }
            
            $repairRate = if ($vendor.AssetCount -gt 0)
            { 
                [math]::Round(($vendor.InRepairAssets / $vendor.AssetCount) * 100, 2) 
            } else
            { 0 
            }
            
            # Simple reliability score (lower is better)
            $reliabilityScore = $supportTicketsPerAsset + ($repairRate / 10)
            
            [PSCustomObject]@{
                VendorName = $vendor.VendorName
                AssetCount = $vendor.AssetCount
                SupportTicketsPerAsset = $supportTicketsPerAsset
                RepairRate = "$repairRate%"
                ReliabilityScore = [math]::Round($reliabilityScore, 2)
                TotalPurchaseValue = $vendor.TotalPurchaseValue
            }
        } | Sort-Object ReliabilityScore
        
        # Calculate summary statistics
        $totalVendors = $vendorAnalysis.Count
        $vendorsWithAssets = ($vendorAnalysis | Where-Object AssetCount -gt 0).Count
        $totalSupportTickets = ($supportVolume | Measure-Object SupportTickets -Sum).Sum
        
        # Create summary
        $vendorSummary = [PSCustomObject]@{
            TotalVendors = $totalVendors
            VendorsWithAssets = $vendorsWithAssets
            TotalSupportTickets = $totalSupportTickets
            TopVendorByAssets = ($vendorAnalysis | Select-Object -First 1).VendorName
            TopVendorBySupportVolume = ($supportVolume | Select-Object -First 1).VendorName
            AverageSupportTicketsPerVendor = if ($vendorsWithAssets -gt 0)
            { 
                [math]::Round($totalSupportTickets / $vendorsWithAssets, 2) 
            } else
            { 0 
            }
            LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        # Display results
        Write-Host "`n=== VENDOR PERFORMANCE SUMMARY ===" -ForegroundColor Cyan
        $vendorSummary | Format-List
        
        Write-Host "`n=== TOP 10 VENDORS BY ASSET COUNT ===" -ForegroundColor Cyan
        $vendorAnalysis | Select-Object -First 10 |
            Format-Table VendorName, AssetCount, TotalPurchaseValue, ActiveAssets, InRepairAssets -AutoSize
        
        Write-Host "`n=== TOP SUPPORT VOLUME VENDORS ===" -ForegroundColor Cyan
        $supportVolume | Select-Object -First 10 |
            Format-Table VendorName, SupportTickets, WarrantyIssues, RepairIssues -AutoSize
        
        Write-Host "`n=== VENDOR RELIABILITY SCORES ===" -ForegroundColor Cyan
        $vendorReliability | Select-Object -First 10 |
            Format-Table VendorName, AssetCount, SupportTicketsPerAsset, RepairRate, ReliabilityScore -AutoSize
        
        # Export if requested
        if ($ExportPath)
        {
            $exportData = $vendorAnalysis + $supportVolume + $vendorReliability
            Export-ToCsv -Data $exportData -FilePath $ExportPath
        }
        
        return @{
            Summary = $vendorSummary
            VendorAnalysis = $vendorAnalysis
            SupportVolume = $supportVolume
            ReliabilityScores = $vendorReliability
            VendorActivities = $vendorActivities
        }
    } catch
    {
        Write-Error "Failed to retrieve vendor performance data: $($_.Exception.Message)"
    }
}

#endregion

#region Main Menu Function

<#
.SYNOPSIS
    Displays an interactive menu for asset management reporting.

.DESCRIPTION
    Provides a comprehensive menu system for accessing all reporting functions
    with options for data export and real-time updates.

.EXAMPLE
    Show-AssetManagementMenu
#>
function Show-AssetManagementMenu
{
    [CmdletBinding()]
    param()
    
    do
    {
        Clear-Host
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host "    ASSET MANAGEMENT REPORTING MENU" -ForegroundColor White
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "1.  Asset Overview" -ForegroundColor Green
        Write-Host "2.  Audit Status" -ForegroundColor Green  
        Write-Host "3.  Cart Utilization" -ForegroundColor Green
        Write-Host "4.  User Activity" -ForegroundColor Green
        Write-Host "5.  Warranty Status" -ForegroundColor Green
        Write-Host "6.  License Management" -ForegroundColor Green
        Write-Host "7.  Location Distribution" -ForegroundColor Green
        Write-Host "8.  Depreciation Overview" -ForegroundColor Green
        Write-Host "9.  Repair and Maintenance" -ForegroundColor Green
        Write-Host "10. Vendor Performance" -ForegroundColor Green
        Write-Host ""
        Write-Host "C.  Configure Connections" -ForegroundColor Yellow
        Write-Host "A.  Run All Reports" -ForegroundColor Magenta
        Write-Host "Q.  Quit" -ForegroundColor Red
        Write-Host ""
        
        $choice = Read-Host "Select an option (1-10, C, A, Q)"
        
        switch ($choice.ToUpper())
        {
            "1"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-AssetOverview 
                } else
                { Get-AssetOverview -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "2"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-AuditStatus 
                } else
                { Get-AuditStatus -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "3"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-CartUtilization 
                } else
                { Get-CartUtilization -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "4"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-UserActivity 
                } else
                { Get-UserActivity -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "5"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                $daysAhead = Read-Host "Days ahead for expiring warranties (default 90)"
                if ([string]::IsNullOrEmpty($daysAhead))
                { $daysAhead = 90 
                }
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-WarrantyStatus -DaysAhead $daysAhead 
                } else
                { Get-WarrantyStatus -DaysAhead $daysAhead -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "6"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                $daysAhead = Read-Host "Days ahead for expiring licenses (default 60)"
                if ([string]::IsNullOrEmpty($daysAhead))
                { $daysAhead = 60 
                }
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-LicenseManagement -DaysAhead $daysAhead 
                } else
                { Get-LicenseManagement -DaysAhead $daysAhead -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "7"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-LocationDistribution 
                } else
                { Get-LocationDistribution -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "8"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-DepreciationOverview 
                } else
                { Get-DepreciationOverview -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "9"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-RepairMaintenance 
                } else
                { Get-RepairMaintenance -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "10"
            { 
                $exportPath = Read-Host "Export path (press Enter to skip)"
                if ([string]::IsNullOrEmpty($exportPath))
                { Get-VendorPerformance 
                } else
                { Get-VendorPerformance -ExportPath $exportPath 
                }
                Read-Host "Press Enter to continue"
            }
            "C"
            {
                Write-Host "`nConfiguring Connections..." -ForegroundColor Yellow
                
                # Snipe-IT Configuration
                $snipeUrl = Read-Host "Enter Snipe-IT Base URL"
                $snipeToken = Read-Host "Enter Snipe-IT API Token" -AsSecureString
                $snipeTokenText = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($snipeToken))
                
                try
                {
                    Initialize-SnipeITConnection -BaseUrl $snipeUrl -ApiToken $snipeTokenText
                    Write-Host "Snipe-IT connection configured successfully!" -ForegroundColor Green
                } catch
                {
                    Write-Host "Failed to configure Snipe-IT: $($_.Exception.Message)" -ForegroundColor Red
                }
                
                # SNITEPS Configuration
                $snitepsUrl = Read-Host "Enter SNITEPS Base URL"
                $snitepsUser = Read-Host "Enter SNITEPS Username"
                $snitepsPass = Read-Host "Enter SNITEPS Password" -AsSecureString
                $snitepsPassText = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($snitepsPass))
                
                try
                {
                    Initialize-SNITEPSConnection -BaseUrl $snitepsUrl -Username $snitepsUser -Password $snitepsPassText
                    Write-Host "SNITEPS connection configured successfully!" -ForegroundColor Green
                } catch
                {
                    Write-Host "Failed to configure SNITEPS: $($_.Exception.Message)" -ForegroundColor Red
                }
                
                Read-Host "Press Enter to continue"
            }
            "A"
            {
                Write-Host "`nRunning all reports..." -ForegroundColor Magenta
                $baseExportPath = Read-Host "Enter base export directory (or press Enter to skip exports)"
                
                $reports = @(
                    @{ Name = "Asset Overview"; Function = "Get-AssetOverview"; File = "asset_overview.csv" },
                    @{ Name = "Audit Status"; Function = "Get-AuditStatus"; File = "audit_status.csv" },
                    @{ Name = "Cart Utilization"; Function = "Get-CartUtilization"; File = "cart_utilization.csv" },
                    @{ Name = "User Activity"; Function = "Get-UserActivity"; File = "user_activity.csv" },
                    @{ Name = "Warranty Status"; Function = "Get-WarrantyStatus"; File = "warranty_status.csv" },
                    @{ Name = "License Management"; Function = "Get-LicenseManagement"; File = "license_management.csv" },
                    @{ Name = "Location Distribution"; Function = "Get-LocationDistribution"; File = "location_distribution.csv" },
                    @{ Name = "Depreciation Overview"; Function = "Get-DepreciationOverview"; File = "depreciation_overview.csv" },
                    @{ Name = "Repair and Maintenance"; Function = "Get-RepairMaintenance"; File = "repair_maintenance.csv" },
                    @{ Name = "Vendor Performance"; Function = "Get-VendorPerformance"; File = "vendor_performance.csv" }
                )
                
                foreach ($report in $reports)
                {
                    Write-Host "Running $($report.Name)..." -ForegroundColor Yellow
                    try
                    {
                        if (![string]::IsNullOrEmpty($baseExportPath))
                        {
                            $exportPath = Join-Path $baseExportPath $report.File
                            & $report.Function -ExportPath $exportPath
                        } else
                        {
                            & $report.Function
                        }
                        Write-Host "$($report.Name) completed successfully!" -ForegroundColor Green
                    } catch
                    {
                        Write-Host "Failed to run $($report.Name): $($_.Exception.Message)" -ForegroundColor Red
                    }
                    Write-Host ""
                }
                
                Read-Host "All reports completed. Press Enter to continue"
            }
            "Q"
            { 
                Write-Host "Goodbye!" -ForegroundColor Green
                break
            }
            default
            { 
                Write-Host "Invalid option. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } while ($true)
}

#endregion

# Export module members
Export-ModuleMember -Function @(
    'Initialize-SnipeITConnection',
    'Initialize-SNITEPSConnection',
    'Get-AssetOverview',
    'Get-AuditStatus',
    'Get-CartUtilization',
    'Get-UserActivity',
    'Get-WarrantyStatus',
    'Get-LicenseManagement',
    'Get-LocationDistribution',
    'Get-DepreciationOverview',
    'Get-RepairMaintenance',
    'Get-VendorPerformance',
    'Show-AssetManagementMenu'
)
