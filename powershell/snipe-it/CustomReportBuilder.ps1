<#
.SYNOPSIS
    Custom Report Builder for Snipe-IT
.DESCRIPTION
    This script provides a flexible report builder for Snipe-IT, allowing users to create
    customized reports for assets, users, licenses, and maintenance records.
.NOTES
    Requires the snipeitps PowerShell module to be installed.
    This script is designed to be called from the main Snipe-IT Management Console,
    but can also be run independently.
.PARAMETER APIKey
    The Snipe-IT API key to use for authentication.
.PARAMETER SnipeURL
    The URL of the Snipe-IT instance.
#>

param (
  [Parameter(Mandatory = $false)]
  [string]$APIKey,
    
  [Parameter(Mandatory = $false)]
  [string]$SnipeURL
)

# Function to create custom reports in Snipe-IT
function Create-CustomReport
{
  # Check if module is loaded
  if (-not (Get-Module -Name "snipeitps"))
  {
    # Try to import the module
    try
    {
      Import-Module snipeitps -ErrorAction Stop
    } catch
    {
      Write-Host "SnipeIT PS module not found or couldn't be loaded. Installing now..." -ForegroundColor Yellow
      try
      {
        Install-Module -Name snipeitps -Force -Scope CurrentUser
        Import-Module snipeitps
      } catch
      {
        Write-Host "Failed to install snipeitps module. Please install it manually." -ForegroundColor Red
        return
      }
    }
  }

  # If API credentials weren't passed as parameters, check if they're already configured
  if ([string]::IsNullOrEmpty($APIKey) -or [string]::IsNullOrEmpty($SnipeURL))
  {
    try
    {
      $configuredAPI = Get-SnipeitInfo
      if (-not $configuredAPI)
      {
        Write-Host "Snipe-IT API connection not configured. Please run this script with valid APIKey and SnipeURL parameters." -ForegroundColor Red
        return
      }
    } catch
    {
      Write-Host "Snipe-IT API connection not configured. Please run this script with valid APIKey and SnipeURL parameters." -ForegroundColor Red
      return
    }
  } else
  {
    # Configure connection with provided parameters
    Set-SnipeitInfo -URL $SnipeURL -APIKey $APIKey
  }

  Clear-Host
  Write-Host "===== Custom Report Builder =====" -ForegroundColor Cyan
    
  try
  {
    # Step 1: Select report type
    Write-Host "Select the type of report to create:" -ForegroundColor Green
    Write-Host "1. Asset Report" -ForegroundColor Yellow
    Write-Host "2. User Report" -ForegroundColor Yellow
    Write-Host "3. License Report" -ForegroundColor Yellow
    Write-Host "4. Maintenance Report" -ForegroundColor Yellow
    Write-Host "5. Return to Menu" -ForegroundColor Yellow
        
    $reportType = Read-Host "Enter your choice (1-5)"
        
    if ($reportType -eq "5")
    {
      return
    }
        
    # Step 2: Define filters and parameters
    $filters = @{}
    $reportName = ""
    $getFunction = ""
    $entityName = ""
        
    # Set up report-specific settings
    switch ($reportType)
    {
      "1" # Asset Report
      {
        $reportName = "Custom Asset Report"
        $getFunction = "Get-SnipeitAsset"
        $entityName = "assets"
                
        # Asset status filter
        $getStatuses = Read-Host "Filter by status? (Y/N)"
        if ($getStatuses -eq "Y" -or $getStatuses -eq "y")
        {
          $statuses = Get-SnipeitStatus
          if ($statuses)
          {
            Write-Host "Available Statuses:" -ForegroundColor Green
            for ($i = 0; $i -lt $statuses.Count; $i++)
            {
              Write-Host "[$i] $($statuses[$i].name)" -ForegroundColor Cyan
            }
                        
            $statusIndex = Read-Host "Select status number (or comma-separated list)"
            $statusIndices = $statusIndex -split ',' | ForEach-Object { $_.Trim() }
            $statusIds = @()
                        
            foreach ($idx in $statusIndices)
            {
              if ([int]::TryParse($idx, [ref]$null) -and [int]$idx -ge 0 -and [int]$idx -lt $statuses.Count)
              {
                $statusIds += $statuses[[int]$idx].id
              }
            }
                        
            if ($statusIds.Count -gt 0)
            {
              if ($statusIds.Count -eq 1)
              {
                $filters.Add("status_id", $statusIds[0])
              } else
              {
                # For multiple statuses, we'll need to filter post-query
                $filters.Add("_status_ids", $statusIds)
              }
            }
          }
        }
                
        # Add other asset filters here (Category, Location, Date range, Model, Free-text search)
        # ... (include the rest of the asset filters from the previous code)
      }
            
      "2" # User Report
      {
        $reportName = "Custom User Report"
        $getFunction = "Get-SnipeitUser"
        $entityName = "users"
                
        # Add user filters here (Active/Inactive, Department, Location, Free-text search)
        # ... (include the user filters from the previous code)
      }
            
      "3" # License Report
      {
        $reportName = "Custom License Report"
        $getFunction = "Get-SnipeitLicense"
        $entityName = "licenses"
                
        # Add license filters here (Expiration status, Manufacturer, Category, Free-text search)
        # ... (include the license filters from the previous code)
      }
            
      "4" # Maintenance Report
      {
        $reportName = "Custom Maintenance Report"
        $getFunction = "Get-SnipeitAssetMaintenance"
        $entityName = "maintenance records"
                
        # Add maintenance filters here (Maintenance type, Status, Date range, Free-text search)
        # ... (include the maintenance filters from the previous code)
      }
            
      default
      {
        Write-Host "Invalid report type. Returning to menu." -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
      }
    }
        
    # Step 3: Report Name
    $customReportName = Read-Host "Enter a name for this report (or leave blank for default)"
    if (-not [string]::IsNullOrWhiteSpace($customReportName))
    {
      $reportName = $customReportName
    }
        
    # Step 4: Retrieve and filter data
    Write-Host "`nRetrieving data for report..." -ForegroundColor Yellow
        
    # Determine which filter parameters can be passed directly to the API
    $apiParams = @{}
    $postFilters = @{}
        
    foreach ($key in $filters.Keys)
    {
      # Special filters that start with underscore require post-processing
      if ($key.StartsWith("_"))
      {
        $postFilters[$key] = $filters[$key]
      } else
      {
        $apiParams[$key] = $filters[$key]
      }
    }
        
    # Retrieve data
    $data = & $getFunction -all @apiParams
        
    if (-not $data -or $data.Count -eq 0)
    {
      Write-Host "No $entityName found matching your criteria." -ForegroundColor Red
      Start-Sleep -Seconds 3
      return
    }
        
    # Apply post-filters
    # ... (include the post-filtering logic from the previous code)
        
    # Step 5: Configure Columns and Format
    # ... (include the column configuration code from the previous code)
        
    # Step 6: Select output format
    Write-Host "`nSelect output format:" -ForegroundColor Green
    Write-Host "1. Display on screen" -ForegroundColor Yellow
    Write-Host "2. Export to CSV" -ForegroundColor Yellow
    Write-Host "3. Export to HTML" -ForegroundColor Yellow
        
    $outputFormat = Read-Host "Enter your choice (1-3)"
        
    switch ($outputFormat)
    {
      "1" # Display on screen
      {
        Clear-Host
        Write-Host "===== $reportName =====" -ForegroundColor Cyan
        Write-Host "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
        Write-Host "Total Records: $($data.Count)" -ForegroundColor Yellow
                
        # Display data in table format
        $data | Select-Object -Property $columns | Format-Table -AutoSize
                
        Write-Host "`nPress any key to continue..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
      }
            
      "2" # Export to CSV
      {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\SnipeIT_${reportName}_${timestamp}.csv"
        $defaultPath = $defaultPath -replace ' ', '_'
                
        $savePath = Read-Host "Enter path to save CSV file (default: $defaultPath)"
        if ([string]::IsNullOrWhiteSpace($savePath))
        {
          $savePath = $defaultPath
        }
                
        try
        {
          # Create custom objects with selected columns for export
          $exportData = $data | ForEach-Object {
            $item = $_
            $exportObj = [PSCustomObject]@{}
                        
            foreach ($column in $columns)
            {
              $value = & $column.Expression $item
              $exportObj | Add-Member -NotePropertyName $column.Label -NotePropertyValue $value
            }
                        
            $exportObj
          }
                    
          # Export to CSV
          $exportData | Export-Csv -Path $savePath -NoTypeInformation
                    
          Write-Host "Report exported successfully to: $savePath" -ForegroundColor Green
        } catch
        {
          Write-Host "Error exporting to CSV: $_" -ForegroundColor Red
        }
                
        Write-Host "`nPress any key to continue..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
      }
            
      "3" # Export to HTML
      {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\SnipeIT_${reportName}_${timestamp}.html"
        $defaultPath = $defaultPath -replace ' ', '_'
                
        $savePath = Read-Host "Enter path to save HTML file (default: $defaultPath)"
        if ([string]::IsNullOrWhiteSpace($savePath))
        {
          $savePath = $defaultPath
        }
                
        try
        {
          # Create custom objects with selected columns for export
          $exportData = $data | ForEach-Object {
            $item = $_
            $exportObj = [PSCustomObject]@{}
                        
            foreach ($column in $columns)
            {
              $value = & $column.Expression $item
              $exportObj | Add-Member -NotePropertyName $column.Label -NotePropertyValue $value
            }
                        
            $exportObj
          }
                    
          # HTML styling
          $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>$reportName</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        .report-info { margin-bottom: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #ddd; }
    </style>
</head>
<body>
    <h1>$reportName</h1>
    <div class="report-info">
        <p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p><strong>Total Records:</strong> $($data.Count)</p>
    </div>
"@
                    
          $htmlFooter = @"
</body>
</html>
"@
                    
          # Export to HTML
          $htmlBody = $exportData | ConvertTo-Html -Fragment
                    
          $htmlContent = $htmlHeader + $htmlBody + $htmlFooter
          $htmlContent | Out-File -FilePath $savePath
                    
          Write-Host "Report exported successfully to: $savePath" -ForegroundColor Green
                    
          # Offer to open the HTML file
          $openFile = Read-Host "Would you like to open the HTML report now? (Y/N)"
          if ($openFile -eq "Y" -or $openFile -eq "y")
          {
            Start-Process $savePath
          }
        } catch
        {
          Write-Host "Error exporting to HTML: $_" -ForegroundColor Red
        }
                
        Write-Host "`nPress any key to continue..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
      }
    }
        
    # Step 7: Option to save report definition
    $saveDefinition = Read-Host "`nWould you like to save this report definition for future use? (Y/N)"
    if ($saveDefinition -eq "Y" -or $saveDefinition -eq "y")
    {
      $reportDefinition = @{
        "ReportType" = $reportType
        "ReportName" = $reportName
        "Filters" = $filters
        "APIParams" = $apiParams
        "PostFilters" = $postFilters
      }
            
      $savedReportsFolder = Join-Path -Path $env:USERPROFILE -ChildPath "Documents\SnipeIT_Reports"
            
      # Create directory if it doesn't exist
      if (-not (Test-Path $savedReportsFolder))
      {
        New-Item -Path $savedReportsFolder -ItemType Directory | Out-Null
      }
            
      $definitionFileName = $reportName -replace '[^a-zA-Z0-9_]', '_'
      $definitionPath = Join-Path -Path $savedReportsFolder -ChildPath "$definitionFileName.xml"
            
      # Save definition as XML
      try
      {
        $reportDefinition | Export-Clixml -Path $definitionPath
        Write-Host "Report definition saved successfully to: $definitionPath" -ForegroundColor Green
      } catch
      {
        Write-Host "Error saving report definition: $_" -ForegroundColor Red
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

# Function to run a saved report definition
function Run-SavedReport
{
  # Check if module is loaded
  if (-not (Get-Module -Name "snipeitps"))
  {
    # Try to import the module
    try
    {
      Import-Module snipeitps -ErrorAction Stop
    } catch
    {
      Write-Host "SnipeIT PS module not found or couldn't be loaded. Installing now..." -ForegroundColor Yellow
      try
      {
        Install-Module -Name snipeitps -Force -Scope CurrentUser
        Import-Module snipeitps
      } catch
      {
        Write-Host "Failed to install snipeitps module. Please install it manually." -ForegroundColor Red
        return
      }
    }
  }

  # If API credentials weren't passed as parameters, check if they're already configured
  if ([string]::IsNullOrEmpty($APIKey) -or [string]::IsNullOrEmpty($SnipeURL))
  {
    try
    {
      $configuredAPI = Get-SnipeitInfo
      if (-not $configuredAPI)
      {
        Write-Host "Snipe-IT API connection not configured. Please run this script with valid APIKey and SnipeURL parameters." -ForegroundColor Red
        return
      }
    } catch
    {
      Write-Host "Snipe-IT API connection not configured. Please run this script with valid APIKey and SnipeURL parameters." -ForegroundColor Red
      return
    }
  } else
  {
    # Configure connection with provided parameters
    Set-SnipeitInfo -URL $SnipeURL -APIKey $APIKey
  }

  Clear-Host
  Write-Host "===== Run Saved Report =====" -ForegroundColor Cyan
    
  # Load saved report definitions
  $savedReportsFolder = Join-Path -Path $env:USERPROFILE -ChildPath "Documents\SnipeIT_Reports"
    
  if (-not (Test-Path $savedReportsFolder))
  {
    Write-Host "No saved reports found. Please create and save a report first." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    return
  }
    
  $savedReports = Get-ChildItem -Path $savedReportsFolder -Filter "*.xml"
    
  if ($savedReports.Count -eq 0)
  {
    Write-Host "No saved reports found. Please create and save a report first." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    return
  }
    
  # Display available saved reports
  Write-Host "Available saved reports:" -ForegroundColor Green
  for ($i = 0; $i -lt $savedReports.Count; $i++)
  {
    # Extract report name from filename
    $reportName = [System.IO.Path]::GetFileNameWithoutExtension($savedReports[$i].Name)
    $reportName = $reportName -replace '_', ' '
        
    Write-Host "[$i] $reportName" -ForegroundColor Yellow
  }
    
  $reportIndex = Read-Host "Select a report to run (or 'q' to return to menu)"
    
  if ($reportIndex -eq 'q')
  {
    return
  }
    
  if (-not [int]::TryParse($reportIndex, [ref]$null) -or 
    [int]$reportIndex -lt 0 -or 
    [int]$reportIndex -ge $savedReports.Count)
  {
    Write-Host "Invalid selection." -ForegroundColor Red
    Start-Sleep -Seconds 2
    return
  }
    
  $selectedReport = $savedReports[[int]$reportIndex]
    
  # Load report definition
  try
  {
    $reportDefinition = Import-Clixml -Path $selectedReport.FullName
        
    $reportType = $reportDefinition.ReportType
    $reportName = $reportDefinition.ReportName
    $filters = $reportDefinition.Filters
    $apiParams = $reportDefinition.APIParams
    $postFilters = $reportDefinition.PostFilters
        
    # Determine entity and function
    $getFunction = ""
    $entityName = ""
        
    switch ($reportType)
    {
      "1" # Asset Report
      {
        $getFunction = "Get-SnipeitAsset"
        $entityName = "assets"
      }
            
      "2" # User Report
      {
        $getFunction = "Get-SnipeitUser"
        $entityName = "users"
      }
            
      "3" # License Report
      {
        $getFunction = "Get-SnipeitLicense"
        $entityName = "licenses"
      }
            
      "4" # Maintenance Report
      {
        $getFunction = "Get-SnipeitAssetMaintenance"
        $entityName = "maintenance records"
      }
    }
        
    # Retrieve data
    Write-Host "`nRunning saved report: $reportName..." -ForegroundColor Yellow
        
    $data = & $getFunction -all @apiParams
        
    if (-not $data -or $data.Count -eq 0)
    {
      Write-Host "No $entityName found matching your criteria." -ForegroundColor Red
      Start-Sleep -Seconds 3
      return
    }
        
    # Apply post-filters
    # ... (include the post-filtering logic from the previous code)
        
    # Configure columns based on report type
    # ... (include the column configuration code from the previous code)
        
    # Output options
    # ... (include the output format code from the previous code)
  } catch
  {
    Write-Host "Error loading or running saved report: $_" -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
}

# Check if the script is being run standalone or imported
if ($MyInvocation.InvocationName -ne '.')
{
  # Script is being run directly, show a menu
  function Show-ReportBuilderMenu
  {
    Clear-Host
    Write-Host "===== Snipe-IT Report Builder =====" -ForegroundColor Cyan
    Write-Host "1. Create Custom Report" -ForegroundColor Green
    Write-Host "2. Run Saved Report" -ForegroundColor Green
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host "=============================" -ForegroundColor Cyan
        
    $choice = Read-Host "Select an option"
        
    switch ($choice)
    {
      "1"
      { Create-CustomReport 
      }
      "2"
      { Run-SavedReport 
      }
      "Q"
      { exit 
      }
      "q"
      { exit 
      }
      default
      { 
        Write-Host "Invalid option. Please try again." -ForegroundColor Red
        Start-Sleep -Seconds 2
        Show-ReportBuilderMenu
      }
    }
        
    # Return to menu after function completes
    Show-ReportBuilderMenu
  }
    
  # Start the standalone menu
  Show-ReportBuilderMenu
}

# Export functions for importing
Export-ModuleMember -Function Create-CustomReport, Run-SavedReport
