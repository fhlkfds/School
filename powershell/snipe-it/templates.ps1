<#
.SYNOPSIS
    Snipe-IT CSV Template Generator
.DESCRIPTION
    This script provides functions to generate CSV templates for importing various entities into Snipe-IT,
    including assets, users, licenses, and maintenance records.
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

# Check if module is loaded and API is configured when script is imported
function Initialize-SnipeITConnection
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
        return $false
      }
    }
  }

  # If API credentials weren't passed as parameters, check if they're already configured
  if ([string]::IsNullOrEmpty($Script:APIKey) -or [string]::IsNullOrEmpty($Script:SnipeURL))
  {
    try
    {
      $configuredAPI = Get-SnipeitInfo
      if (-not $configuredAPI)
      {
        Write-Host "Snipe-IT API connection not configured. Please run this script with valid APIKey and SnipeURL parameters." -ForegroundColor Red
        return $false
      }
    } catch
    {
      Write-Host "Snipe-IT API connection not configured. Please run this script with valid APIKey and SnipeURL parameters." -ForegroundColor Red
      return $false
    }
  } else
  {
    # Configure connection with provided parameters
    Set-SnipeitInfo -URL $Script:SnipeURL -APIKey $Script:APIKey
  }

  return $true
}

# Function to generate an Asset Import Template
function New-AssetImportTemplate
{
  Clear-Host
  Write-Host "===== Generate Asset Import Template =====" -ForegroundColor Cyan
    
  try
  {
    # Get the save location
    $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\snipeit_asset_template.csv"
    Write-Host "This will create a template CSV file for importing assets into Snipe-IT." -ForegroundColor Yellow
    $savePath = Read-Host "Enter a path to save the template (default: $defaultPath)"
        
    if ([string]::IsNullOrWhiteSpace($savePath))
    {
      $savePath = $defaultPath
    }

    # Check if connection is initialized
    if (-not (Initialize-SnipeITConnection))
    {
      return
    }
        
    # Gather available field options from Snipe-IT for user reference
    Write-Host "`nRetrieving reference data from Snipe-IT..." -ForegroundColor Green
        
    # Get models
    $models = Get-SnipeitModel
    $modelOptions = if ($models)
    {
      $models | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get statuses
    $statuses = Get-SnipeitStatus
    $statusOptions = if ($statuses)
    {
      $statuses | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get categories
    $categories = Get-SnipeitCategory
    $categoryOptions = if ($categories)
    {
      $categories | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get manufacturers
    $manufacturers = Get-SnipeitManufacturer
    $manufacturerOptions = if ($manufacturers)
    {
      $manufacturers | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get suppliers
    $suppliers = Get-SnipeitSupplier
    $supplierOptions = if ($suppliers)
    {
      $suppliers | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get locations
    $locations = Get-SnipeitLocation
    $locationOptions = if ($locations)
    {
      $locations | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get custom fields
    $customFields = Get-SnipeitCustomField
    $customFieldOptions = if ($customFields)
    {
      $customFields | Where-Object { $_.field_type -eq "asset" } | Select-Object -Property id, name, format, db_column_name | Sort-Object -Property db_column_name
    } else
    {
      @()
    }
        
    # Create the template content
    $templateHeaders = @(
      "asset_tag",
      "name",
      "model_id",
      "status_id",
      "category_id",
      "manufacturer_id",
      "supplier_id",
      "location_id",
      "purchase_date",
      "purchase_cost",
      "order_number",
      "warranty_months",
      "notes",
      "serial",
      "requestable"
    )
        
    # Add custom field headers
    $customFieldHeaders = @()
    if ($customFieldOptions.Count -gt 0)
    {
      foreach ($field in $customFieldOptions)
      {
        if (-not [string]::IsNullOrWhiteSpace($field.db_column_name))
        {
          $templateHeaders += "_snipeit_$($field.db_column_name)"
          $customFieldHeaders += "_snipeit_$($field.db_column_name)"
        }
      }
    }
        
    # Create header row
    $templateContent = $templateHeaders -join ","
    $templateContent += "`n"
        
    # Create example rows
    $exampleRow1 = @(
      '"A00001"',
      '"Dell Latitude E7470"',
      '1',  # model_id
      '2',  # status_id (ready to deploy)
      '3',  # category_id
      '1',  # manufacturer_id (Dell)
      '1',  # supplier_id
      '1',  # location_id
      '"2023-01-15"',
      '"1200.00"',
      '"PO-12345"',
      '36',
      '"Sample laptop for executive team"',
      '"ABC123XYZ"',
      'TRUE'
    )
        
    # Add custom field examples
    foreach ($header in $customFieldHeaders)
    {
      $exampleRow1 += '""'  # Empty custom field values
    }
        
    $exampleRow2 = @(
      '"A00002"',
      '"iPhone 13"',
      '2',  # model_id
      '2',  # status_id (ready to deploy)
      '7',  # category_id (Mobile Devices)
      '2',  # manufacturer_id (Apple)
      '2',  # supplier_id
      '2',  # location_id
      '"2023-02-20"',
      '"999.00"',
      '"PO-67890"',
      '12',
      '"Sample mobile phone"',
      '"IMEI123456789"',
      'FALSE'
    )
        
    # Add custom field examples
    foreach ($header in $customFieldHeaders)
    {
      $exampleRow2 += '""'  # Empty custom field values
    }
        
    # Add example rows to template
    $templateContent += $exampleRow1 -join ","
    $templateContent += "`n"
    $templateContent += $exampleRow2 -join ","
        
    # Save the template
    try
    {
      $templateContent | Out-File -FilePath $savePath -Encoding UTF8
      Write-Host "`nTemplate saved successfully to: $savePath" -ForegroundColor Green
            
      # Create reference guide for IDs
      $referenceGuide = "`n===== Reference Guide for Snipe-IT IDs =====" +
      "`n`nModels (model_id):`n"
            
      if ($modelOptions.Count -gt 0)
      {
        $referenceGuide += $modelOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No models found in your Snipe-IT instance.`n"
      }
            
      $referenceGuide += "`nStatuses (status_id):`n"
      if ($statusOptions.Count -gt 0)
      {
        $referenceGuide += $statusOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No statuses found in your Snipe-IT instance.`n"
      }
            
      $referenceGuide += "`nCategories (category_id):`n"
      if ($categoryOptions.Count -gt 0)
      {
        $referenceGuide += $categoryOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No categories found in your Snipe-IT instance.`n"
      }
            
      $referenceGuide += "`nManufacturers (manufacturer_id):`n"
      if ($manufacturerOptions.Count -gt 0)
      {
        $referenceGuide += $manufacturerOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No manufacturers found in your Snipe-IT instance.`n"
      }
            
      $referenceGuide += "`nSuppliers (supplier_id):`n"
      if ($supplierOptions.Count -gt 0)
      {
        $referenceGuide += $supplierOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No suppliers found in your Snipe-IT instance.`n"
      }
            
      $referenceGuide += "`nLocations (location_id):`n"
      if ($locationOptions.Count -gt 0)
      {
        $referenceGuide += $locationOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No locations found in your Snipe-IT instance.`n"
      }
            
      if ($customFieldOptions.Count -gt 0)
      {
        $referenceGuide += "`nCustom Fields:`n"
        $referenceGuide += $customFieldOptions | ForEach-Object { 
          "  Name: $($_.name)`n  Column: _snipeit_$($_.db_column_name)`n  Format: $($_.format)`n" 
        } | Out-String
      }
            
      $referenceGuide += "`n===== End of Reference Guide ====="
            
      # Save the reference guide
      $refPath = [System.IO.Path]::ChangeExtension($savePath, ".reference.txt")
      $referenceGuide | Out-File -FilePath $refPath -Encoding UTF8
      Write-Host "Reference guide saved to: $refPath" -ForegroundColor Green
            
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
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
  }
    
  Write-Host "`nPress any key to continue..."
  $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to generate a User Import Template
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

    # Check if connection is initialized
    if (-not (Initialize-SnipeITConnection))
    {
      return
    }
        
    # Gather available field options from Snipe-IT for user reference
    Write-Host "`nRetrieving reference data from Snipe-IT..." -ForegroundColor Green
        
    # Get departments
    $departments = Get-SnipeitDepartment
    $departmentOptions = if ($departments)
    {
      $departments | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get locations
    $locations = Get-SnipeitLocation
    $locationOptions = if ($locations)
    {
      $locations | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get custom fields
    $customFields = Get-SnipeitCustomField
    $customFieldOptions = if ($customFields)
    {
      $customFields | Where-Object { $_.field_type -eq "user" } | Select-Object -Property id, name, format, db_column_name | Sort-Object -Property db_column_name
    } else
    {
      @()
    }
        
    # Create the template content
    $templateHeaders = @(
      "first_name",
      "last_name",
      "username",
      "email",
      "password",
      "phone",
      "jobtitle",
      "employee_num",
      "department_id",
      "location_id",
      "manager_id",
      "ldap_import",
      "activated",
      "notes"
    )
        
    # Add custom field headers
    $customFieldHeaders = @()
    if ($customFieldOptions.Count -gt 0)
    {
      foreach ($field in $customFieldOptions)
      {
        if (-not [string]::IsNullOrWhiteSpace($field.db_column_name))
        {
          $templateHeaders += "_snipeit_$($field.db_column_name)"
          $customFieldHeaders += "_snipeit_$($field.db_column_name)"
        }
      }
    }
        
    # Create header row
    $templateContent = $templateHeaders -join ","
    $templateContent += "`n"
        
    # Create example rows
    $exampleRow1 = @(
      '"John"',
      '"Doe"',
      '"jdoe"',
      '"john.doe@example.com"',
      '"Password123"',  # password
      '"555-1234"',
      '"IT Manager"',
      '"EMP001"',
      '1',  # department_id
      '1',  # location_id
      '',  # manager_id
      'FALSE',
      'TRUE',
      '"Example notes for John"'
    )
        
    # Add custom field examples
    foreach ($header in $customFieldHeaders)
    {
      $exampleRow1 += '""'  # Empty custom field values
    }
        
    $exampleRow2 = @(
      '"Jane"',
      '"Smith"',
      '"jsmith"',
      '"jane.smith@example.com"',
      '"Password456"',  # password
      '"555-5678"',
      '"Developer"',
      '"EMP002"',
      '2',  # department_id
      '1',  # location_id
      '1',  # manager_id (John Doe)
      'FALSE',
      'TRUE',
      '"Example notes for Jane"'
    )
        
    # Add custom field examples
    foreach ($header in $customFieldHeaders)
    {
      $exampleRow2 += '""'  # Empty custom field values
    }
        
    # Add example rows to template
    $templateContent += $exampleRow1 -join ","
    $templateContent += "`n"
    $templateContent += $exampleRow2 -join ","
        
    # Save the template
    try
    {
      $templateContent | Out-File -FilePath $savePath -Encoding UTF8
      Write-Host "`nTemplate saved successfully to: $savePath" -ForegroundColor Green
            
      # Create reference guide for IDs
      $referenceGuide = "`n===== Reference Guide for Snipe-IT IDs =====" +
      "`n`nDepartments (department_id):`n"
            
      if ($departmentOptions.Count -gt 0)
      {
        $referenceGuide += $departmentOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No departments found in your Snipe-IT instance.`n"
      }
            
      $referenceGuide += "`nLocations (location_id):`n"
      if ($locationOptions.Count -gt 0)
      {
        $referenceGuide += $locationOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No locations found in your Snipe-IT instance.`n"
      }
            
      if ($customFieldOptions.Count -gt 0)
      {
        $referenceGuide += "`nCustom Fields:`n"
        $referenceGuide += $customFieldOptions | ForEach-Object { 
          "  Name: $($_.name)`n  Column: _snipeit_$($_.db_column_name)`n  Format: $($_.format)`n" 
        } | Out-String
      }
            
      $referenceGuide += "`n===== Notes =====`n" +
      "- password: If omitted, a random password will be generated.`n" +
      "- ldap_import: Set to TRUE if the user should be managed via LDAP/AD.`n" +
      "- activated: Set to TRUE to make the account active, FALSE for inactive.`n" +
      "- manager_id: The ID of the user's manager (another user ID in Snipe-IT)."
            
      $referenceGuide += "`n===== End of Reference Guide ====="
            
      # Save the reference guide
      $refPath = [System.IO.Path]::ChangeExtension($savePath, ".reference.txt")
      $referenceGuide | Out-File -FilePath $refPath -Encoding UTF8
      Write-Host "Reference guide saved to: $refPath" -ForegroundColor Green
            
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
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
  }
    
  Write-Host "`nPress any key to continue..."
  $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to generate a License Import Template
function New-LicenseImportTemplate
{
  Clear-Host
  Write-Host "===== Generate License Import Template =====" -ForegroundColor Cyan
    
  try
  {
    # Get the save location
    $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\snipeit_license_template.csv"
    Write-Host "This will create a template CSV file for importing licenses into Snipe-IT." -ForegroundColor Yellow
    $savePath = Read-Host "Enter a path to save the template (default: $defaultPath)"
        
    if ([string]::IsNullOrWhiteSpace($savePath))
    {
      $savePath = $defaultPath
    }

    # Check if connection is initialized
    if (-not (Initialize-SnipeITConnection))
    {
      return
    }
        
    # Gather available field options from Snipe-IT for user reference
    Write-Host "`nRetrieving reference data from Snipe-IT..." -ForegroundColor Green
        
    # Get categories (license)
    $categories = Get-SnipeitCategory -category_type 'license'
    $categoryOptions = if ($categories)
    {
      $categories | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get manufacturers
    $manufacturers = Get-SnipeitManufacturer
    $manufacturerOptions = if ($manufacturers)
    {
      $manufacturers | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get suppliers
    $suppliers = Get-SnipeitSupplier
    $supplierOptions = if ($suppliers)
    {
      $suppliers | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Get custom fields
    $customFields = Get-SnipeitCustomField
    $customFieldOptions = if ($customFields)
    {
      $customFields | Where-Object { $_.field_type -eq "license" } | Select-Object -Property id, name, format, db_column_name | Sort-Object -Property db_column_name
    } else
    {
      @()
    }
        
    # Create the template content
    $templateHeaders = @(
      "name",
      "company",
      "product_key",
      "order_number",
      "seats",
      "license_name",
      "license_email",
      "reassignable",
      "category_id",
      "manufacturer_id",
      "supplier_id",
      "expiration_date",
      "purchase_date",
      "purchase_cost",
      "purchase_order",
      "notes",
      "maintained"
    )
        
    # Add custom field headers
    $customFieldHeaders = @()
    if ($customFieldOptions.Count -gt 0)
    {
      foreach ($field in $customFieldOptions)
      {
        if (-not [string]::IsNullOrWhiteSpace($field.db_column_name))
        {
          $templateHeaders += "_snipeit_$($field.db_column_name)"
          $customFieldHeaders += "_snipeit_$($field.db_column_name)"
        }
      }
    }
        
    # Create header row
    $templateContent = $templateHeaders -join ","
    $templateContent += "`n"
        
    # Create example rows
    $exampleRow1 = @(
      '"Microsoft Office 365"',
      '"Example Company"',
      '"XXXX-XXXX-XXXX-XXXX"',  # product_key
      '"ORD-12345"',
      '50',  # seats
      '"IT Department"',  # license_name
      '"it@example.com"',  # license_email
      'TRUE',  # reassignable
      '1',  # category_id
      '1',  # manufacturer_id (Microsoft)
      '1',  # supplier_id
      '"2024-12-31"',  # expiration_date
      '"2023-01-15"',  # purchase_date
      '"10000.00"',  # purchase_cost
      '"PO-12345"',  # purchase_order
      '"Enterprise agreement for Office 365"',  # notes
      'TRUE'  # maintained
    )
        
    # Add custom field examples
    foreach ($header in $customFieldHeaders)
    {
      $exampleRow1 += '""'  # Empty custom field values
    }
        
    $exampleRow2 = @(
      '"Adobe Creative Cloud"',
      '"Example Company"',
      '"AAAA-BBBB-CCCC-DDDD"',  # product_key
      '"ORD-67890"',
      '25',  # seats
      '"Design Team"',  # license_name
      '"design@example.com"',  # license_email
      'TRUE',  # reassignable
      '2',  # category_id
      '2',  # manufacturer_id (Adobe)
      '1',  # supplier_id
      '"2024-06-30"',  # expiration_date
      '"2023-02-20"',  # purchase_date
      '"15000.00"',  # purchase_cost
      '"PO-67890"',  # purchase_order
      '"Creative Cloud licenses for design team"',  # notes
      'TRUE'  # maintained
    )
        
    # Add custom field examples
    foreach ($header in $customFieldHeaders)
    {
      $exampleRow2 += '""'  # Empty custom field values
    }
        
    # Add example rows to template
    $templateContent += $exampleRow1 -join ","
    $templateContent += "`n"
    $templateContent += $exampleRow2 -join ","
        
    # Save the template
    try
    {
      $templateContent | Out-File -FilePath $savePath -Encoding UTF8
      Write-Host "`nTemplate saved successfully to: $savePath" -ForegroundColor Green
            
      # Create reference guide for IDs
      $referenceGuide = "`n===== Reference Guide for Snipe-IT IDs =====" +
      "`n`nCategories (category_id):`n"
            
      if ($categoryOptions.Count -gt 0)
      {
        $referenceGuide += $categoryOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No license categories found in your Snipe-IT instance.`n"
      }
            
      $referenceGuide += "`nManufacturers (manufacturer_id):`n"
      if ($manufacturerOptions.Count -gt 0)
      {
        $referenceGuide += $manufacturerOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No manufacturers found in your Snipe-IT instance.`n"
      }
            
      $referenceGuide += "`nSuppliers (supplier_id):`n"
      if ($supplierOptions.Count -gt 0)
      {
        $referenceGuide += $supplierOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No suppliers found in your Snipe-IT instance.`n"
      }
            
      if ($customFieldOptions.Count -gt 0)
      {
        $referenceGuide += "`nCustom Fields:`n"
        $referenceGuide += $customFieldOptions | ForEach-Object { 
          "  Name: $($_.name)`n  Column: _snipeit_$($_.db_column_name)`n  Format: $($_.format)`n" 
        } | Out-String
      }
            
      $referenceGuide += "`n===== Notes =====`n" +
      "- reassignable: Set to TRUE if the license can be reassigned after checkout, FALSE if not.`n" +
      "- maintained: Set to TRUE if the license is under maintenance/support, FALSE if not.`n" +
      "- Date fields should be in YYYY-MM-DD format."
            
      $referenceGuide += "`n===== End of Reference Guide ====="
            
      # Save the reference guide
      $refPath = [System.IO.Path]::ChangeExtension($savePath, ".reference.txt")
      $referenceGuide | Out-File -FilePath $refPath -Encoding UTF8
      Write-Host "Reference guide saved to: $refPath" -ForegroundColor Green
            
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
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
  }
    
  Write-Host "`nPress any key to continue..."
  $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to generate a Maintenance Import Template
function New-MaintenanceImportTemplate
{
  Clear-Host
  Write-Host "===== Generate Maintenance Import Template =====" -ForegroundColor Cyan
    
  try
  {
    # Get the save location
    $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\snipeit_maintenance_template.csv"
    Write-Host "This will create a template CSV file for importing maintenance records into Snipe-IT." -ForegroundColor Yellow
    $savePath = Read-Host "Enter a path to save the template (default: $defaultPath)"
        
    if ([string]::IsNullOrWhiteSpace($savePath))
    {
      $savePath = $defaultPath
    }

    # Check if connection is initialized
    if (-not (Initialize-SnipeITConnection))
    {
      return
    }
        
    # Gather available field options from Snipe-IT for user reference
    Write-Host "`nRetrieving reference data from Snipe-IT..." -ForegroundColor Green
        
    # Get assets
    $assets = Get-SnipeitAsset -limit 100  # Limit to prevent too much data
    $assetOptions = if ($assets)
    {
      $assets | Select-Object -Property id, asset_tag, name | Sort-Object -Property asset_tag
    } else
    {
      @()
    }
        
    # Get suppliers
    $suppliers = Get-SnipeitSupplier
    $supplierOptions = if ($suppliers)
    {
      $suppliers | Select-Object -Property id, name | Sort-Object -Property name
    } else
    {
      @()
    }
        
    # Create the template content
    $templateHeaders = @(
      "asset_id",
      "asset_tag",
      "supplier_id",
      "supplier_name",
      "maintenance_type",
      "title",
      "start_date",
      "completion_date",
      "cost",
      "notes"
    )
        
    # Create header row
    $templateContent = $templateHeaders -join ","
    $templateContent += "`n"
        
    # Create example rows
    $exampleRow1 = @(
      '1',  # asset_id - ID of the asset in Snipe-IT (use either asset_id or asset_tag)
      '',   # asset_tag - leave blank if using asset_id
      '1',  # supplier_id
      '',   # supplier_name - leave blank if using supplier_id
      '"maintenance"',  # maintenance_type
      '"Annual Preventive Maintenance"',  # title
      '"2023-01-15"',  # start_date
      '"2023-01-16"',  # completion_date
      '"250.00"',  # cost
      '"Regular annual maintenance check"'  # notes
    )
        
    $exampleRow2 = @(
      '',   # asset_id - leave blank if using asset_tag
      '"A00002"',  # asset_tag - tag of the asset in Snipe-IT
      '',   # supplier_id - leave blank if using supplier_name
      '"XYZ IT Services"',  # supplier_name
      '"repair"',  # maintenance_type
      '"Screen Replacement"',  # title
      '"2023-02-20"',  # start_date
      '"2023-02-22"',  # completion_date
      '"150.00"',  # cost
      '"Replaced cracked screen"'  # notes
    )
        
    # Add example rows to template
    $templateContent += $exampleRow1 -join ","
    $templateContent += "`n"
    $templateContent += $exampleRow2 -join ","
        
    # Save the template
    try
    {
      $templateContent | Out-File -FilePath $savePath -Encoding UTF8
      Write-Host "`nTemplate saved successfully to: $savePath" -ForegroundColor Green
            
      # Create reference guide for IDs
      $referenceGuide = "`n===== Reference Guide for Snipe-IT Maintenance Import =====" +
      "`n`nMaintenance Types:`n" +
      "  - maintenance: Regular scheduled maintenance\n" +
      "  - repair: Fixing a damaged/broken item\n" +
      "  - upgrade: Hardware or software upgrades\n" +
      "  - calibration: Accuracy adjustment for devices\n" +
      "  - software_support: Software maintenance/support\n" +
      "  - hardware_support: Hardware service/support\n"
            
      if ($assetOptions.Count -gt 0)
      {
        $referenceGuide += "`nSample Assets (select 10 shown):`n"
        $referenceGuide += $assetOptions | Select-Object -First 10 | ForEach-Object { 
          "  ID: $($_.id) | Tag: $($_.asset_tag) | Name: $($_.name)" 
        } | Out-String
        $referenceGuide += "  ... (more assets available in your Snipe-IT instance)\n"
      } else
      {
        $referenceGuide += "  No assets found in your Snipe-IT instance.\n"
      }
            
      $referenceGuide += "`nSuppliers (supplier_id):`n"
      if ($supplierOptions.Count -gt 0)
      {
        $referenceGuide += $supplierOptions | ForEach-Object { "  $($_.id): $($_.name)" } | Out-String
      } else
      {
        $referenceGuide += "  No suppliers found in your Snipe-IT instance.\n"
      }
            
      $referenceGuide += "`n===== Notes =====`n" +
      "- Use either asset_id OR asset_tag (not both) to identify the asset\n" +
      "- Use either supplier_id OR supplier_name (not both) to identify the supplier\n" +
      "- Date fields should be in YYYY-MM-DD format\n" +
      "- Completion date can be left blank for ongoing maintenance\n" +
      "- Cost should be a decimal number (e.g., 150.00)\n"
            
      $referenceGuide += "`n===== End of Reference Guide ====="
            
      # Save the reference guide
      $refPath = [System.IO.Path]::ChangeExtension($savePath, ".reference.txt")
      $referenceGuide | Out-File -FilePath $refPath -Encoding UTF8
      Write-Host "Reference guide saved to: $refPath" -ForegroundColor Green
            
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
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
  }
    
  Write-Host "`nPress any key to continue..."
  $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to view template documentation
function Show-TemplateDocumentation
{
  Clear-Host
  Write-Host "===== Snipe-IT CSV Import Documentation =====" -ForegroundColor Cyan
    
  $documentation = @"
## Snipe-IT CSV Import Guide

This documentation provides guidance on preparing and importing CSV data into Snipe-IT using the templates generated by this tool.

### General Import Guidelines

1. **File Format**:
   - Files must be in CSV format with UTF-8 encoding
   - First row must contain column headers
   - Text containing commas should be enclosed in double quotes
   - Dates should be in YYYY-MM-DD format

2. **Required vs Optional Fields**:
   - Each template includes both required and optional fields
   - Required fields must have values for successful import
   - Optional fields can be left blank

3. **ID Fields**:
   - Many fields use numeric IDs (e.g., model_id, category_id)
   - The reference guide included with each template lists available IDs
   - For some fields, you can use names instead of IDs (e.g., supplier_name vs supplier_id)

4. **Custom Fields**:
   - Custom fields in Snipe-IT are prefixed with "_snipeit_" in the CSV
   - Format varies based on custom field type (text, date, etc.)

### Asset Import Details

The asset template includes fields for creating new assets in Snipe-IT.

**Required Fields**:
- asset_tag: A unique identifier for the asset
- model_id: The ID of the model this asset belongs to
- status_id: The ID of the status for this asset (e.g., Ready to Deploy)

**Common Optional Fields**:
- name: A descriptive name for the asset
- serial: The serial number of the asset
- purchase_date: When the asset was purchased (YYYY-MM-DD)
- purchase_cost: The cost of the asset
- supplier_id: The ID of the supplier
- location_id: The ID of the default/home location

### User Import Details

The user template includes fields for creating new user accounts in Snipe-IT.

**Required Fields**:
- first_name: The user's first name
- last_name: The user's last name
- username: A unique username for login
- email: The user's email address

**Common Optional Fields**:
- password: User's password (if blank, a random one will be generated)
- department_id: The ID of the user's department
- location_id: The ID of the user's location
- manager_id: The ID of the user's manager
- activated: TRUE/FALSE whether the account is active

### License Import Details

The license template includes fields for creating software licenses in Snipe-IT.

**Required Fields**:
- name: A descriptive name for the license
- seats: The number of license seats available
- category_id: The ID of the license category

**Common Optional Fields**:
- product_key: The license/product key
- expiration_date: When the license expires (YYYY-MM-DD)
- purchase_date: When the license was purchased (YYYY-MM-DD)
- purchase_cost: The cost of the license
- manufacturer_id: The ID of the manufacturer
- supplier_id: The ID of the supplier

### Maintenance Import Details

The maintenance template includes fields for creating maintenance records in Snipe-IT.

**Required Fields**:
- Asset Identifier (either asset_id OR asset_tag)
- maintenance_type: Type of maintenance performed
- title: A descriptive title for the maintenance record
- start_date: When the maintenance started (YYYY-MM-DD)

**Common Optional Fields**:
- completion_date: When maintenance was completed (YYYY-MM-DD)
- supplier_id or supplier_name: Who performed the maintenance
- cost: The cost of the maintenance
- notes: Additional details about the maintenance

### Import Process in Snipe-IT

1. Log in to your Snipe-IT web interface
2. Navigate to the appropriate section (Assets, Users, etc.)
3. Look for the Import option (usually a button labeled "Import")
4. Upload your CSV file
5. Follow the on-screen mapping instructions
6. Preview the data and confirm the import

### Common Errors and Solutions

- **Duplicate asset tags**: Asset tags must be unique
- **Invalid model/category/status IDs**: Verify these exist in your Snipe-IT instance
- **Date format issues**: Ensure dates are in YYYY-MM-DD format
- **Character encoding problems**: Save your CSV with UTF-8 encoding
- **Missing required fields**: Ensure all required fields have values
"@
    
  # Display the documentation
  Write-Host $documentation
    
  # Option to save the documentation
  $saveDoc = Read-Host "`nWould you like to save this documentation to a file? (Y/N)"
    
  if ($saveDoc -eq "Y" -or $saveDoc -eq "y")
  {
    $defaultPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\SnipeIT_Import_Documentation.md"
    $savePath = Read-Host "Enter path to save documentation (default: $defaultPath)"
        
    if ([string]::IsNullOrWhiteSpace($savePath))
    {
      $savePath = $defaultPath
    }
        
    try
    {
      $documentation | Out-File -FilePath $savePath -Encoding UTF8
      Write-Host "Documentation saved to: $savePath" -ForegroundColor Green
            
      # Option to open the file
      $openFile = Read-Host "Would you like to open the documentation file? (Y/N)"
            
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
      Write-Host "Error saving documentation: $_" -ForegroundColor Red
    }
  }
    
  Write-Host "`nPress any key to continue..."
  $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Check if the script is being run standalone or imported
if ($MyInvocation.InvocationName -ne '.')
{
  # Script is being run directly, show a menu
  function Show-TemplatesMenu
  {
    Clear-Host
    Write-Host "===== Snipe-IT CSV Templates =====" -ForegroundColor Cyan
    Write-Host "1. Generate Asset Import Template" -ForegroundColor Green
    Write-Host "2. Generate User Import Template" -ForegroundColor Green
    Write-Host "3. Generate License Import Template" -ForegroundColor Green
    Write-Host "4. Generate Maintenance Import Template" -ForegroundColor Green
    Write-Host "5. View Template Documentation" -ForegroundColor Green
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host "=============================" -ForegroundColor Cyan
        
    $choice = Read-Host "Select an option"
        
    switch ($choice)
    {
      "1"
      { New-AssetImportTemplate 
      }
      "2"
      { New-UserImportTemplate 
      }
      "3"
      { New-LicenseImportTemplate 
      }
      "4"
      { New-MaintenanceImportTemplate 
      }
      "5"
      { Show-TemplateDocumentation 
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
        Show-TemplatesMenu
      }
    }
        
    # Return to menu after function completes
    Show-TemplatesMenu
  }
    
  # If API credentials were passed, set them for the script session
  if (-not [string]::IsNullOrEmpty($APIKey) -and -not [string]::IsNullOrEmpty($SnipeURL))
  {
    $Script:APIKey = $APIKey
    $Script:SnipeURL = $SnipeURL
  }
    
  # Start the standalone menu
  Show-TemplatesMenu
}

# Export functions for importing
Export-ModuleMember -Function New-AssetImportTemplate, New-UserImportTemplate, New-LicenseImportTemplate, New-MaintenanceImportTemplate, Show-TemplateDocumentation
