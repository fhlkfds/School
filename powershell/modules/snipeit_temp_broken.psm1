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

function Show-TempBrokenMenu
{
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

function Get-UserByEmailOrEmployeeNum
{
  param(
    [Parameter(Mandatory=$true)]
    [string]$UserIdentifier
  )
    
  try
  {
    # Search by email or employee number
    $users = Get-SnipeitUser -search $UserIdentifier
        
    if ($users)
    {
      # Handle case where search returns an array of users
      # Convert to array to ensure we can work with it consistently
      $userArray = @($users)
            
      # First, try to find exact matches
      foreach ($user in $userArray)
      {
        # Check if it matches email exactly
        if ($user.email -eq $UserIdentifier)
        {
          Write-Host "Found exact email match: $($user.name) ($($user.email))" -ForegroundColor Green
          return $user
        }
        # Check if it matches employee number exactly
        if ($user.employee_num -eq $UserIdentifier)
        {
          Write-Host "Found exact employee number match: $($user.name) (ID: $($user.employee_num))" -ForegroundColor Green
          return $user
        }
      }
            
      # If no exact match found, but we have results, use the first one with confirmation
      if ($userArray.Count -eq 1)
      {
        $user = $userArray[0]
        Write-Host "Found user: $($user.name) ($($user.email)) [Employee #: $($user.employee_num)]" -ForegroundColor Yellow
        Write-Host "Please verify this is the correct user." -ForegroundColor Yellow
        return $user
      } elseif ($userArray.Count -gt 1)
      {
        Write-Host "Multiple users found matching '$UserIdentifier':" -ForegroundColor Yellow
        for ($i = 0; $i -lt $userArray.Count; $i++)
        {
          $user = $userArray[$i]
          Write-Host "  $($i + 1). $($user.name) ($($user.email)) [Employee #: $($user.employee_num)]" -ForegroundColor White
        }
                
        do
        {
          $selection = Read-Host "Please select the correct user (1-$($userArray.Count)) or 0 to cancel"
          if ($selection -eq "0")
          {
            Write-Host "User selection cancelled." -ForegroundColor Yellow
            return $null
          }
          $selectedIndex = [int]$selection - 1
        } while ($selectedIndex -lt 0 -or $selectedIndex -ge $userArray.Count)
                
        return $userArray[$selectedIndex]
      }
    }
        
    Write-Host "User not found with email or ID number: $UserIdentifier" -ForegroundColor Red
    return $null
  } catch
  {
    Write-Host "Error retrieving user with identifier $UserIdentifier : $($_.Exception.Message)" -ForegroundColor Red
    return $null
  }
}
function Get-AssetByTag
{
  param(
    [Parameter(Mandatory=$true)]
    [string]$AssetTag
  )
    
  try
  {
    $asset = Get-SnipeitAsset -asset_tag $AssetTag
    return $asset
  } catch
  {
    Write-Host "Error retrieving asset with tag $AssetTag : $($_.Exception.Message)" -ForegroundColor Red
    return $null
  }
}

function AssignTempLaptop
{
  Write-Host "`n===========================================" -ForegroundColor Green
  Write-Host "        ASSIGN TEMPORARY LAPTOP            " -ForegroundColor Green
  Write-Host "===========================================" -ForegroundColor Green
    
  # Get User identifier (email or employee ID number)
  $userIdentifier = Read-Host "`nEnter User Email or ID Number (Employee Number)"
  if ([string]::IsNullOrWhiteSpace($userIdentifier))
  {
    Write-Host "User email or ID number cannot be empty." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Validate user exists using enhanced lookup
  $user = Get-UserByEmailOrEmployeeNum -UserIdentifier $userIdentifier
  if (-not $user)
  {
    Start-Sleep -Seconds 3
    return
  }
    
  Write-Host "User found: $($user.name) ($($user.email)) [Employee #: $($user.employee_num)]" -ForegroundColor Green
    
  # Get Asset Tag
  $assetTag = Read-Host "Enter Temporary Laptop Asset Tag"
  if ([string]::IsNullOrWhiteSpace($assetTag))
  {
    Write-Host "Asset tag cannot be empty." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Validate asset exists and is available
  $asset = Get-AssetByTag -AssetTag $assetTag
  if (-not $asset)
  {
    Write-Host "Asset with tag $assetTag not found." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Check if asset is available
  if ($asset.assigned_to)
  {
    Write-Host "Asset $assetTag is already assigned to someone else." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  Write-Host "Asset found: $($asset.name) - $($asset.model)" -ForegroundColor Green
    
  # Set expected return date to current date automatically
  $dueDate = Get-Date -Format "yyyy-MM-dd"
  Write-Host "Expected return date automatically set to: $dueDate" -ForegroundColor Cyan
    
  try
  {
    # Set checkout date to current date and time
    $checkoutDate = Get-Date
        
    # Checkout asset to user using the correct syntax
    $checkoutParams = @{
      id = $asset.id
      assigned_id = $user.id
      checkout_to_type = "user"
      note = "TEMPORARY ASSIGNMENT"
      checkout_at = $checkoutDate
      expected_checkin = [DateTime]::Parse($dueDate)
    }
        
    $result = Set-SnipeitAssetOwner @checkoutParams
        
    if ($result)
    {
      Write-Host "`nSuccess! Temporary laptop $assetTag assigned to user $($user.name)" -ForegroundColor Green
      Write-Host "Checkout Date: $($checkoutDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Green
      Write-Host "Expected Return: $dueDate" -ForegroundColor Green
    } else
    {
      Write-Host "Failed to assign temporary laptop." -ForegroundColor Red
    }
  } catch
  {
    Write-Host "Error assigning temporary laptop: $($_.Exception.Message)" -ForegroundColor Red
  }
    
  Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
  [void][System.Console]::ReadKey($true)
}




# Add this function to your module (after the existing functions, before Export-ModuleMember)
function Get-BrokenStatusId
{
  try
  {
    # First try to find the "Broken" status directly
    $brokenStatus = Get-SnipeitStatus -search "Broken"
    if ($brokenStatus) 
    {
      Write-Host "Found broken status: '$($brokenStatus.name)' (ID: $($brokenStatus.id))" -ForegroundColor Green
      return $brokenStatus.id
    }
        
    # If not found, get all statuses and look for keywords
    $allStatuses = Get-SnipeitStatus
        
    # Look for common broken status names
    $brokenKeywords = @("broken", "defective", "repair", "damaged", "faulty", "undeployable")
        
    foreach ($status in $allStatuses)
    {
      foreach ($keyword in $brokenKeywords)
      {
        if ($status.name -like "*$keyword*" -or $status.type -eq "undeployable")
        {
          Write-Host "Found potential broken status: '$($status.name)' (ID: $($status.id), Type: $($status.type))" -ForegroundColor Yellow
          return $status.id
        }
      }
    }
        
    # If no broken status found, display all statuses for user to choose
    Write-Host "Could not automatically find a 'broken' status. Available statuses:" -ForegroundColor Yellow
    foreach ($status in $allStatuses)
    {
      Write-Host "  ID: $($status.id) - Name: $($status.name) - Type: $($status.type)" -ForegroundColor White
    }
        
    # Return null to indicate manual selection needed
    return $null
  } catch
  {
    Write-Host "Error retrieving status labels: $($_.Exception.Message)" -ForegroundColor Red
    # Since we know ID 6 exists for Broken, return it as fallback
    Write-Host "Using fallback broken status ID: 6" -ForegroundColor Yellow
    return 6
  }
}

# Simplified version that just uses the known broken status ID
function ReportBrokenLaptop
{
  Write-Host "`n===========================================" -ForegroundColor Yellow
  Write-Host "      REPORT BROKEN LAPTOP                " -ForegroundColor Yellow
  Write-Host "===========================================" -ForegroundColor Yellow
    
  # Get User identifier (email or employee ID number)
  $userIdentifier = Read-Host "`nEnter User Email or ID Number (Employee Number)"
  if ([string]::IsNullOrWhiteSpace($userIdentifier))
  {
    Write-Host "User email or ID number cannot be empty." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Validate user exists using enhanced lookup
  $user = Get-UserByEmailOrEmployeeNum -UserIdentifier $userIdentifier
  if (-not $user)
  {
    Start-Sleep -Seconds 3
    return
  }
    
  Write-Host "User found: $($user.name) ($($user.email)) [Employee #: $($user.employee_num)]" -ForegroundColor Green
    
  # Get broken laptop asset tag
  $brokenAssetTag = Read-Host "Enter Broken Laptop Asset Tag"
  if ([string]::IsNullOrWhiteSpace($brokenAssetTag))
  {
    Write-Host "Broken laptop asset tag cannot be empty." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Validate broken asset exists
  $brokenAsset = Get-AssetByTag -AssetTag $brokenAssetTag
  if (-not $brokenAsset)
  {
    Write-Host "Broken asset with tag $brokenAssetTag not found." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Display asset info more safely
  $assetName = if ($brokenAsset.name)
  { $brokenAsset.name 
  } else
  { "N/A" 
  }
  $assetModel = if ($brokenAsset.model -and $brokenAsset.model.name)
  { $brokenAsset.model.name 
  } elseif ($brokenAsset.model)
  { $brokenAsset.model 
  } else
  { "N/A" 
  }
  Write-Host "Broken Asset: $assetName - $assetModel" -ForegroundColor Yellow
    
  # Check if the broken asset is currently assigned
  if (-not $brokenAsset.assigned_to)
  {
    Write-Host "Warning: This asset is not currently assigned to anyone." -ForegroundColor Yellow
    $confirm = Read-Host "Continue anyway? The asset will be marked as broken. (y/N)"
    if ($confirm -ne 'y' -and $confirm -ne 'Y')
    {
      return
    }
  } else
  {
    Write-Host "Currently assigned to: $($brokenAsset.assigned_to.name)" -ForegroundColor Cyan
  }
    
  # Get replacement laptop asset tag
  $newAssetTag = Read-Host "Enter Replacement Laptop Asset Tag"
  if ([string]::IsNullOrWhiteSpace($newAssetTag))
  {
    Write-Host "Replacement laptop asset tag cannot be empty." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Validate replacement asset exists and is available
  $newAsset = Get-AssetByTag -AssetTag $newAssetTag
  if (-not $newAsset)
  {
    Write-Host "Replacement asset with tag $newAssetTag not found." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  if ($newAsset.assigned_to)
  {
    Write-Host "Replacement asset $newAssetTag is already assigned to someone else." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Display replacement asset info more safely
  $newAssetName = if ($newAsset.name)
  { $newAsset.name 
  } else
  { "N/A" 
  }
  $newAssetModel = if ($newAsset.model -and $newAsset.model.name)
  { $newAsset.model.name 
  } elseif ($newAsset.model)
  { $newAsset.model 
  } else
  { "N/A" 
  }
  Write-Host "Replacement Asset: $newAssetName - $newAssetModel" -ForegroundColor Green
    
  # Get issue description
  $issueDescription = Read-Host "Describe what is broken"
  if ([string]::IsNullOrWhiteSpace($issueDescription))
  {
    Write-Host "Issue description cannot be empty." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Get supplier ID for maintenance record
  $supplierIdInput = Read-Host "Enter Supplier ID for maintenance record (or press Enter for default: 1)"
  $supplierId = if ([string]::IsNullOrWhiteSpace($supplierIdInput))
  { 1 
  } else
  { [int]$supplierIdInput 
  }
    
  # Use the known broken status ID (6)
  $brokenStatusId = 6
    
  try
  {
    # Step 1: Process the broken laptop
    Write-Host "Processing broken laptop..." -ForegroundColor Yellow
        
    if ($brokenAsset.assigned_to)
    {
      # Asset is assigned, so check it in first
      Write-Host "Checking in broken laptop from user..." -ForegroundColor Yellow
      $checkinResult = Reset-SnipeitAssetOwner -id $brokenAsset.id -note "BROKEN - $issueDescription - Replaced with $newAssetTag"
            
      if ($checkinResult)
      {
        Write-Host "✓ Broken laptop checked in successfully" -ForegroundColor Green
                
        # Now update the status to broken
        Write-Host "Setting status to broken..." -ForegroundColor Yellow
        try
        {
          $statusParams = @{
            id = $brokenAsset.id
            status_id = $brokenStatusId
          }
          $statusResult = Set-SnipeitAsset @statusParams
          if ($statusResult)
          {
            Write-Host "✓ Status set to broken (ID: $brokenStatusId)" -ForegroundColor Green
          } else
          {
            Write-Host "⚠ Warning: Checked in successfully but failed to set broken status" -ForegroundColor Yellow
          }
        } catch
        {
          Write-Host "⚠ Warning: Error setting broken status: $($_.Exception.Message)" -ForegroundColor Yellow
        }
      } else
      {
        Write-Host "❌ Failed to check in broken laptop. Trying alternative method..." -ForegroundColor Yellow
                
        # Try updating the asset directly to set status and notes
        try
        {
          $updateParams = @{
            id = $brokenAsset.id
            status_id = $brokenStatusId
            notes = "BROKEN - $issueDescription - Replaced with $newAssetTag"
          }
          $updateResult = Set-SnipeitAsset @updateParams
                    
          if ($updateResult)
          {
            Write-Host "✓ Broken laptop status updated successfully" -ForegroundColor Green
            $checkinResult = $true
          } else
          {
            throw "Could not update broken laptop status"
          }
        } catch
        {
          Write-Host "❌ All methods failed: $($_.Exception.Message)" -ForegroundColor Red
          throw "Could not process broken laptop"
        }
      }
    } else
    {
      # Asset is not assigned, just update its status
      Write-Host "Asset not assigned, updating status to broken..." -ForegroundColor Yellow
      $updateParams = @{
        id = $brokenAsset.id
        status_id = $brokenStatusId
        notes = "BROKEN - $issueDescription - Replaced with $newAssetTag"
      }
      $checkinResult = Set-SnipeitAsset @updateParams
            
      if (-not $checkinResult)
      {
        Write-Host "❌ Failed to update broken laptop status" -ForegroundColor Red
        throw "Could not mark asset as broken"
      } else
      {
        Write-Host "✓ Broken laptop status updated successfully" -ForegroundColor Green
      }
    }
        
    if ($checkinResult)
    {
      # Step 2: Create maintenance record
      Write-Host "Creating maintenance record..." -ForegroundColor Yellow
      try
      {
        $maintenanceParams = @{
          asset_id = $brokenAsset.id
          supplier_id = $supplierId
          asset_maintenance_type = "Repair"
          title = "Laptop Repair - $issueDescription"
          start_date = Get-Date
          notes = "Broken laptop reported by $($user.name). Issue: $issueDescription. Replaced with asset $newAssetTag."
        }
                
        $maintenanceResult = New-SnipeitAssetMaintenance @maintenanceParams
                
        if ($maintenanceResult)
        {
          Write-Host "✓ Maintenance record created successfully" -ForegroundColor Green
        } else
        {
          Write-Host "⚠ Warning: Failed to create maintenance record" -ForegroundColor Yellow
        }
      } catch
      {
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
            
      if ($checkoutResult)
      {
        Write-Host "`n🎉 SUCCESS! Replacement process completed:" -ForegroundColor Green
        Write-Host "├─ Broken laptop ${brokenAssetTag}: Processed and marked as broken" -ForegroundColor Green
        Write-Host "├─ Maintenance record: Created for repair tracking" -ForegroundColor Green
        Write-Host "├─ Replacement laptop ${newAssetTag}: Assigned to $($user.name)" -ForegroundColor Green
        Write-Host "├─ Issue description: $issueDescription" -ForegroundColor Green
        Write-Host "└─ Checkout date: $($checkoutDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Green
      } else
      {
        Write-Host "⚠ Warning: Broken laptop processed, but failed to assign replacement" -ForegroundColor Yellow
        Write-Host "You may need to manually assign the replacement laptop $newAssetTag to $($user.name)" -ForegroundColor Yellow
      }
    }
  } catch
  {
    Write-Host "❌ Error processing broken laptop: $($_.Exception.Message)" -ForegroundColor Red
  }
    
  Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
  [void][System.Console]::ReadKey($true)
}

function ViewBrokenLaptopStats
{
  Write-Host "`n===========================================" -ForegroundColor Magenta
  Write-Host "      BROKEN LAPTOP STATISTICS            " -ForegroundColor Magenta
  Write-Host "===========================================" -ForegroundColor Magenta
    
  try
  {
    # Use the known broken status ID (6)
    $brokenStatusId = 6
        
    # Get all assets with the broken status
    $brokenAssets = Get-SnipeitAsset -status_id $brokenStatusId
        
    if ($brokenAssets -and $brokenAssets.Count -gt 0)
    {
      Write-Host "`nTotal Broken Laptops: $($brokenAssets.Count)" -ForegroundColor Red
      Write-Host "`nBroken Assets Details:" -ForegroundColor White
      Write-Host "=" * 80 -ForegroundColor Gray
            
      foreach ($asset in $brokenAssets)
      {
        Write-Host "Asset Tag: $($asset.asset_tag)" -ForegroundColor Yellow
        Write-Host "Name: $($asset.name)" -ForegroundColor White
                
        # Handle model display safely
        $modelName = if ($asset.model -and $asset.model.name)
        { $asset.model.name 
        } elseif ($asset.model)
        { $asset.model 
        } else
        { "N/A" 
        }
        Write-Host "Model: $modelName" -ForegroundColor White
                
        Write-Host "Serial: $($asset.serial)" -ForegroundColor White
                
        # Handle status display safely
        $statusName = if ($asset.status_label -and $asset.status_label.name)
        { $asset.status_label.name 
        } elseif ($asset.status_label)
        { $asset.status_label 
        } else
        { "Broken" 
        }
        Write-Host "Status: $statusName" -ForegroundColor Red
                
        # Get asset history to find latest notes
        try
        {
          $assetDetails = Get-SnipeitAsset -id $asset.id
          if ($assetDetails.notes)
          {
            Write-Host "Notes: $($assetDetails.notes)" -ForegroundColor Cyan
          }
        } catch
        {
          Write-Host "Could not retrieve detailed notes for this asset." -ForegroundColor Gray
        }
                
        Write-Host "-" * 80 -ForegroundColor Gray
      }
    } else
    {
      Write-Host "`nNo broken laptops found in the system." -ForegroundColor Green
    }
        
    # Additional statistics
    Write-Host "`nSummary Statistics:" -ForegroundColor White
    Write-Host "Status ID used for 'broken': $brokenStatusId" -ForegroundColor Gray
    Write-Host "Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Gray
  } catch
  {
    Write-Host "Error retrieving broken laptop statistics: $($_.Exception.Message)" -ForegroundColor Red
  }
    
  Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
  [void][System.Console]::ReadKey($true)
}


function ViewTodayTempAssignments
{
  Write-Host "`n===========================================" -ForegroundColor Blue
  Write-Host "    TODAY'S TEMP LAPTOP ASSIGNMENTS       " -ForegroundColor Blue
  Write-Host "===========================================" -ForegroundColor Blue
    
  try
  {
    $today = Get-Date -Format "yyyy-MM-dd"
    Write-Host "Searching for temp assignments made on: $today" -ForegroundColor Cyan
        
    # Get all assets (without status filter to cast a wider net)
    $allAssets = Get-SnipeitAsset
        
    # Convert to array if single result
    $allAssets = @($allAssets)
        
    # Filter for temp assignments made today
    $tempAssignments = @()
        
    foreach ($asset in $allAssets)
    {
      # Skip if asset is not assigned
      if (-not $asset.assigned_to)
      { continue 
      }
            
      try
      {
        # Get detailed asset information
        $assetDetails = Get-SnipeitAsset -id $asset.id
                
        # Parse checkout date - handle both object and string formats
        $checkoutDate = $null
        $checkoutDateString = $null
                
        if ($assetDetails.last_checkout)
        {
          # Check if it's a hashtable/object with datetime property
          if ($assetDetails.last_checkout -is [hashtable] -or $assetDetails.last_checkout.datetime)
          {
            if ($assetDetails.last_checkout.datetime)
            {
              $checkoutDate = [DateTime]::Parse($assetDetails.last_checkout.datetime)
              $checkoutDateString = $checkoutDate.ToString("yyyy-MM-dd")
            }
          }
          # Check if it's a simple string
          elseif ($assetDetails.last_checkout -is [string])
          {
            try
            {
              $checkoutDate = [DateTime]::Parse($assetDetails.last_checkout)
              $checkoutDateString = $checkoutDate.ToString("yyyy-MM-dd")
            } catch
            {
              # If parsing fails, try to extract the date part
              if ($assetDetails.last_checkout.Length -ge 10)
              {
                $checkoutDateString = $assetDetails.last_checkout.Substring(0, 10)
              }
            }
          }
        }
                
        # Check if checked out today
        if ($checkoutDateString -eq $today)
        {
          # Check for temp assignment indicators in multiple places
          $isTemp = $false
                    
          # Check asset notes
          if ($assetDetails.notes -like "*TEMPORARY ASSIGNMENT*")
          {
            $isTemp = $true
          }
                    
          # Get asset activity/history to check checkout notes
          try
          {
            $assetActivity = Get-SnipeitActivity -item_id $asset.id -item_type "asset"
            if ($assetActivity)
            {
              foreach ($activity in $assetActivity)
              {
                # Parse activity date
                $activityDate = $null
                if ($activity.created_at -is [hashtable] -and $activity.created_at.datetime)
                {
                  $activityDate = [DateTime]::Parse($activity.created_at.datetime)
                } elseif ($activity.created_at -is [string])
                {
                  $activityDate = [DateTime]::Parse($activity.created_at)
                }
                                
                $activityDateString = if ($activityDate)
                { $activityDate.ToString("yyyy-MM-dd") 
                } else
                { "" 
                }
                                
                if ($activity.action_type -eq "checkout" -and 
                  $activity.note -like "*TEMPORARY ASSIGNMENT*" -and
                  $activityDateString -eq $today)
                {
                  $isTemp = $true
                  break
                }
              }
            }
          } catch
          {
            # Continue with other checks
          }
                    
          # If no specific temp indicators found, but asset has expected_checkin date set, consider it a temp assignment
          if (-not $isTemp -and $assetDetails.expected_checkin)
          {
            $expectedDate = $null
            # Handle expected_checkin as hashtable/object
            if ($assetDetails.expected_checkin -is [hashtable] -and $assetDetails.expected_checkin.date)
            {
              try
              {
                $expectedDate = [DateTime]::Parse($assetDetails.expected_checkin.date)
              } catch
              {
                # Try the formatted field
                if ($assetDetails.expected_checkin.formatted)
                {
                  $expectedDate = [DateTime]::Parse($assetDetails.expected_checkin.formatted)
                }
              }
            } elseif ($assetDetails.expected_checkin -is [string])
            {
              $expectedDate = [DateTime]::Parse($assetDetails.expected_checkin)
            }
                        
            if ($expectedDate -and $checkoutDate -and $expectedDate.Date -ge $checkoutDate.Date)
            {
              $isTemp = $true
            }
          }
                    
          if ($isTemp)
          {
            $tempAssignments += $assetDetails
          }
        }
      } catch
      {
        # Continue with next asset if there's an error
        continue
      }
    }
        
    if ($tempAssignments.Count -gt 0)
    {
      Write-Host "`nTemp Laptop Assignments Today: $($tempAssignments.Count)" -ForegroundColor Green
      Write-Host "`nAsset Tag | Checkout Date       | Assigned To" -ForegroundColor White
      Write-Host "----------|-------------------|------------------" -ForegroundColor Gray
            
      foreach ($assignment in $tempAssignments)
      {
        # Get checkout date for display
        $checkoutDisplay = "N/A"
        if ($assignment.last_checkout)
        {
          if ($assignment.last_checkout.formatted)
          {
            $checkoutDisplay = $assignment.last_checkout.formatted
          } elseif ($assignment.last_checkout.datetime)
          {
            # Parse and reformat the datetime
            try
            {
              $checkoutDate = [DateTime]::Parse($assignment.last_checkout.datetime)
              $checkoutDisplay = $checkoutDate.ToString("yyyy-MM-dd HH:mm")
            } catch
            {
              $checkoutDisplay = $assignment.last_checkout.datetime
            }
          } else
          {
            $checkoutDisplay = $assignment.last_checkout
          }
        }
                
        # Get assigned user
        $assignedUser = if ($assignment.assigned_to -and $assignment.assigned_to.name)
        { $assignment.assigned_to.name 
        } else
        { "N/A" 
        }
                
        # Format the output in columns
        $assetTag = $assignment.asset_tag.PadRight(9)
        $checkoutFormatted = $checkoutDisplay.PadRight(17)
        $userFormatted = $assignedUser
                
        Write-Host "$assetTag | $checkoutFormatted | $userFormatted" -ForegroundColor Yellow
      }
    } else
    {
      Write-Host "`nNo temporary laptop assignments found for today ($today)." -ForegroundColor Yellow
    }
        
    Write-Host "`nReport Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Gray
  } catch
  {
    Write-Host "Error retrieving today's temp assignments: $($_.Exception.Message)" -ForegroundColor Red
  }
    
  Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
  [void][System.Console]::ReadKey($true)
}

function ReturnTempLaptop
{
  Write-Host "`n===========================================" -ForegroundColor Green
  Write-Host "       RETURN TEMPORARY LAPTOP            " -ForegroundColor Green
  Write-Host "===========================================" -ForegroundColor Green
    
  # Get Asset Tag
  $assetTag = Read-Host "`nEnter Temporary Laptop Asset Tag to return"
  if ([string]::IsNullOrWhiteSpace($assetTag))
  {
    Write-Host "Asset tag cannot be empty." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Validate asset exists
  $asset = Get-AssetByTag -AssetTag $assetTag
  if (-not $asset)
  {
    Write-Host "Asset with tag $assetTag not found." -ForegroundColor Red
    Start-Sleep -Seconds 3
    return
  }
    
  # Check if asset is currently assigned
  if (-not $asset.assigned_to)
  {
    Write-Host "Asset $assetTag is not currently assigned to anyone." -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    return
  }
    
  Write-Host "Asset: $($asset.name) - $($asset.model.name)" -ForegroundColor White
  Write-Host "Currently assigned to: $($asset.assigned_to.name)" -ForegroundColor Cyan
    
  # Confirm it's a temp assignment
  if ($asset.notes -notlike "*TEMPORARY ASSIGNMENT*")
  {
    Write-Host "Warning: This asset is not marked as a temporary assignment." -ForegroundColor Yellow
    $confirm = Read-Host "Continue with check-in anyway? (y/N)"
    if ($confirm -ne 'y' -and $confirm -ne 'Y')
    {
      return
    }
  }
    
  # Get return notes
  $returnNotes = Read-Host "Enter return condition/notes (optional)"
    
  try
  {
    # Check in the asset using Reset-SnipeitAssetOwner (to location 24)
    $noteText = if ($returnNotes)
    { "TEMP RETURN: $returnNotes" 
    } else
    { "TEMP RETURN" 
    }
    $checkinResult = Reset-SnipeitAssetOwner -id $asset.id -location_id 24 -note $noteText
        
    if ($checkinResult)
    {
      Write-Host "`nSuccess! Temporary laptop $assetTag returned successfully." -ForegroundColor Green
      Write-Host "Return Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Green
      Write-Host "Returned by: $($asset.assigned_to.name)" -ForegroundColor Green
      Write-Host "Checked into location 24" -ForegroundColor Green
      if ($returnNotes)
      {
        Write-Host "Return Notes: $returnNotes" -ForegroundColor Green
      }
    } else
    {
      Write-Host "Failed to return temporary laptop." -ForegroundColor Red
    }
  } catch
  {
    Write-Host "Error returning temporary laptop: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error details: $($_.Exception.ToString())" -ForegroundColor Red
  }
    
  Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
  [void][System.Console]::ReadKey($true)
}

function Process-TempBrokenMenu
{
  $continue = $true
    
  while ($continue)
  {
    Show-TempBrokenMenu
    $selection = Read-Host
        
    switch ($selection)
    {
      "1"
      { AssignTempLaptop 
      }
      "2"
      { ReportBrokenLaptop 
      }
      "3"
      { ViewBrokenLaptopStats 
      }
      "4"
      { ViewTodayTempAssignments 
      }
      "5"
      { ReturnTempLaptop 
      }
      "6"
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

# Export the main function that will be called from the menu script
Export-ModuleMember -Function Process-TempBrokenMenu
