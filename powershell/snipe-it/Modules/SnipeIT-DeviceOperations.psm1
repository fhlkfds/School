# SnipeIT-DeviceOperations.psm1
# Module for handling Snipe-IT device operations (check-in/check-out)

function Invoke-SnipeITRequest
{
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Uri,
        
    [Parameter(Mandatory = $true)]
    [string]$ApiKey,
        
    [Parameter(Mandatory = $true)]
    [string]$Method,
        
    [Parameter(Mandatory = $false)]
    [object]$Body
  )
    
  $headers = @{
    "Accept" = "application/json"
    "Content-Type" = "application/json"
    "Authorization" = "Bearer $ApiKey"
  }
    
  try
  {
    if ($Body)
    {
      $response = Invoke-RestMethod -Uri $Uri -Headers $headers -Method $Method -Body ($Body | ConvertTo-Json) -ErrorAction Stop
    } else
    {
      $response = Invoke-RestMethod -Uri $Uri -Headers $headers -Method $Method -ErrorAction Stop
    }
    return $response
  } catch
  {
    Write-Error "API Request failed: $_"
    return $null
  }
}

function Get-UserByUsername
{
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,
        
    [Parameter(Mandatory = $true)]
    [string]$ApiKey,
        
    [Parameter(Mandatory = $true)]
    [string]$Username
  )
    
  $uri = "$ApiUrl/users?search=$Username"
  $response = Invoke-SnipeITRequest -Uri $uri -ApiKey $ApiKey -Method "GET"
    
  if ($response -and $response.rows.Count -gt 0)
  {
    return $response.rows | Where-Object { $_.username -eq $Username -or $_.name -eq $Username }
  }
    
  return $null
}

function Get-AssetByTag
{
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,
        
    [Parameter(Mandatory = $true)]
    [string]$ApiKey,
        
    [Parameter(Mandatory = $true)]
    [string]$AssetTag
  )
    
  $uri = "$ApiUrl/hardware/bytag/$AssetTag"
  $response = Invoke-SnipeITRequest -Uri $uri -ApiKey $ApiKey -Method "GET"
    
  return $response
}

# PRINTER FUNCTIONS
function Checkout-Printer
{
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,
        
    [Parameter(Mandatory = $true)]
    [string]$ApiKey
  )
    
  Write-Host "===== Printer Checkout =====" -ForegroundColor Cyan
    
  # Get printer asset tag
  $assetTag = Read-Host "Enter Printer Asset Tag"
  $printer = Get-AssetByTag -ApiUrl $ApiUrl -ApiKey $ApiKey -AssetTag $assetTag
    
  if (-not $printer)
  {
    Write-Host "Printer with Asset Tag '$assetTag' not found." -ForegroundColor Red
    return
  }
    
  Write-Host "Found printer: $($printer.name)" -ForegroundColor Green
    
  # Get location
  $location = Read-Host "Enter Location for this Printer"
    
  # Get user to assign to
  $username = Read-Host "Enter Username to assign this printer to (leave blank for location only)"
  $userId = $null
    
  if ($username)
  {
    $user = Get-UserByUsername -ApiUrl $ApiUrl -ApiKey $ApiKey -Username $username
    if (-not $user)
    {
      Write-Host "User '$username' not found." -ForegroundColor Red
      return
    }
    $userId = $user.id
  }
    
  # Build checkout request
  $checkoutBody = @{
    note = "Printer deployed to location: $location"
    checkout_to_type = if ($userId)
    { "user" 
    } else
    { "location" 
    }
  }
    
  if ($userId)
  {
    $checkoutBody.assigned_user = $userId
  } else
  {
    $checkoutBody.assigned_location = $location
  }
    
  # Perform checkout
  $uri = "$ApiUrl/hardware/$($printer.id)/checkout"
  $response = Invoke-SnipeITRequest -Uri $uri -ApiKey $ApiKey -Method "POST" -Body $checkoutBody
    
  if ($response)
  {
    Write-Host "Printer successfully checked out!" -ForegroundColor Green
    Write-Host "Location: $location" -ForegroundColor Green
    if ($userId)
    {
      Write-Host "Assigned to: $username" -ForegroundColor Green
    }
  } else
  {
    Write-Host "Failed to check out printer." -ForegroundColor Red
  }
}

function Checkin-Printer
{
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,
        
    [Parameter(Mandatory = $true)]
    [string]$ApiKey
  )
    
  Write-Host "===== Printer Check-in =====" -ForegroundColor Cyan
    
  # Get printer asset tag
  $assetTag = Read-Host "Enter Printer Asset Tag"
  $printer = Get-AssetByTag -ApiUrl $ApiUrl -ApiKey $ApiKey -AssetTag $assetTag
    
  if (-not $printer)
  {
    Write-Host "Printer with Asset Tag '$assetTag' not found." -ForegroundColor Red
    return
  }
    
  Write-Host "Found printer: $($printer.name)" -ForegroundColor Green
    
  # Add note
  $note = Read-Host "Enter check-in notes (leave blank for none)"
    
  # Perform check-in
  $uri = "$ApiUrl/hardware/$($printer.id)/checkin"
  $checkInBody = @{
    note = if ($note)
    { $note 
    } else
    { "Printer checked in" 
    }
  }
    
  $response = Invoke-SnipeITRequest -Uri $uri -ApiKey $ApiKey -Method "POST" -Body $checkInBody
    
  if ($response)
  {
    Write-Host "Printer successfully checked in!" -ForegroundColor Green
  } else
  {
    Write-Host "Failed to check in printer." -ForegroundColor Red
  }
}

# HOTSPOT FUNCTIONS
function Checkout-Hotspot
{
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,
        
    [Parameter(Mandatory = $true)]
    [string]$ApiKey
  )
    
  Write-Host "===== Hotspot Checkout =====" -ForegroundColor Cyan
    
  # Get hotspot asset tag and ID number
  $assetTag = Read-Host "Enter Hotspot Asset Tag"
  $idNumber = Read-Host "Enter Hotspot ID Number"
    
  $hotspot = Get-AssetByTag -ApiUrl $ApiUrl -ApiKey $ApiKey -AssetTag $assetTag
    
  if (-not $hotspot)
  {
    Write-Host "Hotspot with Asset Tag '$assetTag' not found." -ForegroundColor Red
    return
  }
    
  Write-Host "Found hotspot: $($hotspot.name)" -ForegroundColor Green
    
  # Get user to assign to
  $username = Read-Host "Enter Username to assign this hotspot to"
  $user = Get-UserByUsername -ApiUrl $ApiUrl -ApiKey $ApiKey -Username $username
    
  if (-not $user)
  {
    Write-Host "User '$username' not found." -ForegroundColor Red
    return
  }
    
  # Build checkout request
  $checkoutBody = @{
    note = "Hotspot checked out. ID Number: $idNumber"
    checkout_to_type = "user"
    assigned_user = $user.id
  }
    
  # Perform checkout
  $uri = "$ApiUrl/hardware/$($hotspot.id)/checkout"
  $response = Invoke-SnipeITRequest -Uri $uri -ApiKey $ApiKey -Method "POST" -Body $checkoutBody
    
  if ($response)
  {
    Write-Host "Hotspot successfully checked out!" -ForegroundColor Green
    Write-Host "ID Number: $idNumber" -ForegroundColor Green
    Write-Host "Assigned to: $username" -ForegroundColor Green
  } else
  {
    Write-Host "Failed to check out hotspot." -ForegroundColor Red
  }
}

function Checkin-Hotspot
{
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,
        
    [Parameter(Mandatory = $true)]
    [string]$ApiKey
  )
    
  Write-Host "===== Hotspot Check-in =====" -ForegroundColor Cyan
    
  # Get hotspot asset tag and ID number
  $assetTag = Read-Host "Enter Hotspot Asset Tag"
  $idNumber = Read-Host "Enter Hotspot ID Number"
    
  $hotspot = Get-AssetByTag -ApiUrl $ApiUrl -ApiKey $ApiKey -AssetTag $assetTag
    
  if (-not $hotspot)
  {
    Write-Host "Hotspot with Asset Tag '$assetTag' not found." -ForegroundColor Red
    return
  }
    
  Write-Host "Found hotspot: $($hotspot.name)" -ForegroundColor Green
    
  # Add note
  $note = Read-Host "Enter check-in notes (leave blank for none)"
    
  # Perform check-in
  $uri = "$ApiUrl/hardware/$($hotspot.id)/checkin"
  $checkInBody = @{
    note = if ($note)
    { $note 
    } else
    { "Hotspot checked in. ID Number: $idNumber" 
    }
  }
    
  $response = Invoke-SnipeITRequest -Uri $uri -ApiKey $ApiKey -Method "POST" -Body $checkInBody
    
  if ($response)
  {
    Write-Host "Hotspot successfully checked in!" -ForegroundColor Green
  } else
  {
    Write-Host "Failed to check in hotspot." -ForegroundColor Red
  }
}

# LAPTOP FUNCTIONS
function Checkout-Laptop
{
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,
        
    [Parameter(Mandatory = $true)]
    [string]$ApiKey
  )
    
  Write-Host "===== Laptop/Charger Checkout =====" -ForegroundColor Cyan
    
  # Get laptop info
  $pcAssetTag = Read-Host "Enter Laptop Asset Tag"
  $chargerAssetTag = Read-Host "Enter Charger Asset Tag"
  $idNumber = Read-Host "Enter Laptop ID Number"
    
  # Get laptop
  $laptop = Get-AssetByTag -ApiUrl $ApiUrl -ApiKey $ApiKey -AssetTag $pcAssetTag
  if (-not $laptop)
  {
    Write-Host "Laptop with Asset Tag '$pcAssetTag' not found." -ForegroundColor Red
    return
  }
    
  Write-Host "Found laptop: $($laptop.name)" -ForegroundColor Green
    
  # Get charger
  $charger = Get-AssetByTag -ApiUrl $ApiUrl -ApiKey $ApiKey -AssetTag $chargerAssetTag
  if (-not $charger)
  {
    Write-Host "Charger with Asset Tag '$chargerAssetTag' not found." -ForegroundColor Red
    return
  }
    
  Write-Host "Found charger: $($charger.name)" -ForegroundColor Green
    
  # Get user to assign to
  $username = Read-Host "Enter Username to assign this laptop/charger to"
  $user = Get-UserByUsername -ApiUrl $ApiUrl -ApiKey $ApiKey -Username $username
    
  if (-not $user)
  {
    Write-Host "User '$username' not found." -ForegroundColor Red
    return
  }
    
  # Build checkout request for laptop
  $laptopCheckoutBody = @{
    note = "Laptop checked out. ID Number: $idNumber. With charger: $chargerAssetTag"
    checkout_to_type = "user"
    assigned_user = $user.id
  }
    
  # Build checkout request for charger
  $chargerCheckoutBody = @{
    note = "Charger checked out with laptop: $pcAssetTag"
    checkout_to_type = "user"
    assigned_user = $user.id
  }
    
  # Perform laptop checkout
  $laptopUri = "$ApiUrl/hardware/$($laptop.id)/checkout"
  $laptopResponse = Invoke-SnipeITRequest -Uri $laptopUri -ApiKey $ApiKey -Method "POST" -Body $laptopCheckoutBody
    
  # Perform charger checkout
  $chargerUri = "$ApiUrl/hardware/$($charger.id)/checkout"
  $chargerResponse = Invoke-SnipeITRequest -Uri $chargerUri -ApiKey $ApiKey -Method "POST" -Body $chargerCheckoutBody
    
  if ($laptopResponse -and $chargerResponse)
  {
    Write-Host "Laptop and charger successfully checked out!" -ForegroundColor Green
    Write-Host "ID Number: $idNumber" -ForegroundColor Green
    Write-Host "Laptop Asset Tag: $pcAssetTag" -ForegroundColor Green
    Write-Host "Charger Asset Tag: $chargerAssetTag" -ForegroundColor Green
    Write-Host "Assigned to: $username" -ForegroundColor Green
  } else
  {
    Write-Host "Failed to check out laptop and/or charger." -ForegroundColor Red
  }
}

function Checkin-Laptop
{
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,
        
    [Parameter(Mandatory = $true)]
    [string]$ApiKey
  )
    
  Write-Host "===== Laptop/Charger Check-in =====" -ForegroundColor Cyan
    
  # Get laptop info
  $pcAssetTag = Read-Host "Enter Laptop Asset Tag"
  $chargerAssetTag = Read-Host "Enter Charger Asset Tag"
  $idNumber = Read-Host "Enter Laptop ID Number"
    
  # Get laptop
  $laptop = Get-AssetByTag -ApiUrl $ApiUrl -ApiKey $ApiKey -AssetTag $pcAssetTag
  if (-not $laptop)
  {
    Write-Host "Laptop with Asset Tag '$pcAssetTag' not found." -ForegroundColor Red
    return
  }
    
  Write-Host "Found laptop: $($laptop.name)" -ForegroundColor Green
    
  # Get charger
  $charger = Get-AssetByTag -ApiUrl $ApiUrl -ApiKey $ApiKey -AssetTag $chargerAssetTag
  if (-not $charger)
  {
    Write-Host "Charger with Asset Tag '$chargerAssetTag' not found." -ForegroundColor Red
    return
  }
    
  Write-Host "Found charger: $($charger.name)" -ForegroundColor Green
    
  # Add note
  $note = Read-Host "Enter check-in notes (leave blank for none)"
  $defaultNote = "Laptop checked in. ID Number: $idNumber. With charger: $chargerAssetTag"
    
  # Perform laptop check-in
  $laptopUri = "$ApiUrl/hardware/$($laptop.id)/checkin"
  $laptopCheckInBody = @{
    note = if ($note)
    { $note 
    } else
    { $defaultNote 
    }
  }
    
  $laptopResponse = Invoke-SnipeITRequest -Uri $laptopUri -ApiKey $ApiKey -Method "POST" -Body $laptopCheckInBody
    
  # Perform charger check-in
  $chargerUri = "$ApiUrl/hardware/$($charger.id)/checkin"
  $chargerCheckInBody = @{
    note = if ($note)
    { $note 
    } else
    { "Charger checked in with laptop: $pcAssetTag" 
    }
  }
    
  $chargerResponse = Invoke-SnipeITRequest -Uri $chargerUri -ApiKey $ApiKey -Method "POST" -Body $chargerCheckInBody
    
  if ($laptopResponse -and $chargerResponse)
  {
    Write-Host "Laptop and charger successfully checked in!" -ForegroundColor Green
  } else
  {
    Write-Host "Failed to check in laptop and/or charger." -ForegroundColor Red
  }
}

# Export the functions
Export-ModuleMember -Function Checkout-Printer, Checkin-Printer, Checkout-Hotspot, Checkin-Hotspot, Checkout-Laptop, Checkin-Laptop

