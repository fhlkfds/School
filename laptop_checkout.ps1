# === Configuration ===
$SNIPE_IT_URL = ""
$API_KEY = "" # Your API key

# Define headers
$headers = @{
    'Authorization' = "Bearer $API_KEY"
    'Content-Type' = 'application/json'
}

# === Function definitions ===
# Get user ID
function Get-UserID
{
    param($employeeNum)
    
    $response = Invoke-RestMethod -Uri "$SNIPE_IT_URL/api/v1/users?search=$employeeNum" -Headers $headers -Method Get -ErrorAction SilentlyContinue
    if ($response.rows.Count -eq 0)
    {
        return $null
    }
    
    foreach ($user in $response.rows)
    {
        if ($user.employee_num -eq $employeeNum)
        {
            return $user.id
        }
    }
    return $null
}

# Get asset ID by tag
function Get-AssetID
{
    param($assetTag)
    
    try
    {
        $response = Invoke-RestMethod -Uri "$SNIPE_IT_URL/api/v1/hardware/bytag/$assetTag" -Headers $headers -Method Get -ErrorAction SilentlyContinue
        return $response.id
    } catch
    {
        return $null
    }
}

# Assign asset function
function Assign-Asset
{
    param($assetID, $userID, $assetTag)
    
    Write-Host "Assigning $assetTag..." -NoNewline
    
    $body = @{
        checkout_to_type = "user"
        assigned_user = $userID
    } | ConvertTo-Json
    
    try
    {
        $response = Invoke-RestMethod -Uri "$SNIPE_IT_URL/api/v1/hardware/$assetID/checkout" -Headers $headers -Method Post -Body $body -ErrorAction Stop
        Write-Host " ✅ Success!" -ForegroundColor Green
        return $true
    } catch
    {
        Write-Host " ❌ Failed!" -ForegroundColor Red
        Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Clear screen function
function Clear-WorkScreen
{
    Clear-Host
    Write-Host "===== Snipe-IT Asset Assignment Tool =====" -ForegroundColor Cyan
    Write-Host "Type 'exit' for Employee Number to quit" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
}

# Main loop
while ($true)
{
    Clear-WorkScreen
    
    # Get employee number
    $EMP_NUM = Read-Host "Enter Employee Number (or 'exit' to quit)"
    
    # Check if user wants to exit
    if ($EMP_NUM -eq "exit")
    {
        Write-Host "Exiting program..." -ForegroundColor Cyan
        exit
    }
    
    # Get asset tags
    $LAPTOP_TAG = Read-Host "Enter laptop asset tag"
    $CHARGER_TAG = Read-Host "Enter charger asset tag"
    
    Write-Host "`nProcessing assignment..." -ForegroundColor Cyan
    
    # Get user ID and validate
    $USER_ID = Get-UserID -employeeNum $EMP_NUM
    if (-not $USER_ID)
    {
        Write-Host "❌ No user found with Employee Number $EMP_NUM" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        continue
    }
    
    # Get laptop ID and validate
    $LAPTOP_ID = Get-AssetID -assetTag $LAPTOP_TAG
    if (-not $LAPTOP_ID)
    {
        Write-Host "❌ Could not find laptop with tag $LAPTOP_TAG" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        continue
    }
    
    # Get charger ID and validate
    $CHARGER_ID = Get-AssetID -assetTag $CHARGER_TAG
    if (-not $CHARGER_ID)
    {
        Write-Host "❌ Could not find charger with tag $CHARGER_TAG" -ForegroundColor Red
        Read-Host "Press Enter to continue..."
        continue
    }
    
    # Assign assets
    $LAPTOP_SUCCESS = Assign-Asset -assetID $LAPTOP_ID -userID $USER_ID -assetTag $LAPTOP_TAG
    $CHARGER_SUCCESS = Assign-Asset -assetID $CHARGER_ID -userID $USER_ID -assetTag $CHARGER_TAG
    
    # Summary
    Write-Host ""
    if ($LAPTOP_SUCCESS -and $CHARGER_SUCCESS)
    {
        Write-Host "✅ All assets successfully assigned to employee $EMP_NUM" -ForegroundColor Green
    } else
    {
        Write-Host "⚠️ Some assignments may have failed. Check messages above." -ForegroundColor Yellow
    }
    
    Read-Host "`nPress Enter to continue to next assignment..."
}
