# Check if SnipeitPS module is installed and import it
function Initialize-SnipeIT {
    if (-not (Get-Module -ListAvailable -Name SnipeitPS)) {
        Write-Host "SnipeitPS module is not installed. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name SnipeitPS -Force -Scope CurrentUser
            Write-Host "SnipeitPS module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install SnipeitPS module. Please install it manually with 'Install-Module -Name SnipeitPS'." -ForegroundColor Red
            return $false
        }
    }
    
    Import-Module SnipeitPS
    
    # Import our custom Snipe-IT checking module from modules directory
    # Use $PSScriptRoot instead of Split-Path -Parent $MyInvocation.ScriptName for better reliability
    $scriptDir = $PSScriptRoot
    $moduleDir = Join-Path $scriptDir "modules"
    $modulePath = Join-Path $moduleDir "snipeit_check.psm1"
    
    if (Test-Path $modulePath) {
        Import-Module $modulePath -Force
        Write-Host "Successfully imported snipeit_check module." -ForegroundColor Green
    }
    else {
        Write-Host "Warning: snipeit_check.psm1 module not found in script directory. Some functions may not work." -ForegroundColor Yellow
        Write-Host "Expected path: $modulePath" -ForegroundColor Yellow
    }
    
    # Import the temporary and broken device management module
    $tempBrokenModulePath = Join-Path $moduleDir "snipeit_temp_broken.psm1"
    if (Test-Path $tempBrokenModulePath) {
        Import-Module $tempBrokenModulePath -Force
        Write-Host "Successfully imported snipeit_temp_broken module." -ForegroundColor Green
    }
    else {
        Write-Host "Warning: snipeit_temp_broken.psm1 module not found in modules directory. Temp/Broken Device Management will not be available." -ForegroundColor Yellow
        Write-Host "Expected path: $tempBrokenModulePath" -ForegroundColor Yellow
    }
    
    # Set Snipe-IT API parameters - Replace with your actual values
    $apiUrl = "https://inv.nomma.lan"
    $apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIzIiwianRpIjoiM2MzMGM0NzBmNmRkMGJiODQ0NGI1MDc2NjkwZWY2MzA0MTUzZjQzODhiMDA1MTEzNmM4NzE4NDc2NjQwOTZkYTk2YTQ5MDY4NWRlMzhkY2UiLCJpYXQiOjE3NDc0MDUyODUuNjUzOTY2LCJuYmYiOjE3NDc0MDUyODUuNjUzOTczLCJleHAiOjIyMjA3OTA4ODUuNjMzMTQ5LCJzdWIiOiIxIiwic2NvcGVzIjpbXX0.Bx9nImH2fmqQeWp1KjUi1_hbSpgOF-YJ__BLOthuhct-9OyaJzTgq7wjZ8DvYtEFvVIAr-_wXI37Ihe1PqOPT5SPuYqHl1vES51OQEFHNVHcPSPjQ5gJFraKY4f8Yqs26V5jiEYKo-z7wGfHRpEKAg3MzC8GgfIZUCbh-Xg5OmvdCjtYLQrsFB1G4M2alkGQyBzotI2QV__76JlA1dQIUdAX_6ZNadjxEVG0-GF1CPOO4IrYPZN-YZ6zztCEO8lR0vxSGj-Dtu1WCPqJM4iuE1Jy5TUeyLTCMOtk2Nw_G-LD_w_W6hhEhsxMca8HPwvDnN7V8YHYx1V5uTE5nacHw_gTTpK70kLV-XECljW3rSwfoV0SepHTml3GEECZk4EyNr4vK5DSk5DZwfrjUjzOBGqphyqH1q6mNU296-H5L7OrKfwEIO2HUscuyS6842JDBVZoFH-L2WEYc_PuX2Nndbc0vj0MW8kZUjypVJOn0_biBs2-xEPzgN7mroYGMf5xyaeWgomwEfaA-GX2fYfp7ovWLUhe4KXkkW16kGGgqkKqA62lC8lDYrUbCJuATJGMgBDNGeiSroldB7XlCmmskOwq2AcCUNyKbKMQZWJS89BgYjFmyM__Djv18-3oa0JW6w1norstpOzL8VExCSvM3jE-p2C0r9VMzEYKbaUfrcQ"
    
    try {
        Set-SnipeitInfo -URL $apiUrl -APIKey $apiKey
        Write-Host "Successfully connected to Snipe-IT API." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to connect to Snipe-IT API. Please check your connection settings." -ForegroundColor Red
        return $false
    }
}
