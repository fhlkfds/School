# Configuration
$CsvPath = "/home/liam/laptop_replacements.csv"



# Read the CSV (comma-delimited as corrected)
$records = Import-Csv -Path $CsvPath -Delimiter ","

foreach ($record in $records) {
    # Extract details from CSV
    $firstName = $record.'Cadet First Name:'
    $lastName = $record.'Cadet Last Name:'
    $studentID = $record.'What is your student ID number? (All 7 digits)'
    $brokenAssetTag = $record.'What is the Asset Tag (The numbers on the red sticker) on the broken laptop?'
    $newAssetTag = $record.'What is the Asset Tag (The numbers on the red sticker) on the new laptop?'

    try {
        # 1. Get the User Internal ID
        Write-Output "🔍 Searching for user by employee number: $studentID"
        $user = Get-SnipeitUser -Search $studentID | Where-Object { $_.employee_num -eq $studentID }

        if (!$user) {
            Write-Output "❗ Could not find user with employee number: $studentID"
            continue
        }

        $userID = $user.id
        $userName = $user.name
        Write-Output "🆔 Found User: $userName (ID: $userID)"

        # 2. Get the Broken Laptop Internal ID
        Write-Output "🔍 Searching for broken asset: $brokenAssetTag"
        $brokenAsset = Get-SnipeitAsset -Search $brokenAssetTag | Where-Object { $_.asset_tag -eq $brokenAssetTag }
        
        if (!$brokenAsset) {
            Write-Output "❗ Could not find broken laptop with tag: $brokenAssetTag"
            continue
        }

        $brokenAssetID = $brokenAsset.id
        Write-Output "🆔 Found Broken Laptop: $brokenAssetTag (ID: $brokenAssetID)"
        
        # 3. Unassign and Archive the Broken Laptop
        Write-Output "📦 Archiving broken laptop: $brokenAssetTag (ID: $brokenAssetID)"
        Reset-SnipeitAssetOwner -id $brokenAssetID -status_id 4
        Write-Output "✔️ Archived and unassigned broken laptop: $brokenAssetTag"

        # 4. Get the New Laptop Internal ID
        Write-Output "🔍 Searching for new asset: $newAssetTag"
        $newAsset = Get-SnipeitAsset -Search $newAssetTag | Where-Object { $_.asset_tag -eq $newAssetTag }

        if (!$newAsset) {
            Write-Output "❗ Could not find new laptop with tag: $newAssetTag"
            continue
        }

        $newAssetID = $newAsset.id
        Write-Output "🆔 Found New Laptop: $newAssetTag (ID: $newAssetID)"

        # 5. Assign the New Laptop to the User
        Write-Output "💻 Assigning new laptop: $newAssetTag (ID: $newAssetID) to user: $userID"
        Set-SnipeitAssetOwner -id $newAssetID  -assigned_id $userID 
        Set-SnipeitAsset -id $newAssetID  -status_id 1  
        Write-Output "✔️ Assigned new laptop: $newAssetTag to $userName (User ID: $userID)"
        
    } catch {
        Write-Output "❗ Error processing record for $firstName $lastName ($studentID): $_"
    }
}

