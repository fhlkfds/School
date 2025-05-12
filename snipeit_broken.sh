#!/bin/bash


sheet="/home/liam/broken.csv"
#rclone copy "sheet:Copy of Broken Laptop Form (Responses).xlsx" /home/liam/ --drive-shared-with-me

#ssconvert /home/liam/Copy\ of\ Broken\ Laptop\ Form\ \(Responses\).xlsx /home/liam/broken.csv  
API_URL=""
API_TOKEN=""


# Extract student IDs as a single variable
STUDENT_IDS=$(awk -F, 'NR > 1 {gsub(/"/, "", $4); print $4}' "$sheet")
BROKEN_TAGS=$(awk -F, 'NR > 1 {gsub(/"/, "", $6); if ($6 ~ /^[0-9]{1,4}$/) print $6}' "$sheet")
NEW_TAGS=$(awk -F, 'NR > 1 {gsub(/"/, "", $7); if ($7 ~ /^[0-9]{1,4}$/) print $7}' "$sheet")

# Define headers to avoid repetition
HEADER_AUTH="Authorization: Bearer $API_TOKEN"
HEADER_CT="Content-Type: application/json"
CURL_OPTS="-s --connect-timeout 5 --max-time 10"

#echo "$STUDENT_IDS"
#echo "$BROKEN_TAGS"
#echo "$NEW_TAGS"

# Function to get user ID by student ID
get_user_id() {
    local SEARCH_TERM=$1
    curl $CURL_OPTS -H "$HEADER_AUTH" "$API_URL/users?search=$SEARCH_TERM" | \
        jq -r ".rows[] | select(.employee_num==\"$SEARCH_TERM\" or .username==\"$SEARCH_TERM\") | .id"
}

# Function to get asset ID by tag
get_asset_id() {
    curl $CURL_OPTS -H "$HEADER_AUTH" "$API_URL/hardware/bytag/$1" | \
        jq -r '.id'
}

# Function to assign asset to user
assign_asset() {
    local ASSET_ID=$1
    local USER_ID=$2
    local ASSET_TAG=$3
    
    local RESPONSE=$(curl $CURL_OPTS -X POST "$API_URL/hardware/$ASSET_ID/checkout" \
        -H "$HEADER_AUTH" \
        -H "$HEADER_CT" \
        -d "{\"checkout_to_type\": \"user\", \"assigned_user\": $USER_ID}")
    
    local STATUS=$(echo "$RESPONSE" | jq -r '.status' 2>/dev/null)
    if [[ "$STATUS" == "success" ]]; then
        echo "✅ Successfully assigned asset $ASSET_TAG to user ID $USER_ID"
        return 0
    else
        echo "❌ Failed to assign asset $ASSET_TAG"
        echo "API Response: $RESPONSE"
        return 1
    fi
}

# Function to check in a broken asset
checkin_asset() {
    local ASSET_ID=$1
    local ASSET_TAG=$2
    
    local RESPONSE=$(curl $CURL_OPTS -X POST "$API_URL/hardware/$ASSET_ID/checkin" \
        -H "$HEADER_AUTH" \
        -H "$HEADER_CT" \
        -d "{\"status_id\": 2, \"note\": \"Automatically checked in due to broken status\", \"location_id\": 1}")
    
    local STATUS=$(echo "$RESPONSE" | jq -r '.status' 2>/dev/null)
    if [[ "$STATUS" == "success" ]]; then
        echo "✅ Successfully checked in broken asset $ASSET_TAG"
        return 0
    else
        echo "❌ Failed to check in broken asset $ASSET_TAG"
        echo "API Response: $RESPONSE"
        return 1
    fi
}

# Main loop to process each student and their devices
for i in "${!BROKEN_TAGS[@]}"; do
    TAG="${BROKEN_TAGS[$i]}"
    STUDENT_ID="${STUDENT_IDS[$i]}"
    NEW_TAG="${NEW_TAGS[$i]}"

    # Check in the broken asset first
    BROKEN_ASSET_ID=$(get_asset_id "$TAG")
    if [[ -z "$BROKEN_ASSET_ID" || "$BROKEN_ASSET_ID" == "null" ]]; then
        echo "❌ No broken asset found for tag $TAG"
        continue
    fi
    
    if ! checkin_asset "$BROKEN_ASSET_ID" "$TAG"; then
        echo "❌ Failed to check in broken laptop $TAG"
        continue
    fi

    # Get user ID
    USER_ID=$(get_user_id "$STUDENT_ID")
    if [[ -z "$USER_ID" || "$USER_ID" == "null" ]]; then
        echo "❌ No user found for student ID $STUDENT_ID"
        continue
    fi

    # Get new asset ID
    NEW_ASSET_ID=$(get_asset_id "$NEW_TAG")
    if [[ -z "$NEW_ASSET_ID" || "$NEW_ASSET_ID" == "null" ]]; then
        echo "❌ No asset found for tag $NEW_TAG"
        continue
    fi

    # Attempt to assign the new laptop
    if ! assign_asset "$NEW_ASSET_ID" "$USER_ID" "$NEW_TAG"; then
        echo "❌ Failed to assign laptop $NEW_TAG to student $STUDENT_ID"
        continue
    fi

done
