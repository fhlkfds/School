#!/bin/bash


sheet="/home/liam/broken.csv"
#rclone copy "sheet:Copy of Broken Laptop Form (Responses).xlsx" /home/liam/ --drive-shared-with-me


# Load environment variables for security
API_URL="https://inv.nomma.lan/api/v1"
API_TOKEN="${API_TOKEN:-$(cat /home/liam/.snipeit_token)}"
HEADER_AUTH="Authorization: Bearer $API_TOKEN"
HEADER_CT="Content-Type: application/json"
CURL_OPTS="-s --connect-timeout 5 --max-time 10"

# Extract student IDs, broken tags, and new tags
IFS=$'\n' read -r -d '' -a STUDENT_IDS < <(awk -F, 'NR > 1 {gsub(/"/, "", $4); print $4}' "$sheet" && printf '\0')
IFS=$'\n' read -r -d '' -a BROKEN_TAGS < <(awk -F, 'NR > 1 {gsub(/"/, "", $6); if ($6 ~ /^[0-9]{1,4}$/) print $6}' "$sheet" && printf '\0')
IFS=$'\n' read -r -d '' -a NEW_TAGS < <(awk -F, 'NR > 1 {gsub(/"/, "", $7); if ($7 ~ /^[0-9]{1,4}$/) print $7}' "$sheet" && printf '\0')
curl -s -H "$HEADER_AUTH" "$API_URL/hardware?search=956" | jq .


# Function to get user ID by student ID
get_user_id() {
    local SEARCH_TERM=$1
    local RESPONSE=$(curl $CURL_OPTS -H "$HEADER_AUTH" "$API_URL/users?search=$SEARCH_TERM")
    echo "$RESPONSE" | jq -r ".rows[] | select(.employee_num==\"$SEARCH_TERM\" or .username==\"$SEARCH_TERM\") | .id"
}

# Function to get asset ID by tag
get_asset_id() {
    local TAG=$1
    local RESPONSE=$(curl $CURL_OPTS -H "$HEADER_AUTH" "$API_URL/hardware?search=$TAG")
    
    # Check if the response is valid JSON and contains the asset ID
    if echo "$RESPONSE" | jq . >/dev/null 2>&1; then
        echo "$RESPONSE" | jq -r '.rows[] | select(.asset_tag=="'"$TAG"'") | .id'
    else
        echo "❌ Invalid response for tag $TAG"
        echo "Response body: $RESPONSE"
        return 1
    fi
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

    local STATUS=$(echo "$RESPONSE" | jq -r '.status')
    if [[ "$STATUS" == "success" ]]; then
        echo "✅ Successfully assigned asset $ASSET_TAG to user ID $USER_ID"
    else
        echo "❌ Failed to assign asset $ASSET_TAG"
        echo "API Response: $RESPONSE"
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

    local STATUS=$(echo "$RESPONSE" | jq -r '.status')
    if [[ "$STATUS" == "success" ]]; then
        echo "✅ Successfully checked in broken asset $ASSET_TAG"
    else
        echo "❌ Failed to check in broken asset $ASSET_TAG"
        echo "API Response: $RESPONSE"
    fi
}

# Main loop to process each student and their devices
for ((i=0; i<${#BROKEN_TAGS[@]}; i++)); do
    TAG="${BROKEN_TAGS[$i]}"
    STUDENT_ID="${STUDENT_IDS[$i]}"
    NEW_TAG="${NEW_TAGS[$i]}"

    # Check in the broken asset
    BROKEN_ASSET_ID=$(get_asset_id "$TAG")
    if [[ -z "$BROKEN_ASSET_ID" || "$BROKEN_ASSET_ID" == "null" ]]; then
        echo "❌ No broken asset found for tag $TAG"
        continue
    fi

    checkin_asset "$BROKEN_ASSET_ID" "$TAG"

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
    assign_asset "$NEW_ASSET_ID" "$USER_ID" "$NEW_TAG"
done

