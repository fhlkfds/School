#!/bin/bash

# === Configuration ===
SNIPE_IT_URL=""   # Change to your Snipe-IT URL
API_KEY="" # Your API key


# Define headers - do this once to avoid repeating in functions
HEADER_AUTH="Authorization: Bearer $API_KEY"
HEADER_CT="Content-Type: application/json"
CURL_OPTS="-s --connect-timeout 5 --max-time 10"

# === Function definitions ===
# Get user ID by Employee Number or Email
get_user_id() {
  local SEARCH_TERM=$1
  curl $CURL_OPTS -H "$HEADER_AUTH" "$SNIPE_IT_URL/api/v1/users?search=$SEARCH_TERM" | 
    jq -r ".rows[] | select(.employee_num==\"$SEARCH_TERM\" or .email==\"$SEARCH_TERM\") | .id"
}

# Get asset ID by tag
get_asset_id() {
  curl $CURL_OPTS -H "$HEADER_AUTH" "$SNIPE_IT_URL/api/v1/hardware/bytag/$1" | 
    jq -r '.id'
}

# Assign asset function
assign_asset() {
  local ASSET_ID=$1
  local USER_ID=$2
  local ASSET_TAG=$3
  
  local RESPONSE=$(curl $CURL_OPTS -X POST "$SNIPE_IT_URL/api/v1/hardware/$ASSET_ID/checkout" \
    -H "$HEADER_AUTH" \
    -H "$HEADER_CT" \
    -d "{\"checkout_to_type\": \"user\", \"assigned_user\": $USER_ID}")
  
  # Check if assignment was successful
  local STATUS=$(echo "$RESPONSE" | jq -r '.status' 2>/dev/null)
  if [[ "$STATUS" == "success" ]]; then
    return 0
  else
    return 1
  fi
}

# Main loop
while true; do
  # Get employee ID or email
  EMP_INPUT=$(zenity --entry --title="Employee Lookup" --text="Enter Employee Number or Email (or 'exit' to quit):")
  
  # Check if user wants to exit
  if [[ "$EMP_INPUT" == "exit" || -z "$EMP_INPUT" ]]; then
    zenity --info --title="Exit" --text="Exiting program..."
    exit 0
  fi
  
  # Get user ID and validate
  USER_ID=$(get_user_id "$EMP_INPUT")
  if [[ -z "$USER_ID" ]]; then
    zenity --error --title="Error" --text="❌ No user found with Employee Number or Email '$EMP_INPUT'"
    continue
  fi
  
  # Get laptop asset tag
  LAPTOP_TAG=$(zenity --entry --title="Laptop Tag" --text="Enter laptop asset tag:")
  if [[ -z "$LAPTOP_TAG" ]]; then
    zenity --error --title="Error" --text="❌ Laptop tag cannot be empty!"
    continue
  fi
  
  # Get charger asset tag
  CHARGER_TAG=$(zenity --entry --title="Charger Tag" --text="Enter charger asset tag:")
  if [[ -z "$CHARGER_TAG" ]]; then
    zenity --error --title="Error" --text="❌ Charger tag cannot be empty!"
    continue
  fi
  
  # Get laptop ID and validate
  LAPTOP_ID=$(get_asset_id "$LAPTOP_TAG")
  if [[ -z "$LAPTOP_ID" || "$LAPTOP_ID" == "null" ]]; then
    zenity --error --title="Error" --text="❌ Could not find laptop with tag '$LAPTOP_TAG'"
    continue
  fi
  
  # Get charger ID and validate
  CHARGER_ID=$(get_asset_id "$CHARGER_TAG")
  if [[ -z "$CHARGER_ID" || "$CHARGER_ID" == "null" ]]; then
    zenity --error --title="Error" --text="❌ Could not find charger with tag '$CHARGER_TAG'"
    continue
  fi
  
  # Attempt to assign laptop
  if ! assign_asset "$LAPTOP_ID" "$USER_ID" "$LAPTOP_TAG"; then
    zenity --error --title="Error" --text="❌ Failed to assign laptop with tag '$LAPTOP_TAG'"
    continue
  fi
  
  # Attempt to assign charger
  if ! assign_asset "$CHARGER_ID" "$USER_ID" "$CHARGER_TAG"; then
    zenity --error --title="Error" --text="❌ Failed to assign charger with tag '$CHARGER_TAG'"
    continue
  fi
  
  # Confirm both assets were assigned
  zenity --info --title="Success" --text="✅ Laptop ($LAPTOP_TAG) and charger ($CHARGER_TAG) successfully assigned to user '$EMP_INPUT'"
done

