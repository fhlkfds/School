#!/bin/bash

# === Configuration ===
SNIPE_IT_URL="http://inv.school.org"   # Change to your Snipe-IT URL
API_KEY="" # Your API key



# Define headers - do this once to avoid repeating in functions
HEADER_AUTH="Authorization: Bearer $API_KEY"
HEADER_CT="Content-Type: application/json"
CURL_OPTS="-s --connect-timeout 5 --max-time 10"

# === Function definitions ===
# Get user ID
get_user_id() {
  curl $CURL_OPTS -H "$HEADER_AUTH" "$SNIPE_IT_URL/api/v1/users?search=$1" | 
    jq -r ".rows[] | select(.employee_num==\"$1\") | .id"
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
  
  echo "Assigning $ASSET_TAG..."
  local RESPONSE=$(curl $CURL_OPTS -X POST "$SNIPE_IT_URL/api/v1/hardware/$ASSET_ID/checkout" \
    -H "$HEADER_AUTH" \
    -H "$HEADER_CT" \
    -d "{\"checkout_to_type\": \"user\", \"assigned_user\": $USER_ID}")
  
  # Check if we got an HTML response (error)
  if [[ "$RESPONSE" == "<!DOCTYPE html>"* ]]; then
    echo "❌ Error assigning asset $ASSET_TAG. Server returned an error."
    return 1
  fi
  
  # Check if assignment was successful
  local STATUS=$(echo "$RESPONSE" | jq -r '.status' 2>/dev/null)
  if [[ "$STATUS" == "success" ]]; then
    echo "✅ $ASSET_TAG assigned successfully"
    return 0
  else
    echo "❌ Failed to assign $ASSET_TAG"
    return 1
  fi
}

# Clear screen function
clear_screen() {
  echo -e "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n"
  echo "===== Snipe-IT Asset Assignment Tool ====="
  echo "Type 'exit' for Employee Number to quit"
  echo "=========================================="
  echo
}

# Main loop
while true; do
  clear_screen
  
  # Get employee number
  read -p "Enter Employee Number (or 'exit' to quit): " EMP_NUM
  
  # Check if user wants to exit
  if [[ "$EMP_NUM" == "exit" ]]; then
    echo "Exiting program..."
    exit 0
  fi
  
  # Get asset tags
  read -p "Enter laptop asset tag: " LAPTOP_TAG
  read -p "Enter charger asset tag: " CHARGER_TAG
  
  echo "Processing assignment..."
  
  # Get user ID and validate
  USER_ID=$(get_user_id "$EMP_NUM")
  if [[ -z "$USER_ID" ]]; then
    echo "❌ No user found with Employee Number $EMP_NUM"
    read -p "Press Enter to continue..."
    continue
  fi
  
  # Get laptop ID and validate
  LAPTOP_ID=$(get_asset_id "$LAPTOP_TAG")
  if [[ -z "$LAPTOP_ID" || "$LAPTOP_ID" == "null" ]]; then
    echo "❌ Could not find laptop with tag $LAPTOP_TAG"
    read -p "Press Enter to continue..."
    continue
  fi
  
  # Get charger ID and validate
  CHARGER_ID=$(get_asset_id "$CHARGER_TAG")
  if [[ -z "$CHARGER_ID" || "$CHARGER_ID" == "null" ]]; then
    echo "❌ Could not find charger with tag $CHARGER_TAG"
    read -p "Press Enter to continue..."
    continue
  fi
  
  # Assign assets (sequentially for better error reporting)
  LAPTOP_SUCCESS=true
  CHARGER_SUCCESS=true
  
  if ! assign_asset "$LAPTOP_ID" "$USER_ID" "$LAPTOP_TAG"; then
    LAPTOP_SUCCESS=false
  fi
  
  if ! assign_asset "$CHARGER_ID" "$USER_ID" "$CHARGER_TAG"; then
    CHARGER_SUCCESS=false
  fi
  
  # Summary
  echo
  if $LAPTOP_SUCCESS && $CHARGER_SUCCESS; then
    echo "✅ All assets successfully assigned to employee $EMP_NUM"
  else
    echo "⚠️ Some assignments may have failed. Check messages above."
  fi
  
  read -p "Press Enter to continue to next assignment..."
done
