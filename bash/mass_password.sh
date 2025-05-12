#!/bin/bash

# Script: reset_passwords_ou.sh
# Purpose: Reset passwords for all users in a given Google Workspace OU using GAM
# GAM version: GAMADV-XTD3 (tested with GAM 7+)

# Check if OU was passed as argument, otherwise prompt
if [ -z "$1" ]; then
  read -p "Enter the full OU path (e.g., /Students/Grade10): " OU
else
  OU="$1"
fi

# Check if GAM is installed
if ! command -v gam &> /dev/null; then
  echo "❌ GAM is not installed or not in your PATH. Please install GAMADV-XTD3 first."
  exit 1
fi

# Confirm action
echo "⚠️ You are about to reset passwords for ALL users in OU: $OU"
read -p "Type YES to continue: " CONFIRM

if [ "$CONFIRM" != "YES" ]; then
  echo "Operation cancelled."
  exit 1
fi

# Generate a temp file for user list
USERLIST=$(mktemp)

# Step 1: Get all users in the OU
echo "📥 Getting users in OU: $OU ..."
gam print users query "orgUnitPath='$OU'" fields primaryEmail > "$USERLIST"


# Step 2: Reset each password
echo "🔐 Resetting passwords to 'Password@1' and forcing password change..."
gam csv "$USERLIST" gam update user ~primaryEmail password "Password@1" changepassword on

# Clean up
rm "$USERLIST"

echo "✅ All user passwords in OU '$OU' have been reset to 'Password@1' and will be required to change them at next login!"

