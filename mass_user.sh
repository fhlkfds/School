#!/bin/bash

INPUT="new_cadets.csv"
DOMAIN="test.lan"
PASSWORD="Password@1"

# Map grades to org units
get_org_unit() {
  local grade="$1"
  case "$grade" in
    9) echo "/Students/Grade9" ;;
    10) echo "/Students/Grade10" ;;
    11) echo "/Students/Grade11" ;;
    12) echo "/Students/Grade12" ;;
    *) echo "/Students/Unknown" ;;
  esac
}

# Read CSV and create users
while IFS=, read -r first last grade
do
  # Skip header
  if [[ "$first" == "first_name" ]]; then continue; fi

  # Clean and lowercase names
  first_clean=$(echo "$first" | tr -d '[:space:]' | tr '[:upper:]' '[:lower:]')
  last_clean=$(echo "$last" | tr -d '[:space:]' | tr '[:upper:]' '[:lower:]')

  # Build base username
  base_username="cadet${first_clean:0:1}${last_clean}"
  username="$base_username"
  email="${username}@${DOMAIN}"
  counter=1

  # Check for existing username and find a unique one
  while gam info user "$email" >/dev/null 2>&1; do
    username="${base_username}${counter}"
    email="${username}@${DOMAIN}"
    ((counter++))
  done

  # Determine org unit
  org_unit=$(get_org_unit "$grade")

  # Create user
  echo "Creating user $email in OU $org_unit"
  gam create user "$username" firstname "$first" lastname "$last" password "$PASSWORD" org "$org_unit"

done < "$INPUT"

