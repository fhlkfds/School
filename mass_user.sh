#!/bin/bash

set -euo pipefail

INPUT="$1"
DOMAIN="test.lan"
PASSWORD="Password@1"
FAILED_OUTPUT="failed_users.csv"

declare -r DOMAIN
declare -r PASSWORD
declare -r FAILED_OUTPUT

print_usage() {
  printf "Usage: %s <csv_file>\n" "$(basename "$0")" >&2
}

validate_arguments() {
  if [[ $# -ne 1 ]]; then
    print_usage
    return 1
  fi
  if [[ ! -f "$1" || ! -r "$1" ]]; then
    printf "Error: File '%s' does not exist or is not readable.\n" "$1" >&2
    return 1
  fi
  return 0
}

sanitize_name() {
  local name="$1"
  name=$(printf "%s" "$name" | tr -d '[:space:]' | tr '[:upper:]' '[:lower:]')
  printf "%s" "$(printf "%s" "$name" | tr -cd '[:alnum:]')"
}

sanitize_last_name() {
  local last="$1"
  local last_clean

  if [[ "$last" == *"-"* ]]; then
    last_clean=$(printf "%s" "$last" | tr '[:upper:]' '[:lower:]' | tr -d '[:space:]')
    last_clean=$(printf "%s" "$last_clean" | tr -cd '[:alnum:]-')
  else
    last_clean=$(printf "%s" "$last" | awk '{print $1}' | tr '[:upper:]' '[:lower:]' | tr -cd '[:alnum:]')
  fi

  printf "%s" "$last_clean"
}

get_org_unit() {
  local grade="$1"
  case "$grade" in
    8) echo "/Cadets/8th Grade" ;;
    9) echo "/Cadets/9th Grade" ;;
    10) echo "/Cadets/10th Grade" ;;
    11) echo "/Cadets/11th Grade" ;;
    12) echo "/Cadets/12th Grade" ;;
    *) echo "/Cadets/Unknown" ;;
  esac
}

generate_unique_username() {
  local base="$1"
  local username="$base"
  local counter=1
  while gam info user "$username" >/dev/null 2>&1; do
    username="${base}${counter}"
    ((counter++))
  done
  printf "%s" "$username"
}

create_user() {
  local first="$1"
  local last="$2"
  local grade="$3"

  if [[ -z "$first" || -z "$last" || -z "$grade" ]]; then
    printf "Skipping user due to missing data: %s, %s, %s\n" "$first" "$last" "$grade" >&2
    printf "%s,%s,%s,Missing data\n" "$first" "$last" "$grade" >> "$FAILED_OUTPUT"
    return
  fi

  local first_clean; first_clean=$(sanitize_name "$first")
  local last_clean; last_clean=$(sanitize_last_name "$last")

  local base_username
  if [[ "$last" == *"-"* ]]; then
    base_username="cadet${first_clean:0:1}${last_clean}"
  else
    base_username="cadet${first_clean:0:1}${last_clean}"
  fi

  local username; username=$(generate_unique_username "$base_username")
  local email="${username}@${DOMAIN}"

  local org_unit; org_unit=$(get_org_unit "$grade")
  if [[ "$org_unit" == "/Cadets/Unknown" ]]; then
    printf "Skipping user %s %s due to unknown grade: %s\n" "$first" "$last" "$grade" >&2
    printf "%s,%s,%s,Unknown grade\n" "$first" "$last" "$grade" >> "$FAILED_OUTPUT"
    return
  fi

  printf "Creating user: %s in OU: %s\n" "$email" "$org_unit"
  if ! gam create user "$username" firstname "$first" lastname "$last" password "$PASSWORD" org "$org_unit"; then
    printf "Failed to create user %s %s\n" "$first" "$last" >&2
    printf "%s,%s,%s,gam create failed\n" "$first" "$last" "$grade" >> "$FAILED_OUTPUT"
  else
    printf "Successfully created %s\n" "$email"
  fi
}

main() {
  if ! validate_arguments "$@"; then return 1; fi

  printf "first_name,last_name,grade,reason\n" > "$FAILED_OUTPUT"

  local input="$1"
  while IFS=, read -r first last grade; do
    [[ "$first" == "first_name" ]] && continue
    create_user "$first" "$last" "$grade"
  done < "$input"
}

main "$@"

