#!/usr/bin/env bash

set -euo pipefail
IFS=$'\n\t'

CSV_FILE=""
EMAIL_LIST=()
INACTIVE_OU="/Inactive"

sanitize_and_validate_file() {
    local file="$1"

    if [[ -z "$file" || ! -f "$file" || ! -r "$file" ]]; then
        printf "Error: CSV file is either missing, unreadable, or not specified: %s\n" "$file" >&2
        return 1
    fi

    if ! file "$file" | grep -qi 'text'; then
        printf "Error: The file does not appear to be a valid text-based CSV: %s\n" "$file" >&2
        return 1
    fi
}

extract_emails_from_column() {
    local file="$1"
    local emails

    if ! emails=$(cut -d',' -f3 "$file" | grep -Eoi '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b'); then
        printf "Error: Failed to extract emails from column 3.\n" >&2
        return 1
    fi

    mapfile -t EMAIL_LIST < <(printf "%s\n" "$emails" | sort -u)
}


disable_rename_move_user() {
    local email username domain new_username

    for email in "${EMAIL_LIST[@]}"; do
        if [[ -z "$email" ]]; then
            continue
        fi

        username="${email%@*}"
        domain="${email#*@}"

        if [[ -z "$username" || -z "$domain" ]]; then
            printf "Error: Invalid email format: %s\n" "$email" >&2
            continue
        fi

        new_username="Inactive${username}"
        

        echo "Lift off in 5 4 3 2 1"

        gam update user "$email" username "${new_username}@${domain}"

        sleep 5
        
        echo "Add for my next trick I will be suppended '$new_username'"
        gam update user "$email" Suspended on

        sleep 3
        echo "Moving '$new_username' to '$INACTIVE_OU'"

        gam update user "$email" org "$INACTIVE_OU" 

        echo "Next deleting user alias"

        sleep 2

        gam delete alias user "$username"
        
        sleep 2

        printf "Processed: %s → Inactive%s@%s\n" "$email" "$username" "$domain"
    done
}

main() {
    if [[ $# -ne 1 ]]; then
        printf "Usage: %s <csv_file>\n" "$0" >&2
        return 1
    fi

    CSV_FILE="$1"

    sanitize_and_validate_file "$CSV_FILE" || return 1
    extract_emails_from_column "$CSV_FILE" || return 1
    disable_rename_move_user
}

main "$@"

