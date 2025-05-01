#!/bin/bash

# Path to the CSV file
CSV_FILE="/home/liam/users.csv"

# Check if the CSV file exists
if [[ ! -f "$CSV_FILE" ]]; then
    echo "CSV file not found: $CSV_FILE"
    exit 1
fi

# Read the CSV, skip header
tail -n +2 "$CSV_FILE" | while IFS=',' read -r col1 col2 username col4; do
    # Trim whitespace from username
    username=$(echo "$username" | xargs)

    if [[ -z "$username" ]]; then
        echo "Skipping empty username"
        continue
    fi

    echo "Processing user: $username"

    # Suspend the user
    gam update user "$username" suspended on

    echo "3"
    # Delete all aliases for the user
    gam user "$username" delete aliases

    # Move user to the 'Inactive' organizational unit
    gam update user "$username" org "/Inactive"

    # Extract the domain from the email
    domain="${username##*@}"
    local_part="${username%@*}"
    new_email="inactive${local_part}@${domain}"
    alias_email="email${local_part}@${domain}"


    echo "2"
    # Change the user's primary email address
    gam update user "$username" email "$new_email"
    sleep 10
    
    echo "1"
    # Add the original email as an alias to the new email
# Add the 'inactive' alias to the user
    echo "Deleting alias $original_email from user $new_email"
    gam user "$new_email" delete alias "$original_email"
done


