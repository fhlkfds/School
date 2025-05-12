#!/bin/bash

read -p "Enter the source Google Group email: " source_group
read -p "Enter the destination Google Group email: " destination_group
read -p "Enter the path to the CSV file (with 'email' column): " csv_file

if [ ! -f "$csv_file" ]; then
    echo "CSV file not found!"
    exit 1
fi

echo "Starting to move users from $source_group to $destination_group..."

gam csv "$csv_file" gam update group "$source_group" remove member ~email
gam csv "$csv_file" gam update group "$destination_group" add member ~email

echo "✅ Done moving users."

