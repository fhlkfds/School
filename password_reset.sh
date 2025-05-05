#!/bin/bash

# Usage: ./reset_password.sh user@example.com

EMAIL="$1"

if [ -z "$EMAIL" ]; then
  echo "Usage: $0 user@example.com"
  exit 1
fi

# Generate a secure random password
NEW_PASSWORD=Password@1

# Reset the user's password using GAM
gam update user "$EMAIL" password "$NEW_PASSWORD" changepassword on

