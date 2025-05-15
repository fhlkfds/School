#!/bin/bash

# Admin Menu System
# Created: May 14, 2025

# Clear the screen
clear

# Function to display the main menu
show_menu() {
    echo "===================================="
    echo "       SYSTEM ADMIN TOOLKIT         "
    echo "===================================="
    echo "1. Password Management"
    echo "2. User Management"
    echo "3. Device Management"
    echo "4. Information Retrieval"
    echo "5. Exit"
    echo "===================================="
    echo "Enter your choice [1-5]: "
}

# Function to display password management submenu
password_menu() {
    clear
    echo "===================================="
    echo "      PASSWORD MANAGEMENT           "
    echo "===================================="
    echo "1. Reset Single User Password"
    echo "2. Mass Password Reset"
    echo "3. Return to Main Menu"
    echo "===================================="
    echo "Enter your choice [1-3]: "
}

# Function to display user management submenu
user_menu() {
    clear
    echo "===================================="
    echo "        USER MANAGEMENT             "
    echo "===================================="
    echo "1. Create Single User"
    echo "2. Create Multiple Users"
    echo "3. Archive User"
    echo "4. Archive Multiple Users"
    echo "5. Return to Main Menu"
    echo "===================================="
    echo "Enter your choice [1-5]: "
}

# Function to display device management submenu
device_menu() {
    clear
    echo "===================================="
    echo "       DEVICE MANAGEMENT            "
    echo "===================================="
    echo "1. Lock Laptop"
    echo "2. Lock Multiple Laptops"
    echo "3. Return to Main Menu"
    echo "===================================="
    echo "Enter your choice [1-3]: "
}

# Function to display information retrieval submenu
info_menu() {
    clear
    echo "===================================="
    echo "     INFORMATION RETRIEVAL          "
    echo "===================================="
    echo "1. Get Device Serial Number"
    echo "2. Get User Information"
    echo "3. Return to Main Menu"
    echo "===================================="
    echo "Enter your choice [1-3]: "
}

# Function to reset single user password
reset_password() {
    clear
    echo "Reset Single User Password"
    echo "-------------------------"
    read -p "Enter username: " username
    echo "Resetting password for $username..."
    # Add your actual password reset command here
    # For example: passwd $username
    echo "Password has been reset for $username"
    read -p "Press Enter to continue..."
}

# Function to perform mass password reset
mass_reset_password() {
    clear
    echo "Mass Password Reset"
    echo "------------------"
    read -p "Enter file with usernames (one per line): " userfile
    if [ -f "$userfile" ]; then
        echo "Resetting passwords for users in $userfile..."
        # Add your actual mass password reset logic here
        # For example: while read user; do passwd $user; done < $userfile
        echo "Passwords have been reset for all users in the file"
    else
        echo "File not found!"
    fi
    read -p "Press Enter to continue..."
}

# Function to create a single user
create_user() {
    clear
    echo "Create Single User"
    echo "-----------------"
    read -p "Enter new username: " username
    # Add your actual user creation command here
    # For example: useradd $username
    echo "User $username has been created"
    read -p "Press Enter to continue..."
}

# Function to create multiple users
create_multi_user() {
    clear
    echo "Create Multiple Users"
    echo "--------------------"
    read -p "Enter file with usernames (one per line): " userfile
    if [ -f "$userfile" ]; then
        echo "Creating users from $userfile..."
        # Add your actual multi-user creation logic here
        # For example: while read user; do useradd $user; done < $userfile
        echo "All users have been created"
    else
        echo "File not found!"
    fi
    read -p "Press Enter to continue..."
}

# Function to archive a single user
archive_user() {
    clear
    echo "Archive User"
    echo "------------"
    read -p "Enter username to archive: " username
    # Add your actual user archiving command here
    echo "User $username has been archived"
    read -p "Press Enter to continue..."
}

# Function to archive multiple users
archive_multi_user() {
    clear
    echo "Archive Multiple Users"
    echo "---------------------"
    read -p "Enter file with usernames to archive (one per line): " userfile
    if [ -f "$userfile" ]; then
        echo "Archiving users from $userfile..."
        # Add your actual multi-user archiving logic here
        echo "All users have been archived"
    else
        echo "File not found!"
    fi
    read -p "Press Enter to continue..."
}

# Function to lock a laptop
lock_laptop() {
    clear
    echo "Lock Laptop"
    echo "----------"
    read -p "Enter laptop ID or hostname: " laptop
    # Add your actual laptop locking command here
    echo "Laptop $laptop has been locked"
    read -p "Press Enter to continue..."
}

# Function to lock multiple laptops
lock_multi_laptop() {
    clear
    echo "Lock Multiple Laptops"
    echo "-------------------"
    read -p "Enter file with laptop IDs (one per line): " laptopfile
    if [ -f "$laptopfile" ]; then
        echo "Locking laptops from $laptopfile..."
        # Add your actual multi-laptop locking logic here
        echo "All laptops have been locked"
    else
        echo "File not found!"
    fi
    read -p "Press Enter to continue..."
}

# Function to get device serial number
get_serial() {
    clear
    echo "Get Device Serial Number"
    echo "-----------------------"
    read -p "Enter device hostname or IP: " device
    # Add your actual serial number retrieval command here
    # For example: ssh $device 'dmidecode -s system-serial-number'
    echo "Serial Number: XYZ123456789"  # Replace with actual command output
    read -p "Press Enter to continue..."
}

# Function to get user information
get_user_info() {
    clear
    echo "Get User Information"
    echo "------------------"
    read -p "Enter username: " username
    # Add your actual user info retrieval command here
    # For example: id $username; finger $username
    echo "User information for $username displayed above"
    read -p "Press Enter to continue..."
}

# Main program loop
while true; do
    clear
    show_menu
    read -r choice
    
    case $choice in
        1)  # Password Management
            while true; do
                password_menu
                read -r subchoice
                case $subchoice in
                    1) reset_password ;;
                    2) mass_reset_password ;;
                    3) break ;;
                    *) echo "Invalid option. Press Enter to continue..."; read ;;
                esac
            done
            ;;
            
        2)  # User Management
            while true; do
                user_menu
                read -r subchoice
                case $subchoice in
                    1) create_user ;;
                    2) create_multi_user ;;
                    3) archive_user ;;
                    4) archive_multi_user ;;
                    5) break ;;
                    *) echo "Invalid option. Press Enter to continue..."; read ;;
                esac
            done
            ;;
            
        3)  # Device Management
            while true; do
                device_menu
                read -r subchoice
                case $subchoice in
                    1) lock_laptop ;;
                    2) lock_multi_laptop ;;
                    3) break ;;
                    *) echo "Invalid option. Press Enter to continue..."; read ;;
                esac
            done
            ;;
            
        4)  # Information Retrieval
            while true; do
                info_menu
                read -r subchoice
                case $subchoice in
                    1) get_serial ;;
                    2) get_user_info ;;
                    3) break ;;
                    *) echo "Invalid option. Press Enter to continue..."; read ;;
                esac
            done
            ;;
            
        5)  # Exit
            echo "Exiting the Admin Toolkit. Goodbye!"
            exit 0
            ;;
            
        *)  # Invalid option
            echo "Invalid option. Press Enter to continue..."
            read
            ;;
    esac
done
