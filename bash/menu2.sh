#!/bin/bash

# Google Workspace Admin Toolkit using GAM
# Created: May 14, 2025
# Requirements: GAM must be installed and configured

# Check if GAM is installed
if ! command -v gam &> /dev/null; then
    echo "GAM is not installed or not in path. Please install GAM first."
    echo "Visit https://github.com/GAM-team/GAM for installation instructions."
    exit 1
fi

# Clear the screen
clear

# Function to display the main menu
show_menu() {
    echo "=========================================="
    echo "    GOOGLE WORKSPACE ADMIN TOOLKIT        "
    echo "=========================================="
    echo "1. Password Management"
    echo "2. User Management"
    echo "3. Device Management"
    echo "4. Information Retrieval"
    echo "5. Exit"
    echo "=========================================="
    echo "Enter your choice [1-5]: "
}

# Function to display password management submenu
password_menu() {
    clear
    echo "=========================================="
    echo "         PASSWORD MANAGEMENT              "
    echo "=========================================="
    echo "1. Reset Single User Password"
    echo "2. Mass Password Reset"
    echo "3. Return to Main Menu"
    echo "=========================================="
    echo "Enter your choice [1-3]: "
}

# Function to display user management submenu
user_menu() {
    clear
    echo "=========================================="
    echo "           USER MANAGEMENT                "
    echo "=========================================="
    echo "1. Create Single User"
    echo "2. Create Multiple Users"
    echo "3. Archive User"
    echo "4. Archive Multiple Users"
    echo "5. Suspend User"
    echo "6. Return to Main Menu"
    echo "=========================================="
    echo "Enter your choice [1-6]: "
}

# Function to display device management submenu
device_menu() {
    clear
    echo "=========================================="
    echo "          DEVICE MANAGEMENT               "
    echo "=========================================="
    echo "1. Lock Single Mobile Device"
    echo "2. Lock Multiple Mobile Devices"
    echo "3. Wipe Device"
    echo "4. List All Devices"
    echo "5. Return to Main Menu"
    echo "=========================================="
    echo "Enter your choice [1-5]: "
}

# Function to display information retrieval submenu
info_menu() {
    clear
    echo "=========================================="
    echo "        INFORMATION RETRIEVAL             "
    echo "=========================================="
    echo "1. Get Device Serial Number"
    echo "2. Get User Information"
    echo "3. List All Users"
    echo "4. Check License Information"
    echo "5. Return to Main Menu"
    echo "=========================================="
    echo "Enter your choice [1-5]: "
}

# Function to reset single user password
reset_password() {
    clear
    echo "Reset Single User Password"
    echo "-------------------------"
    read -p "Enter user email: " useremail
    read -p "Enter new password: " password
    echo "Resetting password for $useremail..."
    gam update user "$useremail" password "$password"
    echo "Password has been reset for $useremail"
    read -p "Press Enter to continue..."
}

# Function to perform mass password reset
mass_reset_password() {
    clear
    echo "Mass Password Reset"
    echo "------------------"
    read -p "Enter file with user emails (one per line): " userfile
    read -p "Enter default password or leave blank for random: " defaultpwd
    
    if [ -f "$userfile" ]; then
        echo "Resetting passwords for users in $userfile..."
        
        if [ -z "$defaultpwd" ]; then
            # Use random password for each user
            while IFS= read -r user; do
                randpwd=Password@1
                echo "Resetting password for $user to a random password..."
                gam update user "$user" password "$randpwd"
                echo "$user,$randpwd" >> password_reset_results.csv
            done < "$userfile"
            echo "Passwords have been reset for all users. Results saved to password_reset_results.csv"
        else
            # Use provided default password
            while IFS= read -r user; do
                echo "Resetting password for $user..."
                gam update user "$user" password "$defaultpwd"
            done < "$userfile"
            echo "Passwords have been reset to the default password for all users"
        fi
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
    read -p "Enter first name: " firstname
    read -p "Enter last name: " lastname
    read -p "Enter username (before @domain.com): " username
    read -p "Enter org unit path (leave blank for root): " orgunit
    read -p "Enter password (leave blank for random): " password
    
    domain=$(gam info domain | grep "Primary Domain:" | cut -d ":" -f2 | xargs)
    email="${username}@${domain}"
    
    echo "Creating user $email..."
    
    if [ -z "$password" ]; then
        # Random password
        password=Password@1
        echo "Generated random password: $password"
    fi
    
    cmd="gam create user $email firstname \"$firstname\" lastname \"$lastname\" password \"$password\""
    
    if [ ! -z "$orgunit" ]; then
        cmd="$cmd org \"$orgunit\""
    fi
    
    eval $cmd
    
    echo "User $email has been created with password: $password"
    read -p "Press Enter to continue..."
}

# Function to create multiple users
create_multi_user() {
    clear
    echo "Create Multiple Users"
    echo "--------------------"
    echo "CSV file should be in format: firstname,lastname,username,orgunit(optional)"
    read -p "Enter CSV file with user info: " userfile
    
    if [ -f "$userfile" ]; then
        echo "Creating users from $userfile..."
        domain=$(gam info domain | grep "Primary Domain:" | cut -d ":" -f2 | xargs)
        
        # Create a results file
        resultfile="user_creation_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "firstname,lastname,email,password,status" > "$resultfile"
        
        # Process each line of the CSV
        while IFS=, read -r firstname lastname username orgunit; do
            email="${username}@${domain}"
            password=Password@1
            
            echo "Creating user $email..."
            
            cmd="gam create user $email firstname \"$firstname\" lastname \"$lastname\" password \"$password\""
            
            if [ ! -z "$orgunit" ]; then
                cmd="$cmd org \"$orgunit\""
            fi
            
            if eval $cmd; then
                echo "$firstname,$lastname,$email,$password,success" >> "$resultfile"
            else
                echo "$firstname,$lastname,$email,$password,failed" >> "$resultfile"
            fi
            
        done < "$userfile"
        
        echo "User creation complete. Results saved to $resultfile"
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
    read -p "Enter user email to archive: " useremail
    read -p "Enter admin email to transfer data to: " adminemail
    
    echo "Archiving user $useremail..."
    echo "This will:"
    echo "1. Back up the user's data"
    echo "2. Transfer Drive files to admin"
    echo "3. Suspend the account"
    
    read -p "Continue? (y/n): " confirm
    if [[ $confirm == [Yy]* ]]; then
        echo "Creating backup of user data..."
        gam create datatransfer $useremail gdrive,gmail $adminemail
        
        echo "Suspending account..."
        gam update user $useremail suspended on
        
        echo "User $useremail has been archived"
        echo "Data transfer to $adminemail is in progress"
    else
        echo "Operation cancelled"
    fi
    
    read -p "Press Enter to continue..."
}

# Function to archive multiple users
archive_multi_user() {
    clear
    echo "Archive Multiple Users"
    echo "---------------------"
    read -p "Enter file with user emails to archive (one per line): " userfile
    read -p "Enter admin email to transfer data to: " adminemail
    
    if [ -f "$userfile" ]; then
        echo "Archiving users from $userfile..."
        
        read -p "This will archive all listed users. Continue? (y/n): " confirm
        if [[ $confirm == [Yy]* ]]; then
            while IFS= read -r useremail; do
                echo "Archiving $useremail..."
                gam create datatransfer $useremail gdrive,gmail $adminemail
                gam update user $useremail suspended on
                echo "$useremail archived, data transfer to $adminemail in progress"
            done < "$userfile"
            echo "All users have been archived"
        else
            echo "Operation cancelled"
        fi
    else
        echo "File not found!"
    fi
    read -p "Press Enter to continue..."
}

# Function to suspend a user
suspend_user() {
    clear
    echo "Suspend User"
    echo "-----------"
    read -p "Enter user email to suspend: " useremail
    
    echo "Suspending user $useremail..."
    gam update user $useremail suspended on
    
    echo "User $useremail has been suspended"
    read -p "Press Enter to continue..."
}

# Function to lock a mobile device
lock_device() {
    clear
    echo "Lock Mobile Device"
    echo "-----------------"
    read -p "Enter user email: " useremail
    
    echo "Retrieving mobile devices for $useremail..."
    gam print mobile query "user:$useremail" > devices.txt
    
    echo "Available devices:"
    cat devices.txt | awk -F',' '{print NR ") " $1 " - " $2 " - " $3}'
    
    read -p "Enter device number to lock: " devicenum
    
    deviceid=$(sed -n "${devicenum}p" devices.txt | cut -d',' -f1)
    
    if [ ! -z "$deviceid" ]; then
        echo "Locking device $deviceid..."
        gam update mobile "$deviceid" action accountlock
        echo "Device has been locked"
    else
        echo "Invalid device selection"
    fi
    
    rm devices.txt
    read -p "Press Enter to continue..."
}

# Function to lock multiple mobile devices
lock_multi_device() {
    clear
    echo "Lock Multiple Mobile Devices"
    echo "--------------------------"
    read -p "Enter file with user emails (one per line): " userfile
    
    if [ -f "$userfile" ]; then
        echo "Locking devices for users in $userfile..."
        
        while IFS= read -r useremail; do
            echo "Processing devices for $useremail..."
            gam print mobile query "user:$useremail" | while IFS=, read -r deviceid rest; do
                if [ ! -z "$deviceid" ]; then
                    echo "Locking device $deviceid for $useremail..."
                    gam update mobile "$deviceid" action accountlock
                fi
            done
        done < "$userfile"
        
        echo "All devices have been locked"
    else
        echo "File not found!"
    fi
    read -p "Press Enter to continue..."
}

# Function to wipe a device
wipe_device() {
    clear
    echo "Wipe Device"
    echo "----------"
    read -p "Enter user email: " useremail
    
    echo "Retrieving mobile devices for $useremail..."
    gam print mobile query "user:$useremail" > devices.txt
    
    echo "Available devices:"
    cat devices.txt | awk -F',' '{print NR ") " $1 " - " $2 " - " $3}'
    
    read -p "Enter device number to wipe: " devicenum
    
    deviceid=$(sed -n "${devicenum}p" devices.txt | cut -d',' -f1)
    
    if [ ! -z "$deviceid" ]; then
        echo "WARNING: This will erase all data on the device!"
        read -p "Are you absolutely sure? (type 'WIPE' to confirm): " confirm
        
        if [ "$confirm" = "WIPE" ]; then
            echo "Wiping device $deviceid..."
            gam update mobile "$deviceid" action wipe
            echo "Wipe command has been sent to the device"
        else
            echo "Wipe operation cancelled"
        fi
    else
        echo "Invalid device selection"
    fi
    
    rm devices.txt
    read -p "Press Enter to continue..."
}

# Function to list all devices
list_devices() {
    clear
    echo "List All Devices"
    echo "---------------"
    read -p "Enter output file name [devices_list.csv]: " outfile
    
    if [ -z "$outfile" ]; then
        outfile="devices_list.csv"
    fi
    
    echo "Retrieving all mobile devices..."
    gam print mobile > "$outfile"
    
    echo "Device list has been saved to $outfile"
    read -p "Press Enter to continue..."
}

# Function to get device serial number
get_serial() {
    clear
    echo "Get Device Serial Number"
    echo "-----------------------"
    read -p "Enter user email: " useremail
    
    echo "Retrieving devices for $useremail..."
    gam print cros query "user:$useremail" fields serialNumber,annotatedUser > device_info.csv
    
    if [ -s device_info.csv ]; then
        echo "Device information:"
        cat device_info.csv
        echo "Information saved to device_info.csv"
    else
        echo "No Chrome devices found for $useremail"
        echo "Checking mobile devices..."
        gam print mobile query "user:$useremail" fields serialNumber > mobile_info.csv
        
        if [ -s mobile_info.csv ]; then
            echo "Mobile device information:"
            cat mobile_info.csv
            echo "Information saved to mobile_info.csv"
        else
            echo "No devices found for $useremail"
        fi
    fi
    
    read -p "Press Enter to continue..."
}

# Function to get user information
get_user_info() {
    clear
    echo "Get User Information"
    echo "------------------"
    read -p "Enter user email: " useremail
    
    echo "Retrieving information for $useremail..."
    gam info user "$useremail"
    
    read -p "Press Enter to continue..."
}

# Function to list all users
list_all_users() {
    clear
    echo "List All Users"
    echo "-------------"
    read -p "Enter output file name [all_users.csv]: " outfile
    
    if [ -z "$outfile" ]; then
        outfile="all_users.csv"
    fi
    
    echo "Retrieving all users..."
    gam print users > "$outfile"
    
    echo "User list has been saved to $outfile"
    read -p "Press Enter to continue..."
}

# Function to check license information
check_license() {
    clear
    echo "Check License Information"
    echo "-----------------------"
    read -p "Enter user email (leave blank for all): " useremail
    
    if [ -z "$useremail" ]; then
        echo "Retrieving license information for all users..."
        gam print licenses > license_info.csv
        echo "License information saved to license_info.csv"
    else
        echo "Retrieving license information for $useremail..."
        gam user "$useremail" show licenses
    fi
    
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
                    5) suspend_user ;;
                    6) break ;;
                    *) echo "Invalid option. Press Enter to continue..."; read ;;
                esac
            done
            ;;
            
        3)  # Device Management
            while true; do
                device_menu
                read -r subchoice
                case $subchoice in
                    1) lock_device ;;
                    2) lock_multi_device ;;
                    3) wipe_device ;;
                    4) list_devices ;;
                    5) break ;;
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
                    3) list_all_users ;;
                    4) check_license ;;
                    5) break ;;
                    *) echo "Invalid option. Press Enter to continue..."; read ;;
                esac
            done
            ;;
            
        5)  # Exit
            echo "Exiting the Google Workspace Admin Toolkit. Goodbye!"
            exit 0
            ;;
            
        *)  # Invalid option
            echo "Invalid option. Press Enter to continue..."
            read
            ;;
    esac
done
