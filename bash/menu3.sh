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
    echo "2. Mass Password Reset (CSV)"
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
    echo "2. Create Multiple Users (CSV)"
    echo "3. Archive User"
    echo "4. Archive Multiple Users (CSV)"
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
    echo "2. Lock Multiple Mobile Devices (CSV)"
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

# Function to perform mass password reset using CSV
mass_reset_password() {
    clear
    echo "Mass Password Reset (CSV)"
    echo "------------------------"
    echo "CSV file should have headers: email,password"
    echo "Leave password column blank for random passwords"
    
    read -p "Enter CSV file path: " csvfile
    
    if [ -f "$csvfile" ]; then
        echo "Resetting passwords for users in $csvfile..."
        
        # Create output file for results
        resultfile="password_reset_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "email,new_password,status" > "$resultfile"
        
        # Skip header line and process each user
        tail -n +2 "$csvfile" | while IFS=, read -r email password || [[ -n "$email" ]]; do
            # Remove any whitespace
            email=$(echo "$email" | xargs)
            password=$(echo "$password" | xargs)
            
            echo "Processing $email..."
            
            # Generate random password if not provided
            if [ -z "$password" ]; then
                password=$(openssl rand -base64 12)
            fi
            
            # Reset the password
            if gam update user "$email" password "$password"; then
                echo "$email,$password,success" >> "$resultfile"
                echo "Password reset successful for $email"
            else
                echo "$email,$password,failed" >> "$resultfile"
                echo "Password reset failed for $email"
            fi
        done
        
        echo "Password reset complete. Results saved to $resultfile"
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
        password=$(openssl rand -base64 12)
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

# Function to create multiple users from CSV
create_multi_user() {
    clear
    echo "Create Multiple Users (CSV)"
    echo "-------------------------"
    echo "CSV file should have headers: firstname,lastname,username,orgunit,password"
    echo "Note: orgunit and password columns are optional"
    echo "Leave password blank for random passwords"
    
    read -p "Enter CSV file path: " csvfile
    
    if [ -f "$csvfile" ]; then
        echo "Creating users from $csvfile..."
        domain=$(gam info domain | grep "Primary Domain:" | cut -d ":" -f2 | xargs)
        
        # Create a results file
        resultfile="user_creation_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "firstname,lastname,email,password,status" > "$resultfile"
        
        # Skip header line and process each user
        tail -n +2 "$csvfile" | while IFS=, read -r firstname lastname username orgunit password || [[ -n "$firstname" ]]; do
            # Remove any whitespace
            firstname=$(echo "$firstname" | xargs)
            lastname=$(echo "$lastname" | xargs)
            username=$(echo "$username" | xargs)
            orgunit=$(echo "$orgunit" | xargs)
            password=$(echo "$password" | xargs)
            
            email="${username}@${domain}"
            
            echo "Creating user $email..."
            
            # Generate random password if not provided
            if [ -z "$password" ]; then
                password=$(openssl rand -base64 12)
            fi
            
            cmd="gam create user $email firstname \"$firstname\" lastname \"$lastname\" password \"$password\""
            
            if [ ! -z "$orgunit" ]; then
                cmd="$cmd org \"$orgunit\""
            fi
            
            if eval $cmd; then
                echo "$firstname,$lastname,$email,$password,success" >> "$resultfile"
                echo "User creation successful for $email"
            else
                echo "$firstname,$lastname,$email,$password,failed" >> "$resultfile"
                echo "User creation failed for $email"
            fi
        done
        
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

# Function to archive multiple users using CSV
archive_multi_user() {
    clear
    echo "Archive Multiple Users (CSV)"
    echo "--------------------------"
    echo "CSV file should have headers: email,admin"
    echo "- email: The user email to archive"
    echo "- admin: The admin email to transfer data to"
    
    read -p "Enter CSV file path: " csvfile
    
    if [ -f "$csvfile" ]; then
        echo "Archiving users from $csvfile..."
        
        # Create output file for results
        resultfile="archive_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "email,admin,status" > "$resultfile"
        
        read -p "This will archive all listed users. Continue? (y/n): " confirm
        if [[ $confirm == [Yy]* ]]; then
            # Skip header line and process each user
            tail -n +2 "$csvfile" | while IFS=, read -r email admin || [[ -n "$email" ]]; do
                # Remove any whitespace
                email=$(echo "$email" | xargs)
                admin=$(echo "$admin" | xargs)
                
                echo "Archiving $email..."
                
                if gam create datatransfer $email gdrive,gmail $admin && gam update user $email suspended on; then
                    echo "$email,$admin,success" >> "$resultfile"
                    echo "Archive successful for $email, data transfer to $admin in progress"
                else
                    echo "$email,$admin,failed" >> "$resultfile"
                    echo "Archive failed for $email"
                fi
            done
            
            echo "Archive process complete. Results saved to $resultfile"
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

# Function to lock multiple mobile devices using CSV
lock_multi_device() {
    clear
    echo "Lock Multiple Mobile Devices (CSV)"
    echo "-------------------------------"
    echo "CSV file should have headers: email,deviceid"
    echo "- email: User's email (if deviceid not provided)"
    echo "- deviceid: Optional - specific device ID to lock"
    echo "If deviceid is blank, all devices for that user will be locked"
    
    read -p "Enter CSV file path: " csvfile
    
    if [ -f "$csvfile" ]; then
        echo "Locking devices from $csvfile..."
        
        # Create output file for results
        resultfile="device_lock_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "email,deviceid,status" > "$resultfile"
        
        # Skip header line and process each entry
        tail -n +2 "$csvfile" | while IFS=, read -r email deviceid || [[ -n "$email" ]]; do
            # Remove any whitespace
            email=$(echo "$email" | xargs)
            deviceid=$(echo "$deviceid" | xargs)
            
            if [ -z "$deviceid" ]; then
                echo "Processing all devices for $email..."
                
                # Get all devices for this user and lock them
                gam print mobile query "user:$email" | tail -n +2 | while IFS=, read -r deviceid rest; do
                    if [ ! -z "$deviceid" ]; then
                        echo "Locking device $deviceid for $email..."
                        
                        if gam update mobile "$deviceid" action accountlock; then
                            echo "$email,$deviceid,success" >> "$resultfile"
                        else
                            echo "$email,$deviceid,failed" >> "$resultfile"
                        fi
                    fi
                done
            else
                echo "Locking specific device $deviceid..."
                
                if gam update mobile "$deviceid" action accountlock; then
                    echo "$email,$deviceid,success" >> "$resultfile"
                    echo "Device $deviceid locked successfully"
                else
                    echo "$email,$deviceid,failed" >> "$resultfile"
                    echo "Failed to lock device $deviceid"
                fi
            fi
        done
        
        echo "Device locking complete. Results saved to $resultfile"
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
        echo "Chrome device information:"
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

# Function to create CSV template
create_csv_template() {
    clear
    echo "Create CSV Template"
    echo "------------------"
    echo "1. User Creation Template"
    echo "2. Password Reset Template"
    echo "3. Archive Users Template"
    echo "4. Lock Devices Template"
    echo "5. Return to Main Menu"
    
    read -p "Select template type: " template_type
    
    case $template_type in
        1)  # User Creation Template
            echo "firstname,lastname,username,orgunit,password" > user_creation_template.csv
            echo "John,Doe,jdoe,/Staff," >> user_creation_template.csv
            echo "Jane,Smith,jsmith,/Faculty,TemporaryPwd123" >> user_creation_template.csv
            echo "Template saved as user_creation_template.csv"
            ;;
            
        2)  # Password Reset Template
            echo "email,password" > password_reset_template.csv
            echo "user1@example.com,NewPassword123" >> password_reset_template.csv
            echo "user2@example.com," >> password_reset_template.csv
            echo "Template saved as password_reset_template.csv"
            ;;
            
        3)  # Archive Users Template
            echo "email,admin" > archive_users_template.csv
            echo "user1@example.com,admin@example.com" >> archive_users_template.csv
            echo "user2@example.com,admin@example.com" >> archive_users_template.csv
            echo "Template saved as archive_users_template.csv"
            ;;
            
        4)  # Lock Devices Template
            echo "email,deviceid" > lock_devices_template.csv
            echo "user1@example.com," >> lock_devices_template.csv
            echo "user2@example.com,device_id_123" >> lock_devices_template.csv
            echo "Template saved as lock_devices_template.csv"
            ;;
            
        5)  # Return to Main Menu
            return
            ;;
            
        *)  echo "Invalid option"
            ;;
    esac
    
    read -p "Press Enter to continue..."
}

# Update main menu to include CSV templates option
show_menu() {
    echo "=========================================="
    echo "    GOOGLE WORKSPACE ADMIN TOOLKIT        "
    echo "=========================================="
    echo "1. Password Management"
    echo "2. User Management"
    echo "3. Device Management"
    echo "4. Information Retrieval"
    echo "5. Create CSV Templates"
    echo "6. Exit"
    echo "=========================================="
    echo "Enter your choice [1-6]: "
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
            
        5)  # Create CSV Templates
            create_csv_template
            ;;
            
        6)  # Exit
            echo "Exiting the Google Workspace Admin Toolkit. Goodbye!"
            exit 0
            ;;
            
        *)  # Invalid option
            echo "Invalid option. Press Enter to continue..."
            read
            ;;
    esac
done
