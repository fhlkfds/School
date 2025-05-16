#!/bin/bash

# Google Workspace Admin Toolkit using GAM
# Created: May 14, 2025
# Requirements: GAM must be installed and configured


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
    echo "6. Moving User to a grade OU (CSV)"
    echo "7. Return to Main Menu"
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
    echo "3. Re-enable Single Mobile Device"
    echo "4. Re-enable Multiple Mobile Devices (CSV)"
    echo "5. Wipe Device"
    echo "6. List All Devices"
    echo "7. Return to Main Menu"
    echo "=========================================="
    echo "Enter your choice [1-7]: "
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

mass_reset_password() {
    clear
    echo "Mass Password Reset (CSV)"
    echo "------------------------"
    echo "CSV file should have a single column with user emails"
    echo "All passwords will be reset to: Password@1"
    
    read -p "Enter CSV file path: " csvfile
    
    if [ -f "$csvfile" ]; then
        echo "Resetting passwords for users in $csvfile..."
        
        # Create output file for results
        resultfile="password_reset_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "email,new_password,status" > "$resultfile"
        
        # Set default password
        default_password="Password@1"
        
        # Process each line in the CSV file
        while IFS= read -r line || [[ -n "$line" ]]; do
            # Skip empty lines and possible header
            if [[ -z "$line" || "$line" == "email" ]]; then
                continue
            fi
            
            # Remove any whitespace
            email=$(echo "$line" | xargs)
            
            echo "Processing $email..."
            
            # Reset the password to Password@1
            if gam update user "$email" password "$default_password"; then
                echo "$email,$default_password,success" >> "$resultfile"
                echo "Password reset successful for $email"
            else
                echo "$email,$default_password,failed" >> "$resultfile"
                echo "Password reset failed for $email"
            fi
        done < "$csvfile"
        
        echo "Password reset complete. Results saved to $resultfile"
    else
        echo "File not found!"
    fi
    read -p "Press Enter to continue..."
}
# Function to create a single user with grade/job title-based OU assignment
create_user() {
    clear
    echo "Create Single User"
    echo "-----------------"
    read -p "Enter first name: " firstname
    read -p "Enter last name: " lastname
    
    # Ask if user is a student or staff
    echo "Is this user a student or staff?"
    echo "1. Student"
    echo "2. Staff/Adult"
    read -p "Enter choice (1 or 2): " usertype
    
    # Initialize orgunit and email format
    orgunit=""
    
    if [ "$usertype" = "1" ]; then
        # Student - ask for grade level
        echo "Select student grade level:"
        echo "8. Grade 8"
        echo "9. Grade 9"
        echo "10. Grade 10"
        echo "11. Grade 11"
        echo "12. Grade 12"
        read -p "Enter grade level (8-12): " grade
        
        case $grade in
            8) orgunit="/Cadets/8th grade/" ;;
            9) orgunit="/Cadets/9th grade/" ;;
            10) orgunit="/Cadets/10th grade/" ;;
            11) orgunit="/Cadets/11th grade/" ;;
            12) orgunit="/Cadets/12th grade/" ;;
            *) echo "Invalid grade level. Using default /Cadets"; orgunit="/Cadets" ;;
        esac
        
        # For students, format email as caMoving User to a grade OU (CSV)detflastname@nomma.net
        firstinitial="${firstname:0:1}"
        firstinitial=$(echo "$firstinitial" | tr '[:upper:]' '[:lower:]')  # convert to lowercase
        lastname_lower=$(echo "$lastname" | tr '[:upper:]' '[:lower:]')   # convert to lowercase
        username="cadet${firstinitial}${lastname_lower}"
        domain="@nomma.net"
    elif [ "$usertype" = "2" ]; then
        # Staff/Adult - ask for job title
        echo "Select staff job title:"
        echo "1. Teacher"
        echo "2. IT"
        echo "3. Staff (General)"
        echo "4. Security"
        echo "5. Counselor"
        echo "6. School Admin"
        read -p "Enter choice (1-6): " jobtitle
        
        case $jobtitle in
            1) orgunit="/Faculty & Staff/Teachers" ;;
            2) orgunit="/Faculty & Staff/IT" ;;
            3) orgunit="/Faculty & Staff/General" ;;
            4) orgunit="/Faculty & Staff/Security" ;;
            5) orgunit="/Faculty & Staff/Counselors" ;;
            6) orgunit="/Faculty & Staff/SchoolAdmins" ;;
            *) echo "Invalid selection. Using default /Faculty & Staff"; orgunit="/Faculty & Staff" ;;
        esac
        
        # For staff, ask for username
        read -p "Enter username (before @nomma.net): " username
        domain="@nomma.net"
    else
        echo "Invalid selection. Using root organizational unit."
        read -p "Enter username (before @nomma.net): " username
        domain="@nomma.net"
    fi
    
    email="${username}${domain}"
    
    # Set default password
    password="Password@1"
    
    echo "Creating user $email in organization unit $orgunit..."
    
    cmd="gam create user $email firstname \"$firstname\" lastname \"$lastname\" password \"$password\" org \"$orgunit\""
    
    eval $cmd
    
    echo "User $email has been created in $orgunit with password: $password"
    read -p "Press Enter to continue..."
}

# Function to create multiple users from CSV with grade/job title-based OU assignment
create_multi_user() {
    clear
    echo "Create Multiple Users (CSV)"
    echo "-------------------------"
    echo "CSV file should have headers: firstname,lastname,usertype,grade_or_jobtitle"
    echo "- usertype: 'student' or 'staff'"
    echo "- grade_or_jobtitle: For students: 8-12, For staff: Teacher, IT, Staff, Security, Counselor, SchoolAdmin"
    echo "- For staff, optional username column can be provided, otherwise first initial and last name will be used"
    
    read -p "Enter CSV file path: " csvfile
    
    if [ -f "$csvfile" ]; then
        echo "Creating users from $csvfile..."
        
        # Create a results file
        resultfile="user_creation_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "firstname,lastname,email,orgunit,password,status" > "$resultfile"
        
        # Default password for all users
        default_password="Password@1"
        
        # Skip header line and process each user
        tail -n +2 "$csvfile" | while IFS=, read -r firstname lastname usertype grade_or_jobtitle username || [[ -n "$firstname" ]]; do
            # Remove any whitespace
            firstname=$(echo "$firstname" | xargs)
            lastname=$(echo "$lastname" | xargs)
            usertype=$(echo "$usertype" | tr '[:upper:]' '[:lower:]' | xargs)
            grade_or_jobtitle=$(echo "$grade_or_jobtitle" | xargs)
            username=$(echo "$username" | xargs)
            
            # Determine organizational unit based on usertype and grade/jobtitle
            if [ "$usertype" = "student" ]; then
                case $grade_or_jobtitle in
                    8) orgunit="/Faculty & Staff/New Staff/" ;;
                    9) orgunit="/Faculty &  Staff/New Staff/" ;;
                    10) orgunit="/Faculty & Staff/New Staff/" ;;
                    11) orgunit="/Faculty & Staff/New Staff/" ;;
                    12) orgunit="/Faculty & Staff/New Staff/" ;;
                    *) orgunit="/Faculty & Staff" ;;
                esac
                
                # For students, format email as cadetflastname@nomma.net
                firstinitial="${firstname:0:1}"
                firstinitial=$(echo "$firstinitial" | tr '[:upper:]' '[:lower:]')  # convert to lowercase
                lastname_lower=$(echo "$lastname" | tr '[:upper:]' '[:lower:]')   # convert to lowercase
                username="cadet${firstinitial}${lastname_lower}"
                domain="@nomma.net"
            elif [ "$usertype" = "staff" ]; then
                case $grade_or_jobtitle in
                    "Teacher") orgunit="/Faculty & Staff/New Staff" ;;
                    "IT") orgunit="/Faculty & Staff/New Staff" ;;
                    "Staff") orgunit="/Faculty & Staff/New Staff" ;;
                    "Security") orgunit="/Faculty & Staff/New Staff" ;;
                    "Counselor") orgunit="/Faculty & Staff/New Staff" ;;
                    "SchoolAdmin") orgunit="/Faculty & Staff/New Staff" ;;
                    *) orgunit="/Faculty & Staff" ;;
                esac
                
                # For staff, use provided username or generate one
                if [ -z "$username" ]; then
                    firstinitial="${firstname:0:1}"
                    firstinitial=$(echo "$firstinitial" | tr '[:upper:]' '[:lower:]')
                    lastname_lower=$(echo "$lastname" | tr '[:upper:]' '[:lower:]')
                    username="${firstinitial}${lastname_lower}"
                fi
                domain="@nomma.net"
            else
                orgunit="/"
                # Default email formatting if usertype is unknown
                if [ -z "$username" ]; then
                    firstinitial="${firstname:0:1}"
                    firstinitial=$(echo "$firstinitial" | tr '[:upper:]' '[:lower:]')
                    lastname_lower=$(echo "$lastname" | tr '[:upper:]' '[:lower:]')
                    username="${firstinitial}${lastname_lower}"
                fi
                domain="@nomma.net"
            fi
            
            email="${username}${domain}"
            
            echo "Creating user $email in $orgunit..."
            
            cmd="gam create user $email firstname \"$firstname\" lastname \"$lastname\" password \"$default_password\" org \"$orgunit\""
            
            if eval $cmd; then
                echo "$firstname,$lastname,$email,$orgunit,$default_password,success" >> "$resultfile"
                echo "User creation successful for $email in $orgunit"
            else
                echo "$firstname,$lastname,$email,$orgunit,$default_password,failed" >> "$resultfile"
                echo "User creation failed for $email"
            fi
        done
        
        echo "User creation complete. Results saved to $resultfile"
    else
        echo "File not found!"
    fi
    read -p "Press Enter to continue..."
}

# Function to create CSV template for user creation with grade/jobtitle
create_csv_template() {
    clear
    echo "Create CSV Template"
    echo "------------------"
    echo "1. User Creation Template (with Grade/Job Title)"
    echo "2. Password Reset Template"
    echo "3. Archive Users Template"
    echo "4. Lock Devices Template"
    echo "5. Return to Main Menu"
    
    read -p "Select template type: " template_type
    
    case $template_type in
        1)  # User Creation Template with Grade/Job Title
            echo "firstname,lastname,usertype,grade_or_jobtitle,username" > user_creation_template.csv
            echo "John,Doe,student,9," >> user_creation_template.csv
            echo "Jane,Smith,staff,Teacher,jsmith" >> user_creation_template.csv
            echo "Mike,Johnson,staff,IT," >> user_creation_template.csv
            echo "Template saved as user_creation_template.csv"
            echo "Note: usertype should be 'student' or 'staff'"
            echo "For students: grade_or_jobtitle should be 8-12"
            echo "  - Students will be placed in /cadet/[grade]th grade/"
            echo "  - Student emails will be automatically formatted as cadetflastname@nomma.net"
            echo "For staff: grade_or_jobtitle should be Teacher, IT, Staff, Security, Counselor, or SchoolAdmin"
            echo "  - Staff will be placed in /Faculty & Staff/[JobTitle]"
            echo "  - Staff emails will use the username column (if provided) or first initial + lastname"
            echo "  - Staff domain is @nomma.net"
            ;;
            
        2)  # Password Reset Template
            echo "email" > password_reset_template.csv
            echo "cadetjdoe@nomma.net" >> password_reset_template.csv
            echo "jsmith@nomma.net" >> password_reset_template.csv
            echo "Template saved as password_reset_template.csv"
            echo "Note: All passwords will be reset to Password@1"
            ;;
            
        3)  # Archive Users Template
            echo "email,admin" > archive_users_template.csv
            echo "cadetjdoe@nomma.net,admin@nomma.net" >> archive_users_template.csv
            echo "jsmith@nomma.net,admin@nomma.net" >> archive_users_template.csv
            echo "Template saved as archive_users_template.csv"
            ;;
            
        4)  # Lock Devices Template
            echo "email,deviceid" > lock_devices_template.csv
            echo "cadetjdoe@nomma.net," >> lock_devices_template.csv
            echo "jsmith@nomma.net,device_id_123" >> lock_devices_template.csv
            echo "Template saved as lock_devices_template.csv"
            ;;
        5)  # Grade Movement Template
            echo "email,grade" > grade_movement_template.csv
            echo "cadetjdoe@nomma.net,9" >> grade_movement_template.csv
            echo "cadetasmith@nomma.net,10" >> grade_movement_template.csv
            echo "Template saved as grade_movement_template.csv"
            echo "Note: grade should be 8, 9, 10, 11, or 12"
            echo "  - Users will be moved to /Cadets/[grade]th grade/"
            ;;

            
        6)  # Return to Main Menu
            return
            ;;
            
        *)  echo "Invalid option"
            ;;
    esac
    
    read -p "Press Enter to continue..."
}




# Function to move users to grade OUs based on their emails
move_users_to_grade() {
    clear
    echo "Move Users to Grade OU (CSV)"
    echo "-------------------------"
    echo "CSV file should have headers: email,grade"
    echo "- email: User's email address"
    echo "- grade: Grade level (8, 9, 10, 11, or 12)"
    
    read -p "Enter CSV file path: " csvfile
    
    if [ -f "$csvfile" ]; then
        echo "Moving users from $csvfile to appropriate grade OUs..."
        
        # Create output file for results
        resultfile="grade_move_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "email,grade,new_ou,status" > "$resultfile"
        
        # Skip header line and process each user
        tail -n +2 "$csvfile" | while IFS=, read -r email grade || [[ -n "$email" ]]; do
            # Remove any whitespace
            email=$(echo "$email" | xargs)
            grade=$(echo "$grade" | xargs)
            
            # Skip if email or grade is empty after trimming
            if [ -z "$email" ] || [ -z "$grade" ]; then
                continue
            fi
            
            # Determine the appropriate OU based on grade
            case $grade in
                8) ou="/Cadets/8th grade/" ;;
                9) ou="/Cadets/9th grade/" ;;
                10) ou="/Cadets/10th grade/" ;;
                11) ou="/Cadets/11th grade/" ;;
                12) ou="/Cadets/12th grade/" ;;
                *) 
                    echo "Invalid grade level for $email: $grade. Skipping."
                    echo "$email,$grade,invalid,skipped" >> "$resultfile"
                    continue
                    ;;
            esac
            
            echo "Moving $email to $ou..."
            
            if gam update user "$email" org "$ou"; then
                echo "$email,$grade,$ou,success" >> "$resultfile"
                echo "Successfully moved $email to $ou"
            else
                echo "$email,$grade,$ou,failed" >> "$resultfile"
                echo "Failed to move $email to $ou"
            fi
        done
        
        echo "User movement complete. Results saved to $resultfile"
    else
        echo "File not found!"
    fi
    read -p "Press Enter to continue..."
}

# Update to create_csv_template function to add grade movement template
# Add this case to the existing create_csv_template function's case statement:
#        6)  # Grade Movement Template
#            echo "email,grade" > grade_movement_template.csv
#            echo "cadetjdoe@nomma.net,9" >> grade_movement_template.csv
#            echo "cadetasmith@nomma.net,10" >> grade_movement_template.csv
#            echo "Template saved as grade_movement_template.csv"
#            echo "Note: grade should be 8, 9, 10, 11, or 12"
#            echo "  - Users will be moved to /Cadets/[grade]th grade/"
#            ;;
# Function to archive a single user
archive_user() {
    clear
    echo "Archive User"
    echo "------------"
    read -p "Enter user email to archive: " useremail
    
    # Extract username and domain parts
    username=$(echo "$useremail" | cut -d'@' -f1)
    domain=$(echo "$useremail" | cut -d'@' -f2)
    
    echo "Archiving user $useremail..."
    echo "This will:"
    echo "1. Suspend the account"
    echo "2. Add 'inactive-' prefix to username"
    echo "3. Reset password to random string"
    echo "4. Replace aliases with random string"
    echo "5. Move user to Inactive OU"
    
    read -p "Continue? (y/n): " confirm
    if [[ $confirm == [Yy]* ]]; then
        # No data backup or transfer
        
        echo "Suspending account..."
        gam update user $useremail suspended on
        
        # Generate random string for password and alias
        random_string=$(cat /dev/urandom | tr -dc 'a-zA-Z0-9' | fold -w 12 | head -n 1)
        
        # Generate random password (more complex)
        random_password=$(cat /dev/urandom | tr -dc 'a-zA-Z0-9!@#$%^&*()_+' | fold -w 16 | head -n 1)
        
        # Rename user to add "inactive-" prefix
        new_username="inactive-$username"
        new_email="$new_username@$domain"
        
        
        echo "Resetting password to random string..."
        gam update user $new_email password "$random_password"
        
        echo "Removing aliases and setting to random value..."
        # First, get all aliases
        gam user $new_email print aliases > tmp_aliases.txt
        
        # Remove all existing aliases if any exist
        if [ -s tmp_aliases.txt ]; then
            tail -n +2 tmp_aliases.txt | while read alias; do
                echo "Removing alias: $alias"
                gam user $new_email delete alias $alias
            done
        fi
        
        # Add a single random alias
        random_alias="archived-${random_string}@$domain"
        echo "Adding random alias: $random_alias"
        gam user $new_email add alias $random_alias
        
        # Remove temporary file
        rm -f tmp_aliases.txt
        
        echo "Moving user to Inactive OU..."
        gam update user $new_email org "/Inactive"

        echo "Renaming user to $new_email..."
        gam update user $useremail primaryemail $new_email
       
        echo "User archiving complete:"
        echo "- Original email: $useremail"
        echo "- New email: $new_email"
        echo "- Random alias: $random_alias"
        echo "- Account suspended: Yes"
        echo "- Password reset: Yes"
        echo "- Moved to Inactive OU: Yes"
        
        # Save the information to a log file
        log_file="archive_log_$(date +%Y%m%d_%H%M%S).txt"
        echo "Archive log for $useremail" > "$log_file"
        echo "Timestamp: $(date)" >> "$log_file"
        echo "New email: $new_email" >> "$log_file"
        echo "Random alias: $random_alias" >> "$log_file"
        echo "Log saved to $log_file"
    else
        echo "Operation cancelled"
    fi
    
    read -p "Press Enter to continue..."
}
# Function to archive multiple users using CSV
# Fixed archive_multi_user function with username change as the last step and alias removal
# Fixed archive_multi_user function with username change as the last step and proper alias removal
archive_multi_user() {
    clear
    echo "Archive Multiple Users (CSV)"
    echo "--------------------------"
    echo "CSV file should have headers: email"
    echo "- email: The user email to archive"
    
    read -p "Enter CSV file path: " csvfile
    
    if [ -f "$csvfile" ]; then
        echo "Archiving users from $csvfile..."
        
        # Create output file for results
        resultfile="archive_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "original_email,new_email,status" > "$resultfile"
        
        read -p "This will archive all listed users. Continue? (y/n): " confirm
        if [[ $confirm == [Yy]* ]]; then
            # Skip header line and process each user
            tail -n +2 "$csvfile" | while IFS=, read -r email || [[ -n "$email" ]]; do
                # Remove any whitespace and skip empty lines
                email=$(echo "$email" | xargs)
                
                # Skip if email is empty after trimming
                if [ -z "$email" ]; then
                    continue
                fi
                
                echo "Archiving $email..."
                
                # Extract username and domain parts for later use
                username=$(echo "$email" | cut -d'@' -f1)
                domain=$(echo "$email" | cut -d'@' -f2)
                
                # Generate random password (more complex)
                random_password=$(cat /dev/urandom | tr -dc 'a-zA-Z0-9!@#$%^&*()_+' | fold -w 16 | head -n 1)
                
                # Calculate new email for later (but don't change it yet)
                new_username="inactive-$username"
                new_email="$new_username@$domain"
                
                # Step 1: Suspend the account
                if gam update user $email suspended on; then
                    echo "Successfully suspended $email"
                    
                    # Step 2: Reset password
                    if gam update user $email password "$random_password"; then
                        echo "Successfully reset password for $email"
                        
                        # Step 3: Remove all aliases - using the correct GAM syntax
                        echo "Removing all aliases for $email..."
                        if gam user $email delete aliases; then
                            echo "Successfully removed all aliases for $email"
                        else
                            echo "Note: No aliases to remove or failed to remove aliases for $email"
                            # Continue anyway since this is not critical
                        fi
                        
                        # Step 4: Move to Inactive OU (still using original email)
                        if gam update user $email org "/Inactive"; then
                            echo "Successfully moved $email to Inactive OU"
                            
                            # Step 5: Finally, rename the user as the LAST step
                            if gam update user $email primaryemail $new_email; then
                                echo "Successfully renamed to $new_email"
                                echo "$email,$new_email,success" >> "$resultfile"
                                echo "Archive successful for $email -> $new_email"
                            else
                                echo "$email,failed,failed-rename" >> "$resultfile"
                                echo "Archive partially complete, but failed to rename $email to $new_email"
                            fi
                        else
                            echo "$email,failed,failed-move-to-ou" >> "$resultfile"
                            echo "Failed to move $email to Inactive OU"
                        fi
                    else
                        echo "$email,failed,failed-reset-password" >> "$resultfile"
                        echo "Failed to reset password for $email"
                    fi
                else
                    echo "$email,failed,failed-suspend" >> "$resultfile"
                    echo "Failed to suspend $email"
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

# Function to re-enable a single Chrome OS device
enable_device() {
    clear
    echo "Re-enable Chrome OS Device"
    echo "-------------------------"
    read -p "Enter asset tag: " assettag
    
    echo "Retrieving Chrome OS device with asset ID $assettag..."
    # Find Chrome OS devices with the specified asset ID
    gam print cros query "asset_id:$assettag" > devices.txt
    
    if [ -s devices.txt ]; then
        echo "Available devices:"
        cat devices.txt | awk -F',' '{print NR ") " $1 " - " $2 " - " $5}'
        
        read -p "Enter device number to re-enable: " devicenum
        
        deviceid=$(sed -n "${devicenum}p" devices.txt | cut -d',' -f1)
        
        if [ ! -z "$deviceid" ]; then
            echo "Re-enabling Chrome OS device $deviceid..."
            gam update cros "$deviceid" action reenable
            echo "Device has been re-enabled"
        else
            echo "Invalid device selection"
        fi
    else
        echo "No Chrome OS devices found with asset ID $assettag"
    fi
    
    rm devices.txt 2>/dev/null
    read -p "Press Enter to continue..."
}

# Function to re-enable multiple Chrome OS devices using CSV
enable_multi_device() {
    clear
    echo "Re-enable Multiple Chrome OS Devices (CSV)"
    echo "----------------------------------------"
    echo "CSV file should have headers: assettag,deviceid"
    echo "- assettag: Device's asset ID (if deviceid not provided)"
    echo "- deviceid: Optional - specific device ID to re-enable"
    echo "If deviceid is blank, devices with that asset ID will be found and re-enabled"
    
    read -p "Enter CSV file path: " csvfile
    
    if [ -f "$csvfile" ]; then
        echo "Re-enabling Chrome OS devices from $csvfile..."
        
        # Create output file for results
        resultfile="device_enable_results_$(date +%Y%m%d_%H%M%S).csv"
        echo "assettag,deviceid,status" > "$resultfile"
        
        # Skip header line and process each entry
        tail -n +2 "$csvfile" | while IFS=, read -r assettag deviceid || [[ -n "$assettag" ]]; do
            # Remove any whitespace
            assettag=$(echo "$assettag" | xargs)
            deviceid=$(echo "$deviceid" | xargs)
            
            if [ -z "$deviceid" ]; then
                echo "Finding Chrome OS devices with asset ID $assettag..."
                
                # Find Chrome OS devices with matching asset ID
                gam print cros query "asset_id:$assettag" > matching_devices.txt
                
                if [ -s matching_devices.txt ]; then
                    # Process each matching device
                    tail -n +2 matching_devices.txt | while IFS=, read -r deviceid rest; do
                        if [ ! -z "$deviceid" ]; then
                            echo "Re-enabling Chrome OS device $deviceid with asset ID $assettag..."
                            
                            if gam update cros "$deviceid" action reenable; then
                                echo "$assettag,$deviceid,success" >> "$resultfile"
                            else
                                echo "$assettag,$deviceid,failed" >> "$resultfile"
                            fi
                        fi
                    done
                else
                    echo "No Chrome OS devices found with asset ID $assettag"
                    echo "$assettag,,not_found" >> "$resultfile"
                fi
                
                rm matching_devices.txt 2>/dev/null
            else
                echo "Re-enabling specific Chrome OS device $deviceid..."
                
                if gam update cros "$deviceid" action reenable; then
                    echo "$assettag,$deviceid,success" >> "$resultfile"
                    echo "Chrome OS device $deviceid re-enabled successfully"
                else
                    echo "$assettag,$deviceid,failed" >> "$resultfile"
                    echo "Failed to re-enable Chrome OS device $deviceid"
                fi
            fi
        done
        
        echo "Chrome OS device re-enabling complete. Results saved to $resultfile"
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
    echo "1. User Creation Template (with Grade/Job Title)"
    echo "2. Password Reset Template"
    echo "3. Archive Users Template"
    echo "4. Lock Devices Template"
    echo "5. Return to Main Menu"
    
    read -p "Select template type: " template_type
    
    case $template_type in
        1)  # User Creation Template with Grade/Job Title
            echo "firstname,lastname,usertype,grade_or_jobtitle,username" > user_creation_template.csv
            echo "John,Doe,student,9," >> user_creation_template.csv
            echo "Jane,Smith,staff,Teacher,jsmith" >> user_creation_template.csv
            echo "Mike,Johnson,staff,IT," >> user_creation_template.csv
            echo "Template saved as user_creation_template.csv"
            echo "Note: usertype should be 'student' or 'staff'"
            echo "For students: grade_or_jobtitle should be 8-12"
            echo "  - Students will be placed in /cadet/[grade]th grade/"
            echo "  - Student emails will be automatically formatted as cadetflastname@nomma.net"
            echo "For staff: grade_or_jobtitle should be Teacher, IT, Staff, Security, Counselor, or SchoolAdmin"
            echo "  - Staff will be placed in /Faculty & Staff/[JobTitle]"
            echo "  - Staff emails will use the username column (if provided) or first initial + lastname"
            echo "  - Staff domain is @nomma.net"
            ;;
            
        2)  # Password Reset Template
            echo "email" > password_reset_template.csv
            echo "cadetjdoe@nomma.net" >> password_reset_template.csv
            echo "jsmith@nomma.net" >> password_reset_template.csv
            echo "Template saved as password_reset_template.csv"
            echo "Note: All passwords will be reset to Password@1"
            ;;
            
        3)  # Archive Users Template
            echo "email" > archive_users_template.csv
            echo "cadetjdoe@nomma.net" >> archive_users_template.csv
            echo "jsmith@nomma.net" >> archive_users_template.csv
            echo "Template saved as archive_users_template.csv"
            echo "Note: This template matches the format required by the Archive Multiple Users function"
            ;;
            
        4)  # Lock Devices Template
            echo "Asset ID" > lock_devices_template.csv
            echo "0001" >> lock_devices_template.csv
            echo "1000" >> lock_devices_template.csv
            echo "Template saved as lock_devices_template.csv"
            echo "Note: Enter the device's asset tag or serial number for each device to be locked"
            echo "  - Each line should contain exactly one device identifier"
            echo "  - All devices in this list will be locked"
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
                    6) move_users_to_grade ;;
                    7) break ;;
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
                    3) enable_device ;;
                    4) enable_multi_device ;;
                    5) wipe_device ;;
                    6) list_devices ;;
                    7) break ;;
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
