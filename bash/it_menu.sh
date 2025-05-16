#!/bin/bash

# Snipe-IT Management Menu
# Description: Interactive menu for managing Snipe-IT tasks

# Set terminal colors
BLUE='\033[0;34m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Set Snipe-IT API information - replace with your actual details
SNIPE_URL="https://inv.nomma.lan"
API_KEY="eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIzIiwianRpIjoiZDYxMzBlMzVkZWFkNWI0MWFkNjcxYTRkNDYwZjVlZWExNmViZGJkY2VkMTFiYWNmMDM1YzRjOWIyOTI5ZmFhZDE2MThmNzFhNGZlNDVjMGYiLCJpYXQiOjE3NDc0MDE3MTguMzg1MTM2LCJuYmYiOjE3NDc0MDE3MTguMzg1MTQzLCJleHAiOjIyMjA3ODczMTguMzc4MDE0LCJzdWIiOiIxIiwic2NvcGVzIjpbXX0.A7ncSeJ0_VZNK23MhEGEtny2ziDU081uCgnby8WsEHZSvPqBcdV5zsaLolmqq_FSVZcHqZG7uC2PmNxIhXroq8HGhkqhyldZpjTdYLSZlU8cvpFgLFxB0pTRgndG0p2GkY8nn3Jl_X8xUT6hG1xjH1jj50ANPSy47QlXGuOeDkZrBYB2W4Ed5FCAEDrsSBl3Isv-IMZPF-fakz_8CB9E7sbhtymttiYHhgTDH6Cf-9fsrXjIzg6yNT18QoscHmOjmAXI_tnHihsQWVPiXkxsOcRZhoFAMcTaqBHUD8XT7sxjcU5O32ICM_lE_bvk14yN0p7n3OzMJIPkfomZ7FuwowumVdpfOeIiIXWFI6BOvy8Er5iGCCJs-HT-4TaNDflRTYZHbataxmhPQZ1rcYU-c4XcZuhXaPSqr9WtIKFbkIAX9qSXWSda_3yy_HgAPIiMIa6vXJyj2fWpEaGQQroUc47o3qFaO-qKHt6JqCEB6MHyNNwhv9H1EOF0I3nOXY9GjW0fh5HmAdoCEB33PNJyG9TghOOhGQ4SY6xspc418Y7nw2Lto9hPm4vs-nN8fkQAsD3lQz50zLTwXY0e-ebviy5mpS0qC3iiEIRYhFDDgxgcagSC2CXKX6_CsthbL2Ck8IWd5zjBAXVKwL37KgqvcOJvqtlWxgRZDxgcvX5qurs"

# Function to display the main menu
show_main_menu() {
  clear
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${BLUE}       SNIPE-IT MANAGEMENT MENU         ${NC}"
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${YELLOW}1.${NC} Asset Management Automation"
  echo -e "${YELLOW}2.${NC} User Support & Identity Management"
  echo -e "${YELLOW}3.${NC} Routine Maintenance & Monitoring"
  echo -e "${YELLOW}4.${NC} Reporting"
  echo -e "${YELLOW}5.${NC} CSV Templates"
  echo -e "${YELLOW}0.${NC} Exit"
  echo -e "${BLUE}=========================================${NC}"
  echo -n "Enter your choice [0-5]: "
}

# Function for checking out assets with ID collection
checkout_asset() {
  clear
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${BLUE}          CHECKOUT ASSET TO USER         ${NC}"
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${GREEN}Enter checkout information. Type 'quit' to exit or 'back' to return to menu.${NC}"
  echo -e "${BLUE}=========================================${NC}"
  
  # Create or append to a collection file
  CHECKOUT_FILE="asset_checkout_$(date +%Y%m%d).csv"
  
  # If file doesn't exist, create it with headers
  if [ ! -f "$CHECKOUT_FILE" ]; then
    echo "ID Number,Device ID,Charger ID,Timestamp,Action" > "$CHECKOUT_FILE"
  fi
  
  while true; do
    echo -e "${BLUE}----------------------------------------${NC}"
    echo -n "Enter ID Number (or 'quit'/'back'): "
    read -r id_number
    
    # Check for quit or back commands
    if [[ "$id_number" == "quit" ]]; then
      echo -e "${GREEN}Exiting script. Checkout data saved to $CHECKOUT_FILE${NC}"
      exit 0
    elif [[ "$id_number" == "back" ]]; then
      echo -e "${GREEN}Returning to Asset Management menu. Checkout data saved to $CHECKOUT_FILE${NC}"
      return
    fi
    
    echo -n "Enter Device ID: "
    read -r device_id
    
    echo -n "Enter Charger ID: "
    read -r charger_id
    
    # Record the timestamp
    timestamp=$(date +"%Y-%m-%d %H:%M:%S")
    
    # Save to CSV file
    echo "$id_number,$device_id,$charger_id,$timestamp,CHECKOUT" >> "$CHECKOUT_FILE"
    
    echo -e "${GREEN}Checkout recorded successfully!${NC}"
    
    # Optional: Add API call to Snipe-IT here
    # curl -X POST -H "Authorization: Bearer $API_KEY" -H "Content-Type: application/json" -d '{"assigned_to": "'"$id_number"'"}' "$SNIPE_URL/api/v1/hardware/$device_id/checkout"
    
    echo -e "${BLUE}----------------------------------------${NC}"
  done
}

# Function for checking in assets with ID collection
checkin_asset() {
  clear
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${BLUE}          CHECKIN ASSET FROM USER        ${NC}"
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${GREEN}Enter checkin information. Type 'quit' to exit or 'back' to return to menu.${NC}"
  echo -e "${BLUE}=========================================${NC}"
  
  # Create or append to a collection file
  CHECKIN_FILE="asset_checkin_$(date +%Y%m%d).csv"
  
  # If file doesn't exist, create it with headers
  if [ ! -f "$CHECKIN_FILE" ]; then
    echo "ID Number,Device ID,Charger ID,Timestamp,Action" > "$CHECKIN_FILE"
  fi
  
  while true; do
    echo -e "${BLUE}----------------------------------------${NC}"
    echo -n "Enter ID Number (or 'quit'/'back'): "
    read -r id_number
    
    # Check for quit or back commands
    if [[ "$id_number" == "quit" ]]; then
      echo -e "${GREEN}Exiting script. Checkin data saved to $CHECKIN_FILE${NC}"
      exit 0
    elif [[ "$id_number" == "back" ]]; then
      echo -e "${GREEN}Returning to Asset Management menu. Checkin data saved to $CHECKIN_FILE${NC}"
      return
    fi
    
    echo -n "Enter Device ID: "
    read -r device_id
    
    echo -n "Enter Charger ID: "
    read -r charger_id
    
    # Record the timestamp
    timestamp=$(date +"%Y-%m-%d %H:%M:%S")
    
    # Save to CSV file
    echo "$id_number,$device_id,$charger_id,$timestamp,CHECKIN" >> "$CHECKIN_FILE"
    
    echo -e "${GREEN}Checkin recorded successfully!${NC}"
    
    # Optional: Add API call to Snipe-IT here
    # curl -X POST -H "Authorization: Bearer $API_KEY" -H "Content-Type: application/json" "$SNIPE_URL/api/v1/hardware/$device_id/checkin"
    
    echo -e "${BLUE}----------------------------------------${NC}"
  done
}

# Function for Asset Management Automation submenu
asset_management() {
  clear
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${BLUE}     ASSET MANAGEMENT AUTOMATION        ${NC}"
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${YELLOW}1.${NC} Add New Asset"
  echo -e "${YELLOW}2.${NC} Checkout Asset to User"
  echo -e "${YELLOW}3.${NC} Checkin Asset from User"
  echo -e "${YELLOW}4.${NC} Update Asset Information"
  echo -e "${YELLOW}5.${NC} Bulk Import Assets from CSV"
  echo -e "${YELLOW}6.${NC} Generate Asset Labels"
  echo -e "${YELLOW}7.${NC} Audit Assets"
  echo -e "${YELLOW}0.${NC} Return to Main Menu"
  echo -e "${BLUE}=========================================${NC}"
  echo -n "Enter your choice [0-7]: "
  
  read -r choice
  case $choice in
    1) echo "Adding new asset..." && sleep 2 ;;
    2) checkout_asset ;; # Call the checkout function
    3) checkin_asset ;; # Call the checkin function
    4) echo "Updating asset information..." && sleep 2 ;;
    5) echo "Importing assets from CSV..." && sleep 2 ;;
    6) echo "Generating asset labels..." && sleep 2 ;;
    7) echo "Auditing assets..." && sleep 2 ;;
    0) return ;;
    *) echo -e "${RED}Invalid option. Please try again.${NC}" && sleep 1 ;;
  esac
  asset_management
}

# Function for User Support & Identity Management submenu
user_management() {
  clear
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${BLUE}   USER SUPPORT & IDENTITY MANAGEMENT   ${NC}"
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${YELLOW}1.${NC} Add New User"
  echo -e "${YELLOW}2.${NC} Update User Information"
  echo -e "${YELLOW}3.${NC} Manage User Permissions"
  echo -e "${YELLOW}4.${NC} View User Assets"
  echo -e "${YELLOW}5.${NC} Bulk Import Users from CSV"
  echo -e "${YELLOW}6.${NC} Sync with LDAP/Active Directory"
  echo -e "${YELLOW}0.${NC} Return to Main Menu"
  echo -e "${BLUE}=========================================${NC}"
  echo -n "Enter your choice [0-6]: "
  
  read -r choice
  case $choice in
    1) echo "Adding new user..." && sleep 2 ;;
    2) echo "Updating user information..." && sleep 2 ;;
    3) echo "Managing user permissions..." && sleep 2 ;;
    4) echo "Viewing user assets..." && sleep 2 ;;
    5) echo "Importing users from CSV..." && sleep 2 ;;
    6) echo "Syncing with LDAP/AD..." && sleep 2 ;;
    0) return ;;
    *) echo -e "${RED}Invalid option. Please try again.${NC}" && sleep 1 ;;
  esac
  user_management
}

# Function for Routine Maintenance & Monitoring submenu
maintenance() {
  clear
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${BLUE}    ROUTINE MAINTENANCE & MONITORING    ${NC}"
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${YELLOW}1.${NC} Backup Snipe-IT Database"
  echo -e "${YELLOW}2.${NC} Check System Status"
  echo -e "${YELLOW}3.${NC} Run System Updates"
  echo -e "${YELLOW}4.${NC} Check for License Expirations"
  echo -e "${YELLOW}5.${NC} Maintenance Mode Toggle"
  echo -e "${YELLOW}6.${NC} View Error Logs"
  echo -e "${YELLOW}0.${NC} Return to Main Menu"
  echo -e "${BLUE}=========================================${NC}"
  echo -n "Enter your choice [0-6]: "
  
  read -r choice
  case $choice in
    1) echo "Backing up database..." && sleep 2 ;;
    2) echo "Checking system status..." && sleep 2 ;;
    3) echo "Running system updates..." && sleep 2 ;;
    4) echo "Checking license expirations..." && sleep 2 ;;
    5) echo "Toggling maintenance mode..." && sleep 2 ;;
    6) echo "Viewing error logs..." && sleep 2 ;;
    0) return ;;
    *) echo -e "${RED}Invalid option. Please try again.${NC}" && sleep 1 ;;
  esac
  maintenance
}

# Function for Reporting submenu
reporting() {
  clear
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${BLUE}              REPORTING                 ${NC}"
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${YELLOW}1.${NC} Asset Inventory Report"
  echo -e "${YELLOW}2.${NC} Asset Allocation Report"
  echo -e "${YELLOW}3.${NC} Depreciation Report"
  echo -e "${YELLOW}4.${NC} Maintenance History Report"
  echo -e "${YELLOW}5.${NC} License Compliance Report"
  echo -e "${YELLOW}6.${NC} Custom Report Builder"
  echo -e "${YELLOW}7.${NC} Export Reports to CSV/PDF"
  echo -e "${YELLOW}0.${NC} Return to Main Menu"
  echo -e "${BLUE}=========================================${NC}"
  echo -n "Enter your choice [0-7]: "
  
  read -r choice
  case $choice in
    1) echo "Generating asset inventory report..." && sleep 2 ;;
    2) echo "Generating asset allocation report..." && sleep 2 ;;
    3) echo "Generating depreciation report..." && sleep 2 ;;
    4) echo "Generating maintenance history report..." && sleep 2 ;;
    5) echo "Generating license compliance report..." && sleep 2 ;;
    6) echo "Opening custom report builder..." && sleep 2 ;;
    7) echo "Exporting reports..." && sleep 2 ;;
    0) return ;;
    *) echo -e "${RED}Invalid option. Please try again.${NC}" && sleep 1 ;;
  esac
  reporting
}

# Function for CSV Templates submenu
csv_templates() {
  clear
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${BLUE}            CSV TEMPLATES               ${NC}"
  echo -e "${BLUE}=========================================${NC}"
  echo -e "${YELLOW}1.${NC} Generate Asset Import Template"
  echo -e "${YELLOW}2.${NC} Generate User Import Template"
  echo -e "${YELLOW}3.${NC} Generate License Import Template"
  echo -e "${YELLOW}4.${NC} Generate Component Import Template"
  echo -e "${YELLOW}5.${NC} Generate Accessory Import Template"
  echo -e "${YELLOW}6.${NC} Generate Consumable Import Template"
  echo -e "${YELLOW}7.${NC} View CSV Import Documentation"
  echo -e "${YELLOW}0.${NC} Return to Main Menu"
  echo -e "${BLUE}=========================================${NC}"
  echo -n "Enter your choice [0-7]: "
  
  read -r choice
  case $choice in
    1) 
      echo "Generating asset import template..."
      cat > asset_import_template.csv << EOF
Asset Tag,Item Name,Status,Category,Manufacturer,Model,Serial,Purchase Date,Purchase Cost,Supplier,Notes,Location,Assigned To
A00001,MacBook Pro,Ready to Deploy,Laptop,Apple,MacBook Pro 16,C02XL0GZJGH5,2023-01-15,1999.99,Apple Store,New Device,Main Office,
A00002,ThinkPad T14,Deployed,Laptop,Lenovo,ThinkPad T14,PF2TCZX8,2023-02-10,1299.99,CDW,Standard Issue,Main Office,john.doe@company.com
EOF
      echo -e "${GREEN}Template saved as asset_import_template.csv${NC}"
      sleep 2 
      ;;
    2) 
      echo "Generating user import template..."
      cat > user_import_template.csv << EOF
First Name,Last Name,Username,Email,Employee Number,Department,Location,Manager,Notes,Activated
John,Doe,john.doe,john.doe@company.com,EMP001,IT,Main Office,jane.smith@company.com,Team Lead,1
Jane,Smith,jane.smith,jane.smith@company.com,EMP002,IT,Main Office,robert.johnson@company.com,Department Manager,1
EOF
      echo -e "${GREEN}Template saved as user_import_template.csv${NC}"
      sleep 2 
      ;;
    3)
      echo "Generating license import template..."
      cat > license_import_template.csv << EOF
Name,Product Key,Seats,Company,Manufacturer,Purchase Date,Purchase Cost,Expiration Date,Notes,Category
Microsoft 365,XXXXX-XXXXX-XXXXX-XXXXX-XXXXX,50,Microsoft,Microsoft,2023-01-01,10000.00,2024-01-01,Annual Subscription,Software
Adobe Creative Cloud,YYYYY-YYYYY-YYYYY-YYYYY-YYYYY,10,Adobe,Adobe,2023-02-15,5999.99,2024-02-15,Design Team License,Software
EOF
      echo -e "${GREEN}Template saved as license_import_template.csv${NC}"
      sleep 2 
      ;;
    4) 
      echo "Generating component import template..."
      cat > component_import_template.csv << EOF
Name,Category,Quantity,Min Qty,Location,Order Number,Purchase Date,Purchase Cost,Notes
RAM 16GB DDR4,Memory,20,5,Main Office,PO-12345,2023-01-15,1999.99,For laptop upgrades
SSD 1TB,Storage,15,3,Main Office,PO-12346,2023-01-15,2499.99,For laptop upgrades
EOF
      echo -e "${GREEN}Template saved as component_import_template.csv${NC}"
      sleep 2 
      ;;
    5) 
      echo "Generating accessory import template..."
      cat > accessory_import_template.csv << EOF
Name,Category,Quantity,Min Qty,Location,Purchase Date,Purchase Cost,Supplier,Notes
Laptop Bag,Bags,30,5,Main Office,2023-01-15,899.99,CDW,Standard issue
Wireless Mouse,Peripherals,25,10,Main Office,2023-01-15,499.99,Dell,Standard issue
EOF
      echo -e "${GREEN}Template saved as accessory_import_template.csv${NC}"
      sleep 2 
      ;;
    6) 
      echo "Generating consumable import template..."
      cat > consumable_import_template.csv << EOF
Name,Category,Quantity,Min Qty,Location,Purchase Date,Purchase Cost,Model Number,Notes
Printer Toner,Printing Supplies,15,3,Main Office,2023-01-15,899.99,HP-CF410X,For HP LaserJet Pro M452dn
Copy Paper,Office Supplies,50,10,Main Office,2023-01-15,199.99,COPY-A4,A4 80gsm
EOF
      echo -e "${GREEN}Template saved as consumable_import_template.csv${NC}"
      sleep 2 
      ;;
    7) echo "Opening CSV import documentation..." && sleep 2 ;;
    0) return ;;
    *) echo -e "${RED}Invalid option. Please try again.${NC}" && sleep 1 ;;
  esac
  csv_templates
}

# Main program loop
while true; do
  show_main_menu
  read -r choice
  
  case $choice in
    1) asset_management ;;
    2) user_management ;;
    3) maintenance ;;
    4) reporting ;;
    5) csv_templates ;;
    0) 
      clear
      echo -e "${GREEN}Thank you for using the Snipe-IT Management Menu.${NC}"
      echo -e "${GREEN}Goodbye!${NC}"
      exit 0 
      ;;
    *) echo -e "${RED}Invalid option. Please try again.${NC}" && sleep 1 ;;
  esac
done
