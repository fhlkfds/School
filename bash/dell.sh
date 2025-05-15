#!/bin/bash

# Dell Warranty Ticket Generator Script
# This script creates Dell warranty tickets in a loop until the user types "quit"

# Colors for better readability
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# Function to display script banner
show_banner() {
    echo -e "${BLUE}╔════════════════════════════════════════════════════════╗${NC}"
    echo -e "${BLUE}║               ${GREEN}Dell Warranty Ticket Generator${BLUE}              ║${NC}"
    echo -e "${BLUE}╚════════════════════════════════════════════════════════╝${NC}"
    echo -e "${YELLOW}Type 'quit' at any prompt to exit the program${NC}"
    echo ""
}

# Function to display available issue types
show_issue_types() {
    echo -e "${CYAN}Available issue types:${NC}"
    echo -e "  ${GREEN}lcd${NC}         - LCD screen damage"
    echo -e "  ${GREEN}liquid${NC}      - Liquid spill damage"
    echo -e "  ${GREEN}power${NC}       - Power surge damage"
    echo -e "  ${GREEN}hinges${NC}      - Broken screen hinges"
    echo -e "  ${GREEN}casing${NC}      - Damaged laptop casing"
    echo -e "  ${GREEN}keyboard${NC}    - Keyboard issues"
    echo -e "  ${GREEN}touchpad${NC}    - Touchpad problems"
    echo -e "  ${GREEN}battery${NC}     - Battery not holding charge"
    echo -e "  ${GREEN}port${NC}        - Damaged ports"
    echo ""
}

# Function to get serial number from Snipe-IT using the asset tag
get_serial_number() {
    local asset_tag=$1
    echo -e "${YELLOW}Looking up serial number for asset tag ${asset_tag}...${NC}"
    
    # Replace with your actual Snipe-IT API call
    # This is a placeholder - you'll need to add your actual API endpoint and credentials
    SNIPEIT_API_KEY="your_snipeit_api_key_here"
    SNIPEIT_URL="https://your-snipeit-instance.com/api/v1"
    
    # Make API call to Snipe-IT (this is a placeholder)
    RESPONSE=$(curl -s -H "Authorization: Bearer $SNIPEIT_API_KEY" \
                    -H "Accept: application/json" \
                    "${SNIPEIT_URL}/hardware/bytag/${asset_tag}")
    
    # Extract serial number from the response (adjust jq parsing as needed)
    SERIAL_NUMBER=$(echo "$RESPONSE" | jq -r '.serial')
    
    # Check if serial number was found
    if [ "$SERIAL_NUMBER" == "null" ] || [ -z "$SERIAL_NUMBER" ]; then
        echo -e "${RED}Error: Could not find serial number for asset tag ${asset_tag}${NC}"
        return 1
    fi
    
    echo -e "${GREEN}Found serial number: ${SERIAL_NUMBER}${NC}"
    echo "$SERIAL_NUMBER"
}

# Function to get specific issue description
get_specific_issue() {
    show_specific_issues
    
    echo -e "${CYAN}Enter specific issue number or 'quit' to exit:${NC}"
    read -r ISSUE_OPTION
    
    # Check if user wants to quit
    if [[ "$ISSUE_OPTION" == "quit" ]]; then
        echo -e "${GREEN}Exiting Dell Warranty Ticket Generator. Goodbye!${NC}"
        exit 0
    fi
    
    local issue_description=""
    
    case "$ISSUE_OPTION" in
        1)
            issue_description="Multiple keys on the keyboard are not functioning. The keyboard needs replacement."
            ;;
        2)
            issue_description="The screen is damaged (cracked/displays lines/blank). The screen needs replacement."
            ;;
        3)
            issue_description="The touchpad/mouse is not responding correctly. It needs repair or replacement."
            ;;
        4)
            issue_description="The battery is not holding a charge. The laptop only works when plugged in."
            ;;
        5)
            issue_description="The operating system is not booting correctly. System recovery is needed."
            ;;
        6)
            echo -e "${CYAN}Which internal component is affected? (CPU, RAM, motherboard, etc.):${NC}"
            read -r COMPONENT
            
            # Check if user wants to quit
            if [[ "$COMPONENT" == "quit" ]]; then
                echo -e "${GREEN}Exiting Dell Warranty Ticket Generator. Goodbye!${NC}"
                exit 0
            fi
            issue_description="There is an issue with the $COMPONENT. It needs diagnosis and possible replacement."
            ;;
        7)
            issue_description="One or more ports (USB/HDMI/etc.) are not functioning properly. They need repair."
            ;;
        8)
            issue_description="The audio system is not working correctly. Speakers or audio components need repair."
            ;;
        9)
            echo -e "${CYAN}Enter custom issue description or 'quit' to exit:${NC}"
            read -r issue_description
            
            # Check if user wants to quit
            if [[ "$issue_description" == "quit" ]]; then
                echo -e "${GREEN}Exiting Dell Warranty Ticket Generator. Goodbye!${NC}"
                exit 0
            fi
            ;;
        *)
            echo -e "${RED}Invalid option. Using default issue description.${NC}"
            issue_description="The device is experiencing hardware issues requiring service."
            ;;
    esac
    
    echo "$issue_description"
}

# Function to generate troubleshooting steps based on specific issue
generate_troubleshooting() {
    local issue_option=$1
    local troubleshooting=""
    
    case "$issue_option" in
        1) # Keyboard
            troubleshooting="We've tried cleaning the keyboard and using an external keyboard which works fine, confirming the issue is with the built-in keyboard."
            ;;
        2) # Screen
            troubleshooting="We tested with an external monitor, and the computer works correctly, confirming this is a display hardware issue."
            ;;
        3) # Touchpad/Mouse
            troubleshooting="We've tried adjusting all touchpad settings and using an external mouse which works fine."
            ;;
        4) # Battery
            troubleshooting="We've tested with multiple chargers and confirmed the battery health is poor. Battery replacement is needed."
            ;;
        5) # OS
            troubleshooting="We've tried system restore, safe mode, and recovery options. The OS continues to have issues."
            ;;
        6) # Internal Component
            troubleshooting="We've run hardware diagnostics that indicate a component failure. The laptop needs depot service."
            ;;
        7) # Ports
            troubleshooting="We've tried multiple cables and devices. The ports appear to be physically damaged and need replacement."
            ;;
        8) # Audio
            troubleshooting="We've verified the issue persists across multiple applications and after driver updates."
            ;;
        9) # Other
            troubleshooting="We've performed basic troubleshooting and determined this requires depot service."
            ;;
        *)
            troubleshooting="Basic troubleshooting has been performed. The issue persists and requires depot service."
            ;;
    esac
    
    echo "$troubleshooting"
}

# Function to submit ticket to Dell Tech Direct
submit_dell_ticket() {
    local serial=$1
    local damage_category=$2
    local specific_issue=$3
    local troubleshooting=$4
    
    # Combine damage category and specific issue for the full description
    local full_description="${damage_category} ${specific_issue}"
    
    echo -e "${YELLOW}Preparing to submit Dell Tech Direct ticket for serial number: ${serial}${NC}"
    echo -e "${BLUE}Full Description:${NC} $full_description"
    echo -e "${BLUE}Troubleshooting Steps:${NC} $troubleshooting"
    
    # This is where you would make the actual API call to Dell Tech Direct
    # Replace with the actual Dell Tech Direct API endpoint and method
    echo -e "${YELLOW}Submitting ticket to Dell Tech Direct...${NC}"
    
    # Placeholder for the actual Dell Tech Direct API call
    # DELL_API_KEY="your_dell_api_key_here"
    # DELL_ENDPOINT="https://techdirect.dell.com/api/v1/warranty/create"
    
    # In a real implementation, you would make an API call like:
    # RESPONSE=$(curl -s -X POST "$DELL_ENDPOINT" \
    #                -H "Authorization: Bearer $DELL_API_KEY" \
    #                -H "Content-Type: application/json" \
    #                -d '{
    #                      "serial_number": "'"$serial"'",
    #                      "issue_description": "'"$full_description"'",
    #                      "troubleshooting": "'"$troubleshooting"'",
    #                      "contact": {
    #                          "name": "IT Support",
    #                          "email": "it@yourcompany.com",
    #                          "phone": "555-123-4567"
    #                      }
    #                    }')
    
    # For this demo, we'll just simulate a successful submission
    echo -e "${GREEN}Ticket successfully submitted to Dell Tech Direct!${NC}"
    echo "Ticket Reference: DELL$(date +%Y%m%d%H%M%S)"
}

# Main execution loop
main() {
    show_banner
    
    while true; do
        # Ask for asset tag
        echo -e "${CYAN}Enter asset tag (or 'quit' to exit):${NC}"
        read -r ASSET_TAG
        
        # Check if user wants to quit
        if [[ "$ASSET_TAG" == "quit" ]]; then
            echo -e "${GREEN}Exiting Dell Warranty Ticket Generator. Goodbye!${NC}"
            exit 0
        fi
        
        # Get serial number from Snipe-IT
        SERIAL_NUMBER=$(get_serial_number "$ASSET_TAG")
        
        # Check if serial number lookup was successful
        if [ $? -ne 0 ]; then
            echo -e "${RED}Failed to get serial number. Please try again.${NC}"
            echo ""
            continue
        fi
        
        # Get damage category
        DAMAGE_CATEGORY=$(get_damage_category)
        
        # Get specific issue
        SPECIFIC_ISSUE=$(get_specific_issue)
        
        # Generate troubleshooting steps based on the specific issue option
        TROUBLESHOOTING=$(generate_troubleshooting "$ISSUE_OPTION")
        
        # Display final ticket details for confirmation
        echo -e "${YELLOW}---------------------------------------${NC}"
        echo -e "${BLUE}Serial Number:${NC} $SERIAL_NUMBER"
        echo -e "${BLUE}Damage Category:${NC} $DAMAGE_CATEGORY"
        echo -e "${BLUE}Specific Issue:${NC} $SPECIFIC_ISSUE"
        echo -e "${BLUE}Troubleshooting:${NC} $TROUBLESHOOTING"
        echo -e "${YELLOW}---------------------------------------${NC}"
        
        # Ask for confirmation
        echo -e "${CYAN}Submit this ticket? (y/n or 'quit' to exit):${NC}"
        read -r CONFIRM
        
        # Check if user wants to quit
        if [[ "$CONFIRM" == "quit" ]]; then
            echo -e "${GREEN}Exiting Dell Warranty Ticket Generator. Goodbye!${NC}"
            exit 0
        fi
        
        # Check if user confirms submission
        if [[ "$CONFIRM" == "y" || "$CONFIRM" == "Y" || "$CONFIRM" == "yes" || "$CONFIRM" == "Yes" ]]; then
            # Submit ticket to Dell Tech Direct
            submit_dell_ticket "$SERIAL_NUMBER" "$DAMAGE_CATEGORY" "$SPECIFIC_ISSUE" "$TROUBLESHOOTING"
            
            echo -e "${GREEN}✓ Warranty ticket submitted successfully!${NC}"
        else
            echo -e "${YELLOW}Ticket submission cancelled.${NC}"
        fi
        
        echo -e "${YELLOW}---------------------------------------${NC}"
        echo ""
    done
}

# Start the program
main
