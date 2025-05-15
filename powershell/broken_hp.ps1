#!/bin/bash

# HP Repair Ticket Generator Script
# This script creates HP repair tickets in a loop until the user types "quit"

# Colors for better readability
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
PURPLE='\033[0;35m'
NC='\033[0m' # No Color

# Function to display script banner
show_banner() {
    echo -e "${BLUE}╔════════════════════════════════════════════════════════╗${NC}"
    echo -e "${BLUE}║               ${PURPLE}HP Repair Ticket Generator${BLUE}               ║${NC}"
    echo -e "${BLUE}╚════════════════════════════════════════════════════════╝${NC}"
    echo -e "${YELLOW}Type 'quit' at any prompt to exit the program${NC}"
    echo ""
}

# Function to display damage categories
show_damage_categories() {
    echo -e "${CYAN}Select damage category:${NC}"
    echo -e "  ${GREEN}1${NC} - LCD"
    echo -e "  ${GREEN}2${NC} - Liquid Spill"
    echo -e "  ${GREEN}3${NC} - Power Surge"
    echo -e "  ${GREEN}4${NC} - Hinges"
    echo -e "  ${GREEN}5${NC} - Casing"
    echo -e "  ${GREEN}6${NC} - Other"
    echo ""
}

# Function to show specific issue types
show_specific_issues() {
    echo -e "${CYAN}Select specific issue:${NC}"
    echo -e "  ${GREEN}1${NC} - Keyboard (keys not working)"
    echo -e "  ${GREEN}2${NC} - Screen (cracked, lines, blank)"
    echo -e "  ${GREEN}3${NC} - Touchpad/Mouse"
    echo -e "  ${GREEN}4${NC} - Battery"
    echo -e "  ${GREEN}5${NC} - Operating System"
    echo -e "  ${GREEN}6${NC} - Internal Component"
    echo -e "  ${GREEN}7${NC} - Ports (USB, HDMI, etc.)"
    echo -e "  ${GREEN}8${NC} - Audio"
    echo -e "  ${GREEN}9${NC} - Other (custom issue)"
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

# Function to get damage category description with story
get_damage_category() {
    show_damage_categories
    
    echo -e "${CYAN}Enter damage category number or 'quit' to exit:${NC}"
    read -r DAMAGE_OPTION
    
    # Check if user wants to quit
    if [[ "$DAMAGE_OPTION" == "quit" ]]; then
        echo -e "${GREEN}Exiting HP Repair Ticket Generator. Goodbye!${NC}"
        exit 0
    fi
    
    local damage_description=""
    
    case "$DAMAGE_OPTION" in
        1) # LCD
            damage_description="The laptop was accidentally dropped from a desk onto a hard floor. The impact caused visible damage to the LCD screen. The screen now shows distortion and there are visible cracks in the display panel."
            ;;
        2) # Liquid Spill
            damage_description="A water bottle was accidentally knocked over near the laptop. Liquid spilled onto the keyboard and possibly entered internal components. After the spill, the device started exhibiting problems with certain functions."
            ;;
        3) # Power Surge
            damage_description="During a thunderstorm, there was a power surge that affected several devices in the office. The laptop was plugged in at the time, and after the surge, it started having problems. The surge appears to have damaged internal components."
            ;;
        4) # Hinges
            damage_description="The laptop was opened too forcefully, causing the screen hinges to snap. Now the screen is loose and wobbles when positioned. The hinge assembly needs to be replaced to restore proper functionality."
            ;;
        5) # Casing
            damage_description="The laptop was accidentally bumped off a table and landed on its corner. This impact caused the casing to crack near one of the corners. The damage to the casing could potentially expose internal components to dust or further damage if not repaired."
            ;;
        6) # Other
            echo -e "${CYAN}Enter custom damage story or 'quit' to exit:${NC}"
            read -r damage_description
            
            # Check if user wants to quit
            if [[ "$damage_description" == "quit" ]]; then
                echo -e "${GREEN}Exiting HP Repair Ticket Generator. Goodbye!${NC}"
                exit 0
            fi
            ;;
        *)
            echo -e "${RED}Invalid option. Using default damage category.${NC}"
            damage_description="The device has sustained damage that is affecting its functionality. The exact cause of the damage is unclear, but repair service is needed."
            ;;
    esac
    
    echo "$damage_description"
}

# Function to get specific issue description
get_specific_issue() {
    show_specific_issues
    
    echo -e "${CYAN}Enter specific issue number or 'quit' to exit:${NC}"
    read -r ISSUE_OPTION
    
    # Check if user wants to quit
    if [[ "$ISSUE_OPTION" == "quit" ]]; then
        echo -e "${GREEN}Exiting HP Repair Ticket Generator. Goodbye!${NC}"
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
                echo -e "${GREEN}Exiting HP Repair Ticket Generator. Goodbye!${NC}"
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
                echo -e "${GREEN}Exiting HP Repair Ticket Generator. Goodbye!${NC}"
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
            troubleshooting="We've tried system restore, safe mode, and HP recovery options. The OS continues to have issues."
            ;;
        6) # Internal Component
            troubleshooting="We've run HP hardware diagnostics that indicate a component failure. The laptop needs service."
            ;;
        7) # Ports
            troubleshooting="We've tried multiple cables and devices. The ports appear to be physically damaged and need replacement."
            ;;
        8) # Audio
            troubleshooting="We've verified the issue persists across multiple applications and after driver updates."
            ;;
        9) # Other
            troubleshooting="We've performed basic troubleshooting and determined this requires service."
            ;;
        *)
            troubleshooting="Basic troubleshooting has been performed. The issue persists and requires service."
            ;;
    esac
    
    echo "$troubleshooting"
}

# Function to get HP case priority
get_case_priority() {
    echo -e "${CYAN}Select case priority:${NC}"
    echo -e "  ${GREEN}1${NC} - Critical (Complete system failure, business operations severely impacted)"
    echo -e "  ${GREEN}2${NC} - High (System partially functional, significant impact on business)"
    echo -e "  ${GREEN}3${NC} - Medium (System functional with limitations, moderate impact)"
    echo -e "  ${GREEN}4${NC} - Low (Minor issues, minimal impact on operations)"
    echo ""
    
    echo -e "${CYAN}Enter priority number or 'quit' to exit:${NC}"
    read -r PRIORITY_OPTION
    
    # Check if user wants to quit
    if [[ "$PRIORITY_OPTION" == "quit" ]]; then
        echo -e "${GREEN}Exiting HP Repair Ticket Generator. Goodbye!${NC}"
        exit 0
    fi
    
    local priority_level=""
    
    case "$PRIORITY_OPTION" in
        1)
            priority_level="Critical"
            ;;
        2)
            priority_level="High"
            ;;
        3)
            priority_level="Medium"
            ;;
        4)
            priority_level="Low"
            ;;
        *)
            echo -e "${RED}Invalid option. Using default priority.${NC}"
            priority_level="Medium"
            ;;
    esac
    
    echo "$priority_level"
}

# Function to submit ticket to HP Support Portal
submit_hp_ticket() {
    local serial=$1
    local damage_story=$2
    local specific_issue=$3
    local troubleshooting=$4
    local priority=$5
    
    # Format the full description with the story followed by the technical issue
    local full_description="HOW IT BROKE: $damage_story\n\nSPECIFIC ISSUE: $specific_issue"
    
    echo -e "${YELLOW}Preparing to submit HP Support ticket for serial number: ${serial}${NC}"
    echo -e "${BLUE}Priority:${NC} $priority"
    echo -e "${BLUE}How It Broke:${NC} $damage_story"
    echo -e "${BLUE}Specific Issue:${NC} $specific_issue"
    echo -e "${BLUE}Troubleshooting Steps:${NC} $troubleshooting"
    
    # This is where you would make the actual API call to HP Support Portal
    # Replace with the actual HP API endpoint and method
    echo -e "${YELLOW}Submitting ticket to HP Support Portal...${NC}"
    
    # Placeholder for the actual HP API call
    # HP_API_KEY="your_hp_api_key_here"
    # HP_ENDPOINT="https://support.hp.com/api/v1/cases/create"
    
    # In a real implementation, you would make an API call like:
    # RESPONSE=$(curl -s -X POST "$HP_ENDPOINT" \
    #                -H "Authorization: Bearer $HP_API_KEY" \
    #                -H "Content-Type: application/json" \
    #                -d '{
    #                      "serial_number": "'"$serial"'",
    #                      "issue_description": "'"$full_description"'",
    #                      "troubleshooting": "'"$troubleshooting"'",
    #                      "priority": "'"$priority"'",
    #                      "contact": {
    #                          "name": "IT Support",
    #                          "email": "it@yourcompany.com",
    #                          "phone": "555-123-4567"
    #                      }
    #                    }')
    
    # For this demo, we'll just simulate a successful submission
    echo -e "${GREEN}Ticket successfully submitted to HP Support Portal!${NC}"
    echo "Case Number: HP$(date +%Y%m%d%H%M%S)"
    
    # HP's estimated response times based on priority
    case "$priority" in
        "Critical")
            echo -e "${PURPLE}Estimated response time: 2 hours${NC}"
            ;;
        "High")
            echo -e "${PURPLE}Estimated response time: 4 hours${NC}"
            ;;
        "Medium")
            echo -e "${PURPLE}Estimated response time: 8 hours${NC}"
            ;;
        "Low")
            echo -e "${PURPLE}Estimated response time: 24 hours${NC}"
            ;;
    esac
}

# Function to check warranty status
check_warranty_status() {
    local serial=$1
    
    echo -e "${YELLOW}Checking warranty status for serial number: ${serial}...${NC}"
    
    # This is where you would make the actual API call to HP Warranty Service
    # Replace with the actual HP API endpoint and method
    
    # Placeholder for the actual HP Warranty API call
    # HP_WARRANTY_ENDPOINT="https://support.hp.com/api/v1/warranty/$serial"
    
    # For this demo, we'll just simulate a warranty check
    # Randomly choose whether the device is in warranty or not for demo purposes
    local random_number=$((RANDOM % 10))
    
    if [ $random_number -ge 3 ]; then
        echo -e "${GREEN}✓ Device is under warranty until $(date -d "+$((RANDOM % 24 + 1)) months" +"%B %d, %Y")${NC}"
        echo -e "${GREEN}✓ Hardware support included${NC}"
        return 0
    else
        echo -e "${RED}✗ Device is out of warranty${NC}"
        echo -e "${YELLOW}Would you like to proceed with a billable repair? (y/n)${NC}"
        read -r PROCEED
        
        if [[ "$PROCEED" == "y" || "$PROCEED" == "Y" || "$PROCEED" == "yes" || "$PROCEED" == "Yes" ]]; then
            echo -e "${YELLOW}Proceeding with billable repair request.${NC}"
            return 0
        else
            echo -e "${RED}Repair request cancelled due to warranty status.${NC}"
            return 1
        fi
    fi
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
            echo -e "${GREEN}Exiting HP Repair Ticket Generator. Goodbye!${NC}"
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
        
        # Check warranty status
        check_warranty_status "$SERIAL_NUMBER"
        
        # Check if warranty check was successful or user wants to proceed
        if [ $? -ne 0 ]; then
            echo ""
            continue
        fi
        
        # Get damage category with story
        DAMAGE_STORY=$(get_damage_category)
        
        # Get specific issue
        SPECIFIC_ISSUE=$(get_specific_issue)
        
        # Generate troubleshooting steps based on the specific issue option
        TROUBLESHOOTING=$(generate_troubleshooting "$ISSUE_OPTION")
        
        # Get case priority
        PRIORITY=$(get_case_priority)
        
        # Display final ticket details for confirmation
        echo -e "${YELLOW}---------------------------------------${NC}"
        echo -e "${BLUE}Serial Number:${NC} $SERIAL_NUMBER"
        echo -e "${BLUE}Priority:${NC} $PRIORITY"
        echo -e "${BLUE}How It Broke:${NC} $DAMAGE_STORY"
        echo -e "${BLUE}Specific Issue:${NC} $SPECIFIC_ISSUE"
        echo -e "${BLUE}Troubleshooting:${NC} $TROUBLESHOOTING"
        echo -e "${YELLOW}---------------------------------------${NC}"
        
        # Ask for confirmation
        echo -e "${CYAN}Submit this ticket? (y/n or 'quit' to exit):${NC}"
        read -r CONFIRM
        
        # Check if user wants to quit
        if [[ "$CONFIRM" == "quit" ]]; then
            echo -e "${GREEN}Exiting HP Repair Ticket Generator. Goodbye!${NC}"
            exit 0
        fi
        
        # Check if user confirms submission
        if [[ "$CONFIRM" == "y" || "$CONFIRM" == "Y" || "$CONFIRM" == "yes" || "$CONFIRM" == "Yes" ]]; then
            # Submit ticket to HP Support Portal
            submit_hp_ticket "$SERIAL_NUMBER" "$DAMAGE_STORY" "$SPECIFIC_ISSUE" "$TROUBLESHOOTING" "$PRIORITY"
            
            echo -e "${GREEN}✓ HP repair ticket submitted successfully!${NC}"
        else
            echo -e "${YELLOW}Ticket submission cancelled.${NC}"
        fi
        
        echo -e "${YELLOW}---------------------------------------${NC}"
        echo ""
    done
}

# Start the program
main
