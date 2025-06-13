import time
import logging
from notion_client import Client
from datetime import datetime, timedelta
import urllib3
import requests
import sys
import json
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')


# Completely suppress all logging except our own
logging.getLogger().setLevel(logging.CRITICAL)
for log_name, logger in logging.Logger.manager.loggerDict.items():
    if isinstance(logger, logging.Logger):
        logger.setLevel(logging.CRITICAL)

# Now set up our own clean logging
logging.basicConfig(
    level=logging.INFO,
    format='%(message)s',
    force=True
)

# Suppress all HTTP-related warnings
urllib3.disable_warnings()

# Path to content.json file
CONTENT_JSON_PATH = os.path.join("JSON Files", "content.json")

def store_content_in_json(content_data):
    """Store content data in the JSON file"""
    try:
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(CONTENT_JSON_PATH), exist_ok=True)
        
        # Write content to JSON file
        with open(CONTENT_JSON_PATH, 'w', encoding='utf-8') as f:
            json.dump(content_data, f, indent=4, ensure_ascii=False)
        logging.info(f"Content stored in {CONTENT_JSON_PATH}")
        return True
    except Exception as e:
        logging.error(f"Error storing content in JSON: {e}")
        return False

# ANSI colors for terminal output
class Colors:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    BLUE = "\033[94m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    CYAN = "\033[96m"
    MAGENTA = "\033[95m"
    DIM = "\033[2m"

def log(message, level="info", newline=True):
    """Print a nicely formatted log message with timestamp."""
    timestamp = datetime.now().strftime("%I:%M:%S %p")
    
    if level == "info":
        prefix = f"{Colors.BLUE}ℹ{Colors.RESET}"
        color = Colors.RESET
    elif level == "success":
        prefix = f"{Colors.GREEN}✓{Colors.RESET}"
        color = Colors.GREEN
    elif level == "warn":
        prefix = f"{Colors.YELLOW}⚠{Colors.RESET}"
        color = Colors.YELLOW
    elif level == "error":
        prefix = f"{Colors.RED}✗{Colors.RESET}"
        color = Colors.RED
    elif level == "wait":
        prefix = f"{Colors.CYAN}◔{Colors.RESET}"
        color = Colors.CYAN
    elif level == "header":
        prefix = f"{Colors.MAGENTA}▶{Colors.RESET}"
        color = Colors.MAGENTA + Colors.BOLD
    else:
        prefix = " "
        color = Colors.RESET
    
    log_msg = f"{Colors.DIM}[{timestamp}]{Colors.RESET} {prefix} {color}{message}{Colors.RESET}"
    
    if newline:
        print(log_msg.encode('utf-8', errors='ignore').decode('utf-8'))

    else:
        print(log_msg, end="", flush=True)

# Notion configuration
NOTION_TOKEN = "ntn_cC7520095381SElmcgTOADYsGnrABFn2ph1PrcaGSst2dv"
NOTION_DATABASE_ID = "1a402cd2c14280909384df6c898ddcb3"

class TargetNotionHandler:
    def __init__(self, token, database_id):
        try:
            log("Validating Notion token...", "info")
            self.notion = Client(auth=token, log_level=logging.CRITICAL)
            # Test the connection by trying to access the database
            self.notion.databases.retrieve(database_id)
            log("Notion connection successful", "success")
        except Exception as e:
            log(f"Failed to initialize Notion client: {str(e)}", "error")
            raise Exception(f"Notion initialization failed: {str(e)}")
            
        self.database_id = database_id
        self.pending_voiceovers = []

    def get_block_content(self, block_id):
        """Recursively get content from a block and its children"""
        try:
            content = []
            blocks = self.notion.blocks.children.list(block_id).get('results', [])
            
            for block in blocks:
                # Handle paragraph blocks
                if block['type'] == 'paragraph':
                    text = ''.join([
                        rt.get('text', {}).get('content', '')
                        for rt in block['paragraph'].get('rich_text', [])
                    ])
                    if text.strip():
                        content.append(text)
                
                # Handle other block types that might contain text
                elif block['type'] in ['heading_1', 'heading_2', 'heading_3', 'bulleted_list_item', 'numbered_list_item']:
                    text = ''.join([
                        rt.get('text', {}).get('content', '')
                        for rt in block[block['type']].get('rich_text', [])
                    ])
                    if text.strip():
                        content.append(text)
                
                # Recursively get content from child blocks if they exist
                if block.get('has_children', False):
                    child_content = self.get_block_content(block['id'])
                    content.extend(child_content)
            
            return content
        except Exception as e:
            log(f"Error getting block content: {str(e)}", "error")
            return []

    def get_records_for_voiceover(self):
        """Get records that have content but haven't been processed for voiceover"""
        try:
            # Query only for records that match our criteria
            unvoiced_records = self.notion.databases.query(
                database_id=self.database_id,
                filter={
                    "and": [
                        {"property": "Voiceover", "checkbox": {"equals": False}},
                        {"property": "Script", "checkbox": {"equals": True}},
                        {"property": "New Title", "title": {"is_not_empty": True}}
                    ]
                }
            ).get('results', [])
            
            log(f"Found {len(unvoiced_records)} unvoiced records", "info")
            
            # Process records that have content
            processed_records = []
            content_data = {
                "records": []
            }
            
            for record in unvoiced_records:
                try:
                    # Debug logging
                    log("Processing new record...", "info")
                    
                    # Validate record structure first
                    if not record:
                        log("Record is None", "error")
                        continue
                        
                    if not isinstance(record, dict):
                        log(f"Record is not a dictionary: {type(record)}", "error")
                        continue
                        
                    if 'properties' not in record:
                        log("No properties in record", "error")
                        continue
                        
                    if 'id' not in record:
                        log("No id in record", "error")
                        continue

                    properties = record['properties']
                    if not isinstance(properties, dict):
                        log(f"Properties is not a dictionary: {type(properties)}", "error")
                        continue

                    # Get the title with better error handling
                    title = 'Untitled'
                    try:
                        if 'New Title' in properties:
                            title_prop = properties['New Title']
                            if isinstance(title_prop, dict) and 'title' in title_prop:
                                title_array = title_prop['title']
                                if isinstance(title_array, list) and len(title_array) > 0:
                                    first_title = title_array[0]
                                    if isinstance(first_title, dict) and 'text' in first_title:
                                        text_content = first_title['text']
                                        if isinstance(text_content, dict) and 'content' in text_content:
                                            title = text_content['content']
                    except Exception as e:
                        log(f"Error extracting title: {str(e)}", "error")
                        # Continue with default title
                    
                    # Get all content recursively
                    content = self.get_block_content(record['id'])
                    if not content:
                        log(f"No content found for record: {title}", "error")
                        continue
                        
                    script = ' '.join(content).strip()
                    if not script:
                        log(f"Empty script for record: {title}", "error")
                        continue
                    
                    # Add to processed records
                    processed_records.append(record)
                    
                    # Create the record data
                    record_data = {
                        "id": record['id'],
                        "title": title,
                        "content": script,
                        "channel": ""  # Initialize with empty string
                    }
                    
                    # Safely get channel
                    try:
                        if 'Channel' in properties:
                            channel_prop = properties['Channel']
                            if isinstance(channel_prop, dict) and 'select' in channel_prop:
                                select_prop = channel_prop['select']
                                if isinstance(select_prop, dict) and 'name' in select_prop:
                                    record_data["channel"] = select_prop['name']
                    except Exception as e:
                        log(f"Error extracting channel: {str(e)}", "error")
                        # Keep default empty string for channel
                    
                    # Verify record data before adding
                    if len(record_data["content"]) > 0:
                        content_data["records"].append(record_data)
                        log(f"Found content for: {title}", "success")
                        log(f"Content length: {len(script)} characters", "info")
                        log(f"Content preview: {script[:100]}...", "info")
                    else:
                        log(f"Warning: Empty content for {title}", "warn")
                            
                except Exception as e:
                    log(f"Error processing record {title if 'title' in locals() else 'Unknown'}: {str(e)}", "error")
                    continue

            # Store content in JSON file
            if content_data["records"]:
                try:
                    # Ensure directory exists
                    os.makedirs(os.path.dirname(CONTENT_JSON_PATH), exist_ok=True)
                    
                    # Convert to JSON string first
                    json_str = json.dumps(content_data, indent=4, ensure_ascii=False)
                    
                    # Verify JSON string has content
                    if '"content": ""' in json_str:
                        log("Warning: Empty content detected in JSON string", "warn")
                    
                    # Write to file
                    with open(CONTENT_JSON_PATH, 'w', encoding='utf-8') as f:
                        f.write(json_str)
                    
                    # Verify file contents
                    with open(CONTENT_JSON_PATH, 'r', encoding='utf-8') as f:
                        saved_data = json.load(f)
                        if saved_data["records"][0]["content"]:
                            log(f"Successfully saved content ({len(saved_data['records'][0]['content'])} characters)", "success")
                        else:
                            log("Error: Content is empty in saved file", "error")
                            
                except Exception as e:
                    log(f"Error saving content: {str(e)}", "error")
                    # Emergency backup - try writing to a different file
                    backup_path = os.path.join("JSON Files", "content_backup.json")
                    with open(backup_path, 'w', encoding='utf-8') as f:
                        json.dump(content_data, f, indent=4, ensure_ascii=False)
                    log(f"Created backup file at {backup_path}", "info")
            
            return processed_records
            
        except Exception as e:
            log(f"Error getting records for voiceover: {str(e)}", "error")
            return []

    def update_notion_checkboxes(self, page_id, voiceover=None, ready_to_be_edited=None):
        """Update the Voiceover and Ready to Be Edited checkboxes for a record"""
        try:
            properties = {}
            if voiceover is not None:
                properties["Voiceover"] = {"checkbox": voiceover}
                log(f"Setting Voiceover checkbox to: {voiceover}", "info")
            
            if ready_to_be_edited is not None:
                properties["Ready to Be Edited"] = {"checkbox": ready_to_be_edited}
                log(f"Setting Ready to Be Edited checkbox to: {ready_to_be_edited}", "info")

            if properties:
                log(f"Updating checkboxes for page {page_id}", "info")
                response = self.notion.pages.update(
                    page_id=page_id,
                    properties=properties
                )
                
                # Verify the update
                if "properties" in response:
                    success = True
                    if voiceover is not None and "Voiceover" in response["properties"]:
                        actual_value = response["properties"]["Voiceover"].get("checkbox", None)
                        if actual_value != voiceover:
                            log(f"Voiceover checkbox value mismatch. Expected: {voiceover}, Got: {actual_value}", "error")
                            success = False
                    
                    if ready_to_be_edited is not None and "Ready to Be Edited" in response["properties"]:
                        actual_value = response["properties"]["Ready to Be Edited"].get("checkbox", None)
                        if actual_value != ready_to_be_edited:
                            log(f"Ready to Be Edited checkbox value mismatch. Expected: {ready_to_be_edited}, Got: {actual_value}", "error")
                            success = False
                    
                    return success
                else:
                    log("No properties found in response", "error")
                    return False
            
            return True
            
        except Exception as e:
            log(f"Error updating checkboxes: {str(e)}", "error")
            return False

def monitor_notion_database():
    notion_handler = TargetNotionHandler(NOTION_TOKEN, NOTION_DATABASE_ID)
    
    # Clear screen and print header
    print("\033[H\033[J", end="")
    log("✨ VOICEOVER MONITOR", "header")
    
    while True:
        try:
            records = notion_handler.get_records_for_voiceover()
            
            # Show next check time
            next_check = datetime.now() + timedelta(seconds=15)
            log(f"Next check at {next_check.strftime('%I:%M:%S %p')}", "wait")
            time.sleep(15)
            
        except KeyboardInterrupt:
            log("Stopping monitor...", "info")
            break
            
        except Exception as e:
            log(f"Error: {str(e)}", "error")
            time.sleep(30)

if __name__ == "__main__":
    try:
        monitor_notion_database()
    except KeyboardInterrupt:
        print("\n")
        log("Monitor stopped", "info")
        log("Goodbye! ✌️", "header")