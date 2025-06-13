import os
import time
import logging
import traceback
import socket
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from pyairtable import Api
import sys
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
from notion_client import Client
import re
import warnings
from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
from queue import Queue, Empty

# Fix the file_cache warning by adding this import and setting
warnings.filterwarnings('ignore', message='file_cache is only supported with oauth2client<4.0.0')

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configuration
DESKTOP_DIRECTORY = os.path.expanduser("~/Desktop")
MAX_WORDS_PER_BLOCK = 250

# Voice URL mapping for different channels
VOICE_URLS = {
    "Red White & Real": "https://app.play.ht/studio/file/daOlVjalacOXwvuBCcVK?voice=s3://voice-cloning-zero-shot/5191967a-e431-4bae-a82e-ed030495963e/original/manifest.json",
    "Rachel Zegler": "https://app.play.ht/studio/file/3AzblQYbrq5z6R0FoL9e?voice=s3://voice-cloning-zero-shot/541946ca-d3c9-49c7-975b-09a4e42a991f/original/manifest.json",
    "Meghan Markle": "https://app.play.ht/studio/file/vvdIaXkaZ9PcRPITZisC?voice=s3://voice-cloning-zero-shot/541946ca-d3c9-49c7-975b-09a4e42a991f/original/manifest.json",
    "Knuckle Talk": "https://play.ht/studio/files/f57e51c2-bf57-4c8f-8094-a22991c8b45f?voice=s3%3A%2F%2Fvoice-cloning-zero-shot%2F54a23eaa-05de-4c08-a9dc-090513a7dc3b%2Foriginal%2Fmanifest.json",
    "Royal Family": "https://play.ht/studio/files/d121f4fc-a6ab-490b-8a3f-9c76bf6688a8?voice=s3%3A%2F%2Fvoice-cloning-zero-shot%2F904e22a9-69cf-4e1c-b6c0-eb8cbfb394ec%2Foriginal%2Fmanifest.json"
}

# Folder paths for audio files
CELEBRITY_VO_PATH = "E:\\Celebrity Voice Overs"
VOICEOVERS_MOVED_PATH = "E:\\Voiceovers Moved"

# Create a queue for new file events
new_file_queue = Queue()

# Default voice URL as fallback
DEFAULT_VOICE_URL = "https://app.play.ht/studio/file/vvdIaXkaZ9PcRPITZisC?voice=s3://voice-cloning-zero-shot/541946ca-d3c9-49c7-975b-09a4e42a991f/original/manifest.json"

# Add these constants
SCOPES = ['https://www.googleapis.com/auth/documents.readonly']
TOKEN_PATH = 'token.pickle'
CREDENTIALS_PATH = 'credentials.json'  # You'll need to get this from Google Cloud Console

# Notion configuration
NOTION_TOKEN = "ntn_cC7520095381SElmcgTOADYsGnrABFn2ph1PrcaGSst2dv"
NOTION_DATABASE_ID = "1a402cd2c14280909384df6c898ddcb3"  # Updated to correct database ID

# Add these constants at the top with your other constants
PLAYHT_COOKIES_FILE = "Dakota.pkl"
PLAYHT_LOGIN_URL = "https://app.play.ht/login"

# Add these constants with your other constants
SERVICE_ACCOUNT_FILE = 'helpful-data-459308-b9-56243ce0290d.json'  # Using the correct service account credentials
GOOGLE_DRIVE_FOLDER_ID = "WrongID"  # Folder ID for voiceover uploads

class NotionHandler:
    def __init__(self, token, database_id):
        self.notion = Client(auth=token)
        self.database_id = database_id

    def get_done_items(self):
        """Get items from the DONE column."""
        try:
            response = self.notion.databases.query(
                database_id=self.database_id,
                filter={
                    "property": "UPDATE",
                    "status": {  # Using status type as in the old code
                        "equals": "DONE"
                    }
                }
            )
            return response['results']
        except Exception as e:
            logging.error(f"Error querying Notion database: {str(e)}")
            return []

    def get_google_docs_link(self, page_id):
        """Extract Google Docs link from page comments."""
        try:
            comments = self.notion.comments.list(block_id=page_id)
            for comment in comments['results']:
                text = comment['rich_text'][0]['text']['content']
                if 'docs.google.com' in text:
                    urls = re.findall(r'https://docs.google.com/\S+', text)
                    if urls:
                        return urls[0]
            return None
        except Exception as e:
            logging.error(f"Error getting comments for page {page_id}: {str(e)}")
            return None

class TargetNotionHandler:
    def __init__(self, token, database_id):
        self.notion = Client(auth=token)
        self.database_id = database_id
        # Log database structure for debugging
        self.log_database_schema()

    def log_database_schema(self):
        """Get and log database schema to verify property names"""
        try:
            database = self.notion.databases.retrieve(self.database_id)
            logging.info("Retrieved database schema for debugging")
            if "properties" in database:
                properties = database["properties"]
                logging.info("Database properties:")
                for name, prop in properties.items():
                    prop_type = prop.get("type", "unknown")
                    logging.info(f"  - '{name}' (type: {prop_type})")
                
                # Log checkbox properties with extra detail
                logging.info("CHECKBOX PROPERTIES (detailed):")
                for name, prop in properties.items():
                    if prop.get("type") == "checkbox":
                        logging.info(f"  - Checkbox: '{name}' (ID: {prop.get('id', 'unknown')})")
                        # Dump the full property definition for detailed inspection
                        import json
                        logging.info(f"    Full definition: {json.dumps(prop)}")
                
                # Try to find the "Ready to Be Edited" property with case-insensitive matching
                ready_prop = None
                for name, prop in properties.items():
                    if "ready" in name.lower() and "edit" in name.lower():
                        logging.info(f"Found potential match for 'Ready to Be Edited': '{name}'")
                        ready_prop = name
                
                if ready_prop:
                    logging.info(f"Will use '{ready_prop}' as the property name for 'Ready to Be Edited'")
                    # Store this for later use
                    self.ready_to_be_edited_prop_name = ready_prop
                else:
                    logging.warning("Could not find a property matching 'Ready to Be Edited'")
                    self.ready_to_be_edited_prop_name = "Ready to Be Edited"  # Default fallback
            else:
                logging.warning("No properties found in database schema")
        except Exception as e:
            logging.error(f"Error retrieving database schema: {str(e)}")
            self.ready_to_be_edited_prop_name = "Ready to Be Edited"  # Default fallback

    def split_into_sentences(self, text):
        """Split text into sentences, preserving sentence boundaries"""
        # First split by obvious sentence endings
        sentences = []
        current = []
        
        # Split by words to preserve spacing
        words = text.split()
        for word in words:
            current.append(word)
            # Check for sentence endings
            if word.endswith('.') or word.endswith('!') or word.endswith('?'):
                sentences.append(' '.join(current))
                current = []
        
        # Add any remaining words as the last sentence
        if current:
            sentences.append(' '.join(current))
        
        return sentences

    def create_content_blocks(self, text):
        """Create content blocks that respect sentence boundaries and stay under 2000 chars"""
        sentences = self.split_into_sentences(text)
        blocks = []
        current_block = []
        current_length = 0
        
        for sentence in sentences:
            # +1 for the space we'll add between sentences
            sentence_length = len(sentence) + 1
            
            # If adding this sentence would exceed 2000 chars, start a new block
            if current_length + sentence_length > 2000 and current_block:
                blocks.append(' '.join(current_block))
                current_block = []
                current_length = 0
            
            current_block.append(sentence)
            current_length += sentence_length
        
        # Add the last block if it exists
        if current_block:
            blocks.append(' '.join(current_block))
        
        return blocks

    def create_record(self, docs_url, new_script="", new_title="", voiceover=False):
        """Create a new record in the target Notion database"""
        try:
            # First create the page with basic properties
            properties = {
                "Docs": {"url": docs_url},
                "New Title": {"title": [{"text": {"content": new_title}}]} if new_title else {"title": []},
                "Voiceover": {"checkbox": voiceover}
            }

            # Create the page first
            response = self.notion.pages.create(
                parent={"database_id": self.database_id},
                properties=properties
            )
            
            # If we have a script, update the page content
            if new_script:
                # Split content into blocks that respect sentence boundaries
                content_blocks = self.create_content_blocks(new_script)
                
                # Create a paragraph block for each content block
                children = [
                    {
                        "object": "block",
                        "type": "paragraph",
                        "paragraph": {
                            "rich_text": [{"type": "text", "text": {"content": block.strip()}}]
                        }
                    }
                    for block in content_blocks
                ]
                
                self.notion.blocks.children.append(
                    response['id'],
                    children=children
                )
            
            logging.info(f"Created new Notion record with Docs URL: {docs_url}")
            return True
        except Exception as e:
            logging.error(f"Error creating Notion record: {e}")
            return False

    def update_record(self, page_id, new_script="", new_title="", voiceover=None, ready_to_be_edited=None):
        """Update an existing record in the target Notion database"""
        try:
            # Update properties except New Script
            properties = {}
            if new_title:
                properties["New Title"] = {"title": [{"text": {"content": new_title}}]}
                logging.info(f"Adding 'New Title' to update: {new_title}")
            if voiceover is not None:
                properties["Voiceover"] = {"checkbox": voiceover}
                logging.info(f"Adding 'Voiceover' checkbox to update: {voiceover}")
            if ready_to_be_edited is not None:
                ready_prop_name = getattr(self, 'ready_to_be_edited_prop_name', "Ready to Be Edited")
                properties[ready_prop_name] = {"checkbox": ready_to_be_edited}
                logging.info(f"Adding '{ready_prop_name}' checkbox to update: {ready_to_be_edited}")

            if properties:
                logging.info(f"Updating page {page_id} with properties: {properties}")
                response = self.notion.pages.update(
                    page_id=page_id,
                    properties=properties
                )
                logging.info(f"Notion API response status: success")
                # Log the actual property values from the response to verify
                if "properties" in response:
                    if voiceover is not None and "Voiceover" in response["properties"]:
                        checkbox_value = response["properties"]["Voiceover"].get("checkbox", None)
                        logging.info(f"Confirmed 'Voiceover' is now set to: {checkbox_value}")
                    if ready_to_be_edited is not None:
                        ready_prop_name = getattr(self, 'ready_to_be_edited_prop_name', "Ready to Be Edited")
                        if ready_prop_name in response["properties"]:
                            checkbox_value = response["properties"][ready_prop_name].get("checkbox", None)
                            logging.info(f"Confirmed '{ready_prop_name}' is now set to: {checkbox_value}")
            
            # If we have a script, update the page content
            if new_script:
                # First, check existing content
                existing_blocks = self.notion.blocks.children.list(page_id)
                existing_content = ""
                
                # Extract the existing content
                for block in existing_blocks.get('results', []):
                    if block['type'] == 'paragraph':
                        for text in block['paragraph'].get('rich_text', []):
                            existing_content += text.get('text', {}).get('content', '')
                
                # Only update content if it's different
                if existing_content.strip() != new_script.strip():
                    logging.info(f"Content differs - updating page {page_id}")
                    
                    # Delete existing blocks
                    for block in existing_blocks.get('results', []):
                        self.notion.blocks.delete(block['id'])
                    
                    # Split content into blocks that respect sentence boundaries
                    content_blocks = self.create_content_blocks(new_script)
                    
                    # Create a paragraph block for each content block
                    children = [
                        {
                            "object": "block",
                            "type": "paragraph",
                            "paragraph": {
                                "rich_text": [{"type": "text", "text": {"content": block.strip()}}]
                            }
                        }
                        for block in content_blocks
                    ]
                    
                    # Add new content as multiple blocks
                    self.notion.blocks.children.append(
                        page_id,
                        children=children
                    )
                else:
                    logging.info(f"Content already matches - skipping update for page {page_id}")

            logging.info(f"Updated Notion record: {page_id}")
            return True
        except Exception as e:
            logging.error(f"Error updating Notion record: {str(e)}")
            # Print more details about the error
            logging.error(f"Traceback: {traceback.format_exc()}")
            return False

    def get_existing_docs_urls(self):
        """Get all existing doc URLs from the database"""
        try:
            results = self.notion.databases.query(database_id=self.database_id)
            existing_urls = set()
            for page in results.get('results', []):
                url = page['properties'].get('Docs', {}).get('url')
                if url:
                    existing_urls.add(url)
            return existing_urls
        except Exception as e:
            logging.error(f"Error getting existing URLs: {e}")
            return set()

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
            
            # Process records that have content
            processed_records = []
            for record in unvoiced_records:
                blocks = self.notion.blocks.children.list(record['id'])
                if blocks.get('results', []):  # If there are any blocks (content)
                    processed_records.append(record)
                    title_prop = record['properties'].get('New Title', {}).get('title', [])
                    if title_prop:
                        title = title_prop[0].get('text', {}).get('content', 'Untitled')
                        logging.info(f"Record {title} is ready for voiceover processing")

            logging.info(f"Found {len(processed_records)} records ready for voiceover processing")
            return processed_records
            
        except Exception as e:
            logging.error(f"Error getting records for voiceover: {e}")
            return []

    def update_notion_checkboxes(self, page_id, voiceover=None, ready_to_be_edited=None):
        """Specific method just for updating checkboxes to ensure it works correctly"""
        try:
            properties = {}
            if voiceover is not None:
                properties["Voiceover"] = {"checkbox": voiceover}
                logging.info(f"Setting Voiceover checkbox to: {voiceover}")
            
            if ready_to_be_edited is not None:
                ready_prop_name = getattr(self, 'ready_to_be_edited_prop_name', "Ready to Be Edited")
                properties[ready_prop_name] = {"checkbox": ready_to_be_edited}
                logging.info(f"Adding '{ready_prop_name}' checkbox to update: {ready_to_be_edited}")

            if properties:
                logging.info(f"CHECKBOX UPDATE: Updating page {page_id} with checkboxes: {properties}")
                response = self.notion.pages.update(
                    page_id=page_id,
                    properties=properties
                )
                
                # Verify checkbox was updated
                if "properties" in response:
                    if voiceover is not None and "Voiceover" in response["properties"]:
                        value = response["properties"]["Voiceover"].get("checkbox", None)
                        logging.info(f"CHECKBOX UPDATE: Voiceover checkbox is now: {value}")
                        if value == voiceover:
                            logging.info("Voiceover checkbox was successfully updated")
                            return True
                        else:
                            logging.error(f"Voiceover checkbox value mismatch. Expected: {voiceover}, Got: {value}")
                            return False
                    else:
                        logging.error("Voiceover property not found in response")
                        return False
                else:
                    logging.error("No properties found in response")
                    return False
            
        except Exception as e:
            logging.error(f"Error updating Voiceover checkbox: {str(e)}")
            logging.error(f"Traceback: {traceback.format_exc()}")
            return False

    def check_page_properties(self, page_id):
        """Retrieve and log a page's current properties for debugging"""
        try:
            page = self.notion.pages.retrieve(page_id)
            logging.info(f"Retrieved page {page_id} to check properties")
            if "properties" in page:
                properties = page["properties"]
                logging.info("Current page properties:")
                # Check for checkboxes specifically
                for name, prop in properties.items():
                    if prop.get("type") == "checkbox":
                        value = prop.get("checkbox", None)
                        logging.info(f"  - Checkbox '{name}': {value}")
            else:
                logging.warning("No properties found in page data")
                
            return True
        except Exception as e:
            logging.error(f"Error retrieving page properties: {str(e)}")
            return False

    def update_notion_with_drive_link(self, page_id, drive_url):
        """Update the Voice Drive Link column in Notion."""
        try:
            properties = {
                "Voice Drive Link": {"url": drive_url}
            }
            
            response = self.notion.pages.update(
                page_id=page_id,
                properties=properties
            )
            
            logging.info(f"Updated Notion with Drive URL: {drive_url}")
            return True
        except Exception as e:
            logging.error(f"Error updating Notion with Drive URL: {e}")
            return False

    def get_unprocessed_records(self):
        """Get records that have Google Docs but haven't been processed"""
        try:
            return self.notion.databases.query(
                database_id=self.database_id,
                filter={
                    "and": [
                        {"property": "Voiceover", "checkbox": {"equals": False}},
                        {"property": "Docs", "url": {"is_not_empty": True}}  # Changed from Docs to Docs
                    ]
                }
            ).get('results', [])
        except Exception as e:
            logging.error(f"Error getting unprocessed records: {e}")
            return []

def remove_whitespace(text):
    return ' '.join(text.split())

def split_text(text, max_words=150):
    # Split text into paragraphs first (preserve original paragraph breaks)
    paragraphs = text.split('\n')
    chunks = []
    
    for paragraph in paragraphs:
        # Skip empty paragraphs
        if not paragraph.strip():
            continue
            
        # Clean the paragraph
        cleaned_para = remove_whitespace(paragraph)
        
        # Split into sentences (looking for ., !, ?)
        sentences = []
        current_sentence = []
        words = cleaned_para.split()
        
        for word in words:
            current_sentence.append(word)
            if word.endswith('.') or word.endswith('!') or word.endswith('?'):
                sentences.append(' '.join(current_sentence))
                current_sentence = []
        
        # Add any remaining words as a sentence
        if current_sentence:
            sentences.append(' '.join(current_sentence))
        
        # Group sentences into chunks
        current_chunk = []
        word_count = 0
        
        for sentence in sentences:
            sentence_words = len(sentence.split())
            
            # If adding this sentence would exceed limit, save current chunk and start new one
            if word_count + sentence_words > max_words and current_chunk:
                chunks.append(' '.join(current_chunk))
                current_chunk = [sentence]
                word_count = sentence_words
            else:
                current_chunk.append(sentence)
                word_count += sentence_words
        
        # Add the last chunk if it exists
        if current_chunk:
            chunks.append(' '.join(current_chunk))
    
    # Log chunk information
    for i, chunk in enumerate(chunks):
        word_count = len(chunk.split())
        logging.info(f"Chunk {i+1}: {word_count} words, {len(chunk)} characters")
        if word_count > 150:
            logging.warning(f"Chunk {i+1} has {word_count} words - check for very long sentences")

    return chunks

def find_available_port(start_port=9222, max_port=9299):
    """Find an available port for Chrome DevTools Protocol."""
    for port in range(start_port, max_port + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(('127.0.0.1', port))
                return port
            except socket.error:
                continue
    return None

def cleanup_chrome_processes():
    """Kill any existing Chrome processes that might interfere with automation"""
    try:
        if sys.platform == "win32":  # Windows
            os.system("taskkill /f /im chrome.exe")
            os.system("taskkill /f /im chromedriver.exe")
        elif sys.platform == "darwin":  # macOS
            os.system("pkill -f 'Google Chrome'")
            os.system("pkill -f 'chromedriver'")
        time.sleep(2)  # Give processes time to close
        logging.info("Cleaned up existing Chrome processes")
    except Exception as e:
        logging.warning(f"Error during Chrome cleanup: {e}")

def setup_chrome_driver():
    try:
        # First cleanup any existing Chrome processes
        cleanup_chrome_processes()
        
        # Find an available port for DevTools
        debug_port = find_available_port()
        if not debug_port:
            raise Exception("Could not find an available port for Chrome DevTools")
        logging.info(f"Using port {debug_port} for Chrome DevTools")
        
        chrome_options = Options()
        
        # Remove automation control bar and other automation indicators
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        
        # Maximize window properly
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--window-size=1920,1080")
        
        # Add options to help with Google OAuth
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument("--disable-notifications")
        
        # Important for OAuth authentication
        chrome_options.add_argument("--enable-features=NetworkService,NetworkServiceInProcess")
        chrome_options.add_argument("--disable-features=IsolateOrigins,site-per-process")
        
        # Set download preferences
        chrome_options.add_experimental_option('prefs', {
            'download.default_directory': CELEBRITY_VO_PATH,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True,
            'profile.default_content_setting_values.notifications': 2,
            'credentials_enable_service': True,
            'profile.password_manager_enabled': True
        })
        
        # Clean up and recreate the Chrome profile directory
        temp_profile_dir = os.path.join(os.path.expanduser("~/Desktop"), "dakota_chrome_profile")
        if os.path.exists(temp_profile_dir):
            try:
                shutil.rmtree(temp_profile_dir)
                logging.info("Removed existing Chrome profile directory")
            except Exception as e:
                logging.warning(f"Could not remove existing profile directory: {e}")
        
        # Create fresh profile directory
        try:
            os.makedirs(temp_profile_dir, exist_ok=True)
            logging.info(f"Created fresh Chrome profile at: {temp_profile_dir}")
        except Exception as e:
            logging.error(f"Could not create profile directory: {e}")
            temp_profile_dir = None
        
        if temp_profile_dir:
            chrome_options.add_argument(f"--user-data-dir={temp_profile_dir}")
        
        # Set more realistic user agent
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")
        
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        
    except Exception as e:
        logging.error(f"Failed to start Chrome: {str(e)}")
        logging.error(f"Traceback: {traceback.format_exc()}")
        try:
            cleanup_chrome_processes()
        except:
            pass
        raise

def setup_chrome_with_manager(chrome_options):
    """Setup Chrome using webdriver-manager"""
    try:
        service = Service(
            ChromeDriverManager().install(),
            log_path=os.devnull  # Suppress chromedriver logs
        )
        logging.info("ChromeDriverManager setup successful")
        return webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        logging.error(f"ChromeDriverManager setup failed: {e}")
        raise

def setup_chrome_system(chrome_options):
    """Setup Chrome using system installation"""
    try:
        # Try to find Chrome in common installation paths
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expanduser("~\\AppData\\Local\\Google\\Chrome\\Application\\chrome.exe")
        ]
        
        chrome_path = None
        for path in chrome_paths:
            if os.path.exists(path):
                chrome_path = path
                break
        
        if chrome_path:
            chrome_options.binary_location = chrome_path
            logging.info(f"Using Chrome at: {chrome_path}")
        
        service = Service()
        return webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        logging.error(f"System Chrome setup failed: {e}")
        raise

def setup_chrome_minimal(chrome_options):
    """Setup Chrome with minimal options as last resort"""
    try:
        # Create minimal options
        minimal_options = Options()
        minimal_options.add_argument("--no-sandbox")
        minimal_options.add_argument("--disable-dev-shm-usage")
        minimal_options.add_argument("--disable-gpu")
        
        # Copy download preferences
        minimal_options.add_experimental_option('prefs', {
            'download.default_directory': CELEBRITY_VO_PATH,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True
        })
        
        service = Service()
        logging.info("Trying minimal Chrome setup")
        return webdriver.Chrome(service=service, options=minimal_options)
    except Exception as e:
        logging.error(f"Minimal Chrome setup failed: {e}")
        raise

def wait_for_element(driver, by, value, timeout=30):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def wait_and_click(driver, by, value, timeout=30):
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    driver.execute_script("arguments[0].click();", element)

def is_audio_ready(driver):
    try:
        # Check for the presence of the audio player and absence of loading indicators
        audio_player = driver.find_element(By.XPATH, "//div[contains(@class, 'audio-player')]")
        loading_indicators = driver.find_elements(By.XPATH, "//div[contains(@class, 'loading')]")
        
        # Check if the Export button is present and enabled
        export_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Export')]")
        
        is_ready = (
            audio_player.is_displayed() and 
            len(loading_indicators) == 0 and 
            export_button.is_enabled()
        )
        
        if is_ready:
            logging.info("Audio is ready for export")
        return is_ready
        
    except (NoSuchElementException, StaleElementReferenceException) as e:
        logging.debug(f"Audio not ready yet: {str(e)}")
        return False

def wait_for_audio_generation(driver, timeout=300):
    start_time = time.time()
    while time.time() - start_time < timeout:
        if is_audio_ready(driver):
            logging.info("Audio generation completed")
            return True
        time.sleep(2)
        logging.info("Waiting for audio generation to complete...")
    
    logging.error("Timeout waiting for audio generation")
    return False

def try_export(driver):
    try:
        # First check if driver session is still alive
        if not is_driver_alive(driver):
            logging.warning("Driver session is dead in try_export")
            return False
            
        # Handle any error dialogs before clicking export
        handle_error_dialogs(driver)
            
        # Try to click Export button
        logging.info("Looking for Export button...")
        export_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Export')]")
        if not export_buttons:
            logging.info("Export button not found")
            return False
            
        export_button = export_buttons[0]
        if not export_button.is_enabled():
            logging.info("Export button found but not enabled")
            return False

        logging.info("Found enabled Export button, attempting to click...")
        driver.execute_script("arguments[0].click();", export_button)
        logging.info("Clicked Export button")
        
        # Handle any error dialogs after clicking
        handle_error_dialogs(driver)
        return True

    except Exception as e:
        error_msg = str(e).lower()
        if any(error_text in error_msg for error_text in ["connection", "session", "refused", "10061"]):
            logging.error(f"Connection error in export process: {str(e)}")
            logging.error("This indicates the WebDriver session has been lost")
        else:
            logging.error(f"Error in export process: {str(e)}")
        return False

def try_generate(driver):
    try:
        # First check if driver session is still alive
        if not is_driver_alive(driver):
            logging.warning("Driver session is dead in try_generate")
            return False
            
        # Handle any error dialogs before clicking generate
        handle_error_dialogs(driver)
            
        # Just look for and click the generate button - don't clear text
        generate_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Generate')]"))
        )
        if generate_button.is_enabled():
            driver.execute_script("arguments[0].click();", generate_button)
            logging.info("Clicked Generate button")
            
            # Handle any error dialogs after clicking
            handle_error_dialogs(driver)
            return True
    except Exception as e:
        error_msg = str(e).lower()
        if any(error_text in error_msg for error_text in ["connection", "session", "refused", "10061"]):
            logging.error(f"Connection error in generate process: {e}")
            logging.error("This indicates the WebDriver session has been lost")
        else:
            logging.debug(f"Generate button not clickable: {e}")
        return False

def get_recent_download(driver, desktop_dir, text_snippet, doc_title, notion_handler, record_id, timeout=60):
    """Check for a recent file and rename it."""
    default_download_path = os.path.join(os.path.expanduser("~"), "Downloads")
    celebrity_vo_path = "E:\\Celebrity Voice Overs"
    
    # Record the start time to only look for files created after this point
    start_time = time.time()
    end_time = start_time + timeout
    last_check_files = set()
    
    while time.time() < end_time:
        try:
            # First check Downloads folder
            if os.path.exists(default_download_path):
                current_files = set(os.listdir(default_download_path))
                new_files = current_files - last_check_files
                
                for filename in new_files:
                    if filename.startswith("PlayAI_") and filename.endswith(".wav"):
                        source_path = os.path.join(default_download_path, filename)
                        
                        # Wait for file to be completely downloaded
                        file_size = -1
                        for _ in range(10):  # Check file size stability
                            try:
                                current_size = os.path.getsize(source_path)
                                if current_size == file_size and current_size > 0:  # File size hasn't changed and is not empty
                                    break
                                file_size = current_size
                                time.sleep(1)
                            except Exception as e:
                                logging.debug(f"Error checking file size: {e}")
                                time.sleep(1)
                        
                        # Get file extension
                        _, ext = os.path.splitext(filename)
                        # Create new filename with doc title
                        new_filename = f"{doc_title}{ext}"
                        target_path = os.path.join(celebrity_vo_path, new_filename)
                        
                        try:
                            # If a file with the same name exists in target directory, add a number
                            counter = 1
                            while os.path.exists(target_path):
                                new_filename = f"{doc_title}_{counter}{ext}"
                                target_path = os.path.join(celebrity_vo_path, new_filename)
                                counter += 1
                            
                            # Try to move the file
                            max_retries = 3
                            for retry in range(max_retries):
                                try:
                                    # First try to copy the file
                                    import shutil
                                    shutil.copy2(source_path, target_path)
                                    logging.info(f"Copied file to: {target_path}")
                                    
                                    # If copy successful, try to remove the original
                                    try:
                                        os.remove(source_path)
                                        logging.info("Removed original file from Downloads folder")
                                    except Exception as e:
                                        logging.warning(f"Could not remove original file: {e}")
                                    
                                    break
                                except Exception as e:
                                    if retry == max_retries - 1:
                                        raise
                                    logging.warning(f"Move attempt {retry + 1} failed: {e}")
                                    time.sleep(2)
                            
                            # Mark Notion record as complete
                            try:
                                notion_handler.update_notion_checkboxes(record_id, voiceover=True, ready_to_be_edited=True)
                                logging.info("Marked Notion record as complete")
                            except Exception as e:
                                logging.error(f"Error updating Notion: {e}")
                                return False
                            
                            return True
                                
                        except Exception as e:
                            logging.error(f"Error processing file: {str(e)}")
                            return False
                
                # Update last checked files
                last_check_files = current_files
            
            # Log progress
            remaining = end_time - time.time()
            if remaining > 0:
                logging.info(f"Waiting for download... {int(remaining)}s remaining")
            
            time.sleep(1)
            
        except Exception as e:
            logging.error(f"Error checking downloads: {e}")
            time.sleep(1)
    
    logging.error("Download timeout reached")
    return False

def save_cookies(driver, cookie_file):
    """Save current browser cookies to a file"""
    try:
        # Wait a moment to ensure we're fully logged in
        time.sleep(2)
        
        # Get all cookies
        cookies = driver.get_cookies()
        
        # Filter cookies for play.ht domain
        playht_cookies = [cookie for cookie in cookies if '.play.ht' in cookie.get('domain', '')]
        
        if playht_cookies:
            # Create directory if it doesn't exist
            cookie_dir = os.path.dirname(cookie_file)
            if cookie_dir and not os.path.exists(cookie_dir):
                os.makedirs(cookie_dir)
            
            # Save cookies
            with open(cookie_file, 'wb') as f:
                pickle.dump(playht_cookies, f)
            logging.info(f"Saved {len(playht_cookies)} cookies to {cookie_file}")
            return True
        else:
            logging.warning("No Play.ht cookies found to save")
            return False
    except Exception as e:
        logging.error(f"Error saving cookies: {e}")
        return False

def load_cookies(driver, cookie_file, domain=None):
    """Load cookies from file"""
    if not os.path.exists(cookie_file):
        logging.warning(f"Cookie file {cookie_file} not found")
        return
        
    try:
        with open(cookie_file, 'rb') as f:
            cookies = pickle.load(f)
            
        # Load only domain-specific cookies if domain is specified
        if domain:
            cookies = [cookie for cookie in cookies if domain in cookie.get('domain', '')]
            
        for cookie in cookies:
            try:
                # Remove problematic attributes that might cause issues
                if 'expiry' in cookie:
                    del cookie['expiry']
                driver.add_cookie(cookie)
            except Exception as e:
                logging.warning(f"Error adding cookie: {e}")
                
        logging.info(f"Loaded {len(cookies)} cookies from {cookie_file}")
    except Exception as e:
        logging.error(f"Error loading cookies from {cookie_file}: {e}")
        return False
    
    return True

def handle_playht_login(driver):
    """Handle PlayHT login if needed"""
    try:
        # Check if we have cookies
        if os.path.exists(PLAYHT_COOKIES_FILE):
            logging.info("Found existing cookie file, loading cookies and proceeding...")
            try:
                # Load cookies
                with open(PLAYHT_COOKIES_FILE, 'rb') as f:
                    cookies = pickle.load(f)
                
                # First navigate to play.ht to set cookies
                driver.get("https://play.ht")
                time.sleep(2)
                
                # Add cookies
                for cookie in cookies:
                    try:
                        clean_cookie = {
                            'name': cookie.get('name'),
                            'value': cookie.get('value'),
                            'domain': '.play.ht',
                            'path': '/'
                        }
                        driver.add_cookie(clean_cookie)
                    except Exception as e:
                        logging.debug(f"Skipping invalid cookie: {e}")
                
                logging.info("Cookies loaded successfully")
                return True
                
            except Exception as e:
                logging.error(f"Error loading cookies: {e}")
                return False
        
        # If we get here, we need manual login because no cookies exist
        logging.info("No cookie file found, manual login required...")
        print("\n*** MANUAL LOGIN REQUIRED ***")
        print("Please log in using your Google account in the browser window")
        print("Click the 'Log in with Google' button")
        print("You have 60 seconds to complete the login")
        print("The script will save your login cookies after successful login\n")
        
        # Navigate to login page
        driver.get(PLAYHT_LOGIN_URL)
        time.sleep(2)
        
        # Try to click the Google login button
        try:
            google_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Log in with Google') or contains(@class, 'google')]"))
            )
            google_button.click()
            logging.info("Clicked Google login button")
        except Exception as e:
            logging.warning(f"Could not click Google login button automatically: {e}")
            print("Please click the 'Log in with Google' button manually")

        # Wait for 60 seconds with countdown
        wait_time = 60
        start_time = time.time()
        logged_in = False

        while time.time() - start_time < wait_time:
            remaining = int(wait_time - (time.time() - start_time))
            print(f"\rTime remaining: {remaining} seconds...", end='', flush=True)
            
            # Check if we're logged in
            try:
                if driver.find_element(By.XPATH, "//div[@role='textbox']").is_displayed():
                    logged_in = True
                    print("\nLogin detected!")
                    break
            except:
                pass
            
            time.sleep(1)
        
        if not logged_in:
            print("\nLogin timeout reached")
            return False
            
        print("\nSaving cookies...")
        
        # Save cookies after manual login
        try:
            # Get all cookies
            all_cookies = driver.get_cookies()
            
            # Clean and filter cookies before saving
            clean_cookies = []
            for cookie in all_cookies:
                if any(domain in str(cookie.get('domain', '')).lower() for domain in ['.play.ht', 'play.ht']):
                    clean_cookie = {
                        'name': cookie.get('name'),
                        'value': cookie.get('value'),
                        'domain': '.play.ht',
                        'path': '/'
                    }
                    clean_cookies.append(clean_cookie)
            
            if clean_cookies:
                with open(PLAYHT_COOKIES_FILE, 'wb') as f:
                    pickle.dump(clean_cookies, f)
                logging.info(f"Saved {len(clean_cookies)} clean cookies to file")
                print(f"Successfully saved {len(clean_cookies)} cookies")
                return True
            else:
                logging.error("No valid cookies found to save")
                return False
                
        except Exception as e:
            logging.error(f"Error saving cookies: {e}")
            return False
            
    except Exception as e:
        logging.error(f"Error in login handling: {e}")
        logging.error(f"Traceback: {traceback.format_exc()}")
        return False

def get_audio_files():
    """Get list of audio files in the directory"""
    try:
        if os.path.exists(CELEBRITY_VO_PATH):
            return [f for f in os.listdir(CELEBRITY_VO_PATH) 
                    if f.endswith(('.wav', '.mp3'))]
        return []
    except Exception as e:
        logging.error(f"Error getting audio files: {e}")
        return []

def wait_for_new_audio_file(timeout=30):
    """Wait for a new audio file by checking directory contents"""
    try:
        start_time = time.time()
        initial_files = set(get_audio_files())
        initial_count = len(initial_files)
        logging.info(f"Starting with {initial_count} audio files")

        while True:
            # Remove timeout check - wait indefinitely
            # Just continue checking for new files

            current_files = set(get_audio_files())
            new_files = current_files - initial_files
            
            if new_files:
                # Get the newest file based on creation time
                new_file = max(new_files, key=lambda f: os.path.getctime(os.path.join(CELEBRITY_VO_PATH, f)))
                new_file_path = os.path.join(CELEBRITY_VO_PATH, new_file)
                logging.info(f"Found new audio file: {new_file_path}")
                return new_file_path

            # Log progress every 5 seconds
            elapsed = int(time.time() - start_time)
            if elapsed % 5 == 0:
                logging.info(f"Still waiting for file... ({elapsed}s elapsed)")
            time.sleep(0.5)

    except Exception as e:
        logging.error(f"Error waiting for new audio file: {e}")
        return None

def handle_error_dialogs(driver):
    """Handle any error dialogs that pop up by clicking OK or Cancel"""
    try:
        # Look for common dialog buttons
        buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'OK') or contains(text(), 'Cancel') or contains(text(), 'Dismiss')]")
        for button in buttons:
            if button.is_displayed():
                try:
                    driver.execute_script("arguments[0].click();", button)
                    logging.info("Clicked dialog button: " + button.text)
                    time.sleep(0.5)  # Short wait after clicking
                except:
                    pass
    except Exception as e:
        logging.debug(f"No error dialogs found or error handling them: {e}")

def process_voiceover(driver, chunks, doc_title, record, target_notion, channel=None):
    try:
        last_generate_click = 0
        refresh_count = 0
        second_refresh_time = 0
        export_clicked = False
        file_found = False

        # Clear the new file queue before starting
        while not new_file_queue.empty():
            new_file_queue.get()

        # Get initial file count for monitoring
        initial_files = set(get_audio_files())
        logging.info(f"Starting with {len(initial_files)} audio files")

        # Determine which voice URL to use
        voice_url = DEFAULT_VOICE_URL
        if channel and channel in VOICE_URLS:
            voice_url = VOICE_URLS[channel]
            logging.info(f"Using voice URL for channel: {channel}")
        else:
            logging.info("Using default voice URL")

        # Navigate directly to the voice URL
        logging.info(f"Navigating directly to voice URL: {voice_url}")
        try:
            driver.get(voice_url)
            # Wait for the page to be fully loaded
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # Handle any error dialogs that might appear
            handle_error_dialogs(driver)
            
            # Verify we're on the correct page
            if "play.ht" not in driver.current_url:
                raise Exception("Failed to navigate to Play.ht")
            
            logging.info(f"Successfully loaded voice URL: {driver.current_url}")
        except Exception as e:
            logging.error(f"Navigation error: {e}")
            return False

        # Wait for editor to be fully loaded and interactive
        try:
            # Wait for editor container
            editor_container = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'editor')]"))
            )
            logging.info("Editor container found")
            
            # Wait for actual editor textbox
            editor = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//div[@role='textbox']"))
            )
            logging.info("Editor textbox found")
            
            # Wait for editor to be clickable
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@role='textbox']"))
            )
            logging.info("Editor is clickable")
            
            # Clear the editor using keyboard shortcuts
            driver.execute_script("arguments[0].focus();", editor)
            time.sleep(0.5)
            actions = ActionChains(driver)
            if sys.platform == "darwin":
                actions.key_down(Keys.COMMAND).send_keys('a').key_up(Keys.COMMAND)
            else:
                actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)
            actions.send_keys(Keys.DELETE).perform()
            time.sleep(0.5)
            
            # Input text chunks with clipboard
            for i, chunk in enumerate(chunks):
                try:
                    # Focus the editor
                    driver.execute_script("arguments[0].focus();", editor)
                    
                    if i > 0:
                        actions = ActionChains(driver)
                        actions.send_keys(Keys.RETURN).perform()
                        time.sleep(0.2)
                    
                    # Use JavaScript to paste the chunk
                    script = """
                        var textarea = arguments[0];
                        var text = arguments[1];
                        var dataTransfer = new DataTransfer();
                        dataTransfer.setData('text', text);
                        textarea.dispatchEvent(new ClipboardEvent('paste', {
                            clipboardData: dataTransfer,
                            bubbles: true,
                            cancelable: true
                        }));
                    """
                    driver.execute_script(script, editor, chunk)
                    time.sleep(0.2)
                    
                    # Verify using the editor's text content
                    current_text = editor.text
                    if not chunk.strip() in current_text:
                        # Try alternative paste method if first attempt failed
                        actions = ActionChains(driver)
                        if sys.platform == "darwin":
                            actions.key_down(Keys.COMMAND).send_keys('v').key_up(Keys.COMMAND)
                        else:
                            actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL)
                        actions.perform()
                        time.sleep(0.2)
                        
                        # Verify again
                        current_text = editor.text
                        if not chunk.strip() in current_text:
                            logging.error(f"Failed to verify chunk {i+1}")
                            return False
                    
                    logging.info(f"Successfully added chunk {i+1}")
                    
                except Exception as e:
                    logging.error(f"Error adding chunk {i+1}: {str(e)}")
                    return False
            
            # Wait for Generate button to be clickable
            generate_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Generate')]"))
            )
            logging.info("Generate button is ready")
            
            # Initial Generate click
            if try_generate(driver):
                last_generate_click = time.time()
                logging.info("Clicked Generate button initially")
            else:
                logging.error("Failed to click initial Generate button")
                return False

            # Main processing loop
            while True:
                try:
                    current_time = time.time()

                    # Handle refreshes (after 60 seconds from initial generate)
                    if refresh_count < 2 and current_time - last_generate_click >= 60:
                        # Store the original chunks text to check after refresh
                        original_text = ""
                        try:
                            editor_before = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//div[@role='textbox']"))
                            )
                            original_text = editor_before.text
                            logging.info("Saved original text before refresh")
                        except Exception as e:
                            logging.warning(f"Could not get original text: {e}")

                        refresh_count += 1
                        logging.info(f"Refreshing page (refresh {refresh_count}/2)")
                        driver.refresh()
                        time.sleep(5)

                        if refresh_count == 2:
                            # After second refresh, try to click generate again
                            if try_generate(driver):
                                second_refresh_time = current_time
                                logging.info("Clicked generate after second refresh, waiting 20 seconds before export attempts")
                            else:
                                logging.warning("Failed to click generate after second refresh")

                    # After second refresh and 20 seconds, start trying export every 10 seconds
                    if refresh_count == 2 and not file_found and current_time - second_refresh_time >= 20:
                        # Check for new files
                        current_files = set(get_audio_files())
                        new_files = current_files - initial_files
                        
                        if new_files:
                            file_found = True
                            new_file = max(new_files, key=lambda f: os.path.getctime(os.path.join(CELEBRITY_VO_PATH, f)))
                            new_file_path = os.path.join(CELEBRITY_VO_PATH, new_file)
                            logging.info(f"New file detected: {new_file_path}")
                            
                            # Wait a moment for file to be fully written
                            time.sleep(3)
                            
                            # Process the file (rename, update Notion, etc.)
                            # ... (keep your existing file processing code) ...
                            return True
                        
                        # If no file found yet, try clicking export
                        last_export_time = getattr(process_voiceover, 'last_export_time', 0)
                        if current_time - last_export_time >= 10:  # Try every 10 seconds
                            logging.info("Attempting export...")
                            if try_export(driver):
                                process_voiceover.last_export_time = current_time
                                logging.info("Clicked Export button")
                            # Handle any error dialogs that might appear
                            handle_error_dialogs(driver)

                    time.sleep(0.5)  # Short sleep to prevent CPU overuse

                except Exception as e:
                    logging.error(f"Error in processing loop: {e}")
                    if driver is None:  # If browser is already closed, stop retrying
                        return False
                    time.sleep(1)

        except Exception as e:
            logging.error(f"Error during text preparation: {e}")
            return False
            
    except Exception as e:
        logging.error(f"Error in process_voiceover: {e}")
        return False
    finally:
        # Make sure to quit the browser if something goes wrong
        try:
            if driver:
                driver.quit()
                logging.info("Cleaned up browser in finally block")
        except:
            pass

def is_driver_alive(driver):
    """Check if the WebDriver session is still alive and responsive."""
    if driver is None:
        return False
    
    try:
        # Try a simple operation to test if driver is responsive
        driver.current_url  # This will fail if session is dead
        return True
    except Exception as e:
        logging.warning(f"Driver session appears to be dead: {e}")
        return False

def sanitize_filename(filename):
    """
    Sanitize filename by removing or replacing characters that are not allowed in Windows filenames.
    
    Windows doesn't allow: < > : " | ? * / \
    Also removes leading/trailing spaces and dots, and limits length.
    """
    if not filename:
        return "untitled"
    
    # Replace invalid characters with underscores or remove them
    invalid_chars = ['<', '>', ':', '"', '|', '?', '*', '/', '\\']
    sanitized = filename
    
    for char in invalid_chars:
        sanitized = sanitized.replace(char, '_')
    
    # Replace multiple consecutive underscores with single underscore
    import re
    sanitized = re.sub(r'_+', '_', sanitized)
    
    # Remove leading/trailing spaces and dots
    sanitized = sanitized.strip('. ')
    
    # Limit filename length (Windows has 255 char limit, but let's be conservative)
    if len(sanitized) > 200:
        sanitized = sanitized[:200]
    
    # If empty after sanitization, use default
    if not sanitized:
        sanitized = "untitled"
    
    logging.info(f"Sanitized filename: '{filename}' -> '{sanitized}'")
    return sanitized

def mark_as_processed(table, record_id):
    table.update(record_id, {"Voiceover": True})
    logging.info(f"Marked record {record_id} as processed")

def clean_script(script):
    # List of variations to remove (case-insensitive)
    variations = [
        "Real Sound",
        "real sound",
        "Real Sound Clip",
        "real sound clip"
    ]
    
    # Clean the text
    cleaned_text = script
    for variation in variations:
        cleaned_text = cleaned_text.replace(variation, "")
    
    # Remove any double spaces or extra whitespace that might be left
    cleaned_text = ' '.join(cleaned_text.split())
    
    logging.info("Removed 'Real Sound' variations from script")
    return cleaned_text

def preprocess_text(text):
    """Preprocess text to ensure chunks of ~150 words ending with full stops"""
    # First clean the text
    text = remove_whitespace(text)
    
    # Split into sentences
    sentences = []
    current = []
    
    # Split by potential sentence endings
    for word in text.split():
        current.append(word)
        if word.endswith('.') or word.endswith('!') or word.endswith('?'):
            sentences.append(' '.join(current))
            current = []
    
    # Add any remaining text as a sentence
    if current:
        sentences.append(' '.join(current))
    
    # Group sentences into chunks of ~150 words
    chunks = []
    current_chunk = []
    current_word_count = 0
    
    for sentence in sentences:
        sentence_words = len(sentence.split())
        
        # If adding this sentence would exceed limit, save chunk and start new one
        if current_word_count + sentence_words > 150:
            if current_chunk:  # Only save if we have content
                chunk_text = ' '.join(current_chunk)
                chunks.append(chunk_text)
                # Log the chunk details
                logging.info(f"Created chunk with {len(chunk_text.split())} words")
            current_chunk = [sentence]
            current_word_count = sentence_words
        else:
            current_chunk.append(sentence)
            current_word_count += sentence_words
    
    # Add the last chunk if it exists
    if current_chunk:
        chunk_text = ' '.join(current_chunk)
        chunks.append(chunk_text)
        logging.info(f"Created final chunk with {len(chunk_text.split())} words")
    
    # Verify all chunks
    for i, chunk in enumerate(chunks, 1):
        word_count = len(chunk.split())
        logging.info(f"Chunk {i}: {word_count} words")
        if word_count > 150:
            logging.warning(f"Chunk {i} exceeds 150 words: {word_count} words")
    
    return chunks

def get_google_creds():
    """Get or refresh Google API credentials"""
    creds = None
    
    # Delete the existing token file if it exists and is invalid
    if os.path.exists(TOKEN_PATH):
        try:
            with open(TOKEN_PATH, 'rb') as token:
                creds = pickle.load(token)
            
            # If credentials are invalid and can't be refreshed, delete the token file
            if not creds or not creds.valid:
                if not creds or not creds.refresh_token or not creds.expired:
                    logging.info("Removing invalid token file")
                    os.remove(TOKEN_PATH)
                    creds = None
        except Exception as e:
            logging.error(f"Error loading credentials, removing token file: {e}")
            os.remove(TOKEN_PATH)
            creds = None
    
    # If no valid credentials available, run the OAuth flow
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                logging.error(f"Error refreshing credentials: {e}")
                os.remove(TOKEN_PATH)
                creds = None
        
        if not creds:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
                creds = flow.run_local_server(port=0)
                # Save the new credentials
                with open(TOKEN_PATH, 'wb') as token:
                    pickle.dump(creds, token)
                logging.info("New credentials obtained and saved")
            except Exception as e:
                logging.error(f"Error running OAuth flow: {e}")
                raise
    
    return creds

def get_doc_content(doc_url):
    """Fetch content and title from a document URL"""
    try:
        parsed_url = urlparse(doc_url)
        
        if "docs.google.com" in parsed_url.netloc:
            # Extract document ID from Google Docs URL, handling both u/0 and direct paths
            path_parts = parsed_url.path.split('/')
            doc_id = None
            for part in path_parts:
                if len(part) > 25:  # Google Doc IDs are typically long strings
                    doc_id = part
                    break
            
            if not doc_id:
                raise ValueError(f"Could not extract document ID from URL: {doc_url}")
            
            logging.info(f"Extracted document ID: {doc_id}")
            
            creds = get_google_creds()
            service = build('docs', 'v1', credentials=creds)
            document = service.documents().get(documentId=doc_id).execute()
            doc_title = document.get('title', '')
            
            content = []
            skip_next = False
            
            real_sound_variants = [
                "REAL SOUND",
                "Real Sound",
                "real sound",
                "REAL SOUND CLIP",
                "Real Sound Clip",
                "real sound clip",
                "[Real Sound]",
                "(Real Sound)",
                "Real Sound:",
                "real sound:",
                "REAL SOUND:",
                "Real sound clip",
                "Real Sound Clip:",
                "REAL SOUND CLIP:"
            ]

            for element in document.get('body').get('content'):
                if 'paragraph' in element:
                    paragraph = element['paragraph']
                    para_segments = []
                    
                    for para_element in paragraph['elements']:
                        if 'textRun' in para_element:
                            text = para_element['textRun']['content']
                            text_style = para_element['textRun'].get('textStyle', {})
                            font_size = text_style.get('fontSize', {}).get('magnitude', 11)
                            
                            # Only skip this specific text segment if it's a headline
                            if font_size <= 13:
                                para_segments.append(text)
                    
                    # Combine all non-headline segments
                    para_text = ''.join(para_segments).strip()
                    
                    # Skip if empty
                    if not para_text:
                        continue
                    
                    # Only skip the "Real Sound" marker itself, not the content after it
                    if any(variant in para_text for variant in real_sound_variants):
                        continue
                    
                    # Add all other content
                    content.append(para_text)
            
            final_content = ' '.join(content)
            final_content = ' '.join(final_content.split())
            
            logging.info(f"Processed content length: {len(final_content)} characters")
            
            return {
                'content': final_content.strip(),
                'title': doc_title
            }
        else:
            response = requests.get(doc_url)
            response.raise_for_status()
            return {
                'content': response.text.strip(),
                'title': ''
            }
            
    except Exception as e:
        logging.error(f"Error fetching doc content: {e}")
        return None

def update_new_script(table, record_id, content, title):
    """Update the New Script and New Title fields with content from doc"""
    try:
        update_fields = {
            "New Script": content,
            "Name": title
        }
        table.update(record_id, update_fields)
        logging.info(f"Updated New Script and Name for record {record_id}")
        return True
    except Exception as e:
        logging.error(f"Error updating fields: {e}")
        return False

def update_airtable_docs(table, docs_url):
    """Create new record in Airtable with the Google Docs link if it doesn't exist."""
    try:
        # Check if this URL already exists in Airtable
        existing_records = table.all(
            formula=f"{{Docs}} = '{docs_url}'"
        )
        
        if existing_records:
            logging.info(f"URL already exists in Airtable, skipping: {docs_url}")
            return True
            
        # If URL doesn't exist, create new record
        new_record = {
            "Docs": docs_url,
        }
        table.create(new_record)
        logging.info(f"Created new Airtable record with Docs URL: {docs_url}")
        return True
    except Exception as e:
        logging.error(f"Error creating Airtable record: {e}")
        return False

def get_existing_docs_urls(table):
    """Get all existing doc URLs from Airtable to avoid duplicates."""
    try:
        records = table.all()
        existing_urls = set()
        for record in records:
            url = record['fields'].get('Docs', '').strip()
            if url:
                existing_urls.add(url)
        return existing_urls
    except Exception as e:
        logging.error(f"Error getting existing URLs: {e}")
        return set()

class AudioFileHandler(FileSystemEventHandler):
    def __init__(self, initial_count):
        self.initial_count = initial_count
        self.last_created_time = time.time()
        
    def on_created(self, event):
        if not event.is_directory:
            current_time = time.time()
            # Prevent duplicate events by checking time delta
            if current_time - self.last_created_time > 1:
                self.last_created_time = current_time
                file_path = event.src_path
                # Check if it's a wav or mp3 file and exists
                if file_path.endswith(('.wav', '.mp3')) and os.path.exists(file_path):
                    try:
                        # Count current files
                        current_files = len([f for f in os.listdir(CELEBRITY_VO_PATH) 
                                          if f.endswith(('.wav', '.mp3'))])
                        
                        # If we have more files than we started with
                        if current_files > self.initial_count:
                            logging.info(f"New audio file detected by watchdog: {file_path}")
                            new_file_queue.put(file_path)
                    except Exception as e:
                        logging.error(f"Error processing new file in watchdog: {e}")

def get_initial_file_count():
    """Get the initial count of audio files in the directory"""
    try:
        if os.path.exists(CELEBRITY_VO_PATH):
            count = len([f for f in os.listdir(CELEBRITY_VO_PATH) 
                        if f.endswith(('.wav', '.mp3'))])
            logging.info(f"Initial audio file count: {count}")
            return count
        else:
            os.makedirs(CELEBRITY_VO_PATH)
            logging.info("Created Celebrity Voice Overs directory")
            return 0
    except Exception as e:
        logging.error(f"Error getting initial file count: {e}")
        return 0

def setup_watchdog():
    """Set up watchdog observer for the Celebrity Voice Overs folder"""
    try:
        # Get initial file count
        initial_count = get_initial_file_count()
        
        event_handler = AudioFileHandler(initial_count)
        observer = Observer()
        observer.schedule(event_handler, CELEBRITY_VO_PATH, recursive=False)
        observer.start()
        logging.info(f"Started watchdog observer for {CELEBRITY_VO_PATH}")
        
        # Verify the observer is running
        if not observer.is_alive():
            logging.error("Observer failed to start")
            return None
            
        return observer
    except Exception as e:
        logging.error(f"Error setting up watchdog: {e}")
        return None

def main():
    # Initialize Notion handler
    notion_handler = TargetNotionHandler(NOTION_TOKEN, NOTION_DATABASE_ID)
    driver = None
    max_retries = 3
    script_records = []  # Initialize script_records
    new_docs_added = False  # Initialize new_docs_added

    while True:
        try:
            # Set up watchdog observer
            observer = setup_watchdog()
            if not observer:
                logging.error("Failed to set up watchdog observer")
                time.sleep(60)
                continue

            # Check if there are any voiceovers to process
            voiceover_records = notion_handler.get_records_for_voiceover()
            logging.info(f"Found {len(voiceover_records)} records that need voiceover")
            
            if voiceover_records:
                # Process all voiceover records
                for record in voiceover_records:
                    title_prop = record['properties'].get('New Title', {}).get('title', [])
                    doc_title = title_prop[0].get('text', {}).get('content', '').strip() if title_prop else ''
                    
                    # Get content from the page body
                    blocks = notion_handler.notion.blocks.children.list(record['id'])
                    script = ''
                    for block in blocks.get('results', []):
                        if block['type'] == 'paragraph':
                            for text in block['paragraph']['rich_text']:
                                script += text.get('text', {}).get('content', '')
                    
                    if script and doc_title:
                        logging.info(f"Processing voiceover for: {doc_title}")
                        
                        # Initialize Chrome if not already running or if session is dead
                        if not driver or not is_driver_alive(driver):
                            if driver:
                                logging.warning("Existing driver session is dead, reinitializing...")
                                try:
                                    driver.quit()
                                except:
                                    pass
                                cleanup_chrome_processes()
                                driver = None
                                time.sleep(5)  # Wait before reinitializing
                            
                            try:
                                driver = setup_chrome_driver()
                                if not driver:
                                    logging.error("Failed to initialize Chrome driver")
                                    time.sleep(60)  # Wait before retrying
                                    continue
                                
                                # Load cookies and handle login
                                if driver:  # Double check driver is valid
                                    load_cookies(driver, PLAYHT_COOKIES_FILE)
                                    if not handle_playht_login(driver):
                                        logging.error("Failed to log in")
                                        cleanup_chrome_processes()
                                        driver = None
                                        time.sleep(60)  # Wait before retrying
                                        continue
                            except Exception as e:
                                logging.error(f"Error setting up Chrome: {e}")
                                if driver:
                                    cleanup_chrome_processes()
                                    driver = None
                                time.sleep(60)  # Wait before retrying
                                continue

                        # Get channel from properties if available
                        channel = None
                        if "Channel" in record['properties']:
                            channel_prop = record['properties']["Channel"].get("select", {})
                            if channel_prop:
                                channel = channel_prop.get("name")
                        
                        # Process the voiceover
                        chunks = split_text(script)
                        if not chunks:
                            logging.warning(f"No valid chunks found for {doc_title}")
                            continue
                            
                        start_time = time.time()
                        success = process_voiceover(driver, chunks, doc_title, record, notion_handler, channel=channel)
                        
                        if success:
                            # Update both Voiceover and Ready to Be Edited checkboxes
                            checkbox_update_success = notion_handler.update_notion_checkboxes(
                                record['id'], 
                                voiceover=True,
                                ready_to_be_edited=True
                            )
                            
                            if checkbox_update_success:
                                logging.info(f"Successfully processed and updated record: {doc_title}")
                            else:
                                logging.warning(f"Failed to update checkboxes for: {doc_title}")
                            
                            processing_time = time.time() - start_time
                            logging.info(f"Processing time for {doc_title}: {processing_time:.2f} seconds")
                        else:
                            logging.error(f"Failed to process voiceover for: {doc_title}")
                            cleanup_chrome_processes()
                            driver = None
                            time.sleep(60)  # Wait before retrying
                    else:
                        logging.warning(f"Skipping record {record['id']} - missing script or title")
            
            # Stop the watchdog observer
            if observer:
                observer.stop()
                observer.join()
            
            time.sleep(10)  # Short sleep between checks
            
        except Exception as e:
            logging.error(f"Error in main loop: {str(e)}")
            if driver:
                cleanup_chrome_processes()
                driver = None
            time.sleep(60)  # Wait before retrying

if __name__ == "__main__":
    main() 