from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import logging
import time
import os
import json
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from queue import Queue, Empty
import re

# Create a logger specific to this module
logger = logging.getLogger(__name__)

# Configurable paths and delays
CELEBRITY_VO_PATH = "E:\\Celebrity Voice Overs"
CONTENT_JSON_PATH = "JSON Files/content.json"
DELAY_BEFORE_EXPORT = 180  # Delay before clicking Export
EXPORT_TIMEOUT = 6000  # Maximum time to wait for export file

# Queue for new file events
new_file_queue = Queue()

class AudioFileHandler(FileSystemEventHandler):
    def __init__(self, initial_count):
        self.initial_count = initial_count
        self.last_created_time = time.time()
        logger.info(f"Initialized AudioFileHandler with initial count: {initial_count}")
        
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
                            logger.info(f"New audio file detected: {file_path}")
                            new_file_queue.put(file_path)
                    except Exception as e:
                        logger.error(f"Error processing new file: {e}")

def get_initial_file_count():
    """Get the initial count of audio files in the directory"""
    try:
        if os.path.exists(CELEBRITY_VO_PATH):
            count = len([f for f in os.listdir(CELEBRITY_VO_PATH) 
                        if f.endswith(('.wav', '.mp3'))])
            logger.info(f"Initial audio file count: {count}")
            return count
        else:
            os.makedirs(CELEBRITY_VO_PATH)
            logger.info("Created Celebrity Voice Overs directory")
            return 0
    except Exception as e:
        logger.error(f"Error getting initial file count: {e}")
        return 0

def setup_watchdog():
    """Set up watchdog observer for the Celebrity Voice Overs folder"""
    try:
        logger.info("Setting up watchdog observer...")
        # Get initial file count
        initial_count = get_initial_file_count()
        
        event_handler = AudioFileHandler(initial_count)
        observer = Observer()
        observer.schedule(event_handler, CELEBRITY_VO_PATH, recursive=False)
        observer.start()
        logger.info(f"Started watchdog observer for {CELEBRITY_VO_PATH}")
        
        # Verify the observer is running
        if not observer.is_alive():
            logger.error("Observer failed to start")
            return None
            
        return observer
    except Exception as e:
        logger.error(f"Error setting up watchdog: {e}")
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
                    logger.info(f"Clicked dialog button: {button.text}")
                    time.sleep(0.5)  # Short wait after clicking
                except:
                    pass
    except Exception as e:
        logger.debug(f"No error dialogs found or error handling them: {e}")

def is_driver_alive(driver):
    """Check if the WebDriver session is still alive and responsive."""
    if driver is None:
        logger.warning("Driver is None")
        return False
    
    try:
        # Try a simple operation to test if driver is responsive
        driver.current_url  # This will fail if session is dead
        return True
    except Exception as e:
        logger.warning(f"Driver session appears to be dead: {e}")
        return False

def check_for_error_dialog(driver):
    """Check specifically for audio not ready dialog and return True if found"""
    try:
        # Look for error messages or dialog boxes about audio not being ready
        error_messages = driver.find_elements(By.XPATH, 
            "//div[contains(text(), 'not ready') or contains(text(), 'still processing') or contains(text(), 'please wait')]"
        )
        error_buttons = driver.find_elements(By.XPATH, 
            "//button[contains(text(), 'OK') or contains(text(), 'Cancel') or contains(text(), 'Dismiss')]"
        )
        
        if error_messages or error_buttons:
            logger.info("Found audio not ready dialog")
            # Click the button to dismiss the dialog
            for button in error_buttons:
                if button.is_displayed():
                    driver.execute_script("arguments[0].click();", button)
                    logger.info("Dismissed error dialog")
            return True
        return False
    except Exception as e:
        logger.debug(f"Error checking for dialog: {e}")
        return False

def try_export(driver):
    """Try to click the Export button with multiple fallback methods"""
    try:
        # First check if driver session is still alive
        if not is_driver_alive(driver):
            logger.warning("Driver session is dead in try_export")
            return False
            
        # Initial delay before first attempt
        logger.info(f"Waiting {DELAY_BEFORE_EXPORT} seconds before attempting export...")
        for remaining in range(DELAY_BEFORE_EXPORT, 0, -5):
            logger.info(f"Export will begin in {remaining} seconds...")
            time.sleep(5)
        
        # Reload the page after the initial delay
        logger.info("Reloading page before export...")
        driver.refresh()
        
        # Wait for page to reload and stabilize
        logger.info("Waiting 20 seconds after reload...")
        time.sleep(20)
        
        max_retries = 20  # Maximum number of retry attempts (20 * 15 seconds = 5 minutes max)
        retry_count = 0
        
        while retry_count < max_retries:
            # Try to click Export button
            logger.info("Looking for Export button...")
            export_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Export')]")
            
            if not export_buttons:
                logger.info("Export button not found")
                return False
            
            export_button = export_buttons[0]
            if not export_button.is_enabled():
                logger.info("Export button found but not enabled")
                time.sleep(15)
                retry_count += 1
                continue

            logger.info("Found enabled Export button, attempting to click...")
            driver.execute_script("arguments[0].click();", export_button)
            logger.info("Clicked Export button")
            
            # Wait a moment for any error dialog to appear
            time.sleep(2)
            
            # Check for error dialog
            if check_for_error_dialog(driver):
                logger.info(f"Audio not ready, waiting 15 seconds before retry {retry_count + 1}/{max_retries}")
                time.sleep(15)
                retry_count += 1
                continue
            else:
                logger.info("No error dialog found, export appears successful")
                return True

        logger.error("Maximum retry attempts reached, export failed")
        return False

    except Exception as e:
        error_msg = str(e).lower()
        if any(error_text in error_msg for error_text in ["connection", "session", "refused", "10061"]):
            logger.error(f"Connection error in export process: {str(e)}")
            logger.error("This indicates the WebDriver session has been lost")
        else:
            logger.error(f"Error in export process: {str(e)}")
        return False

def get_title_from_json():
    """Get the title from content.json"""
    try:
        with open(CONTENT_JSON_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
            title = data['records'][0]['title']
            logger.info(f"Found title from JSON: {title}")
            return title
    except Exception as e:
        logger.error(f"Error reading title from JSON: {e}")
        return None

def sanitize_filename(filename):
    """Sanitize filename for Windows compatibility"""
    # Remove or replace invalid characters for Windows filenames
    # Invalid characters: < > : " | ? * \ /
    invalid_chars = r'[<>:"|?*\\/]'
    filename = re.sub(invalid_chars, '', filename)
    
    # Replace multiple spaces with single space
    filename = re.sub(r'\s+', ' ', filename)
    
    # Remove leading/trailing whitespace and dots
    filename = filename.strip(' .')
    
    # Limit filename length (Windows has a 255 character limit for the full path)
    # Keep it shorter to be safe with the full path
    max_length = 200
    if len(filename) > max_length:
        filename = filename[:max_length].strip()
    
    # If filename is empty after sanitization, use a default
    if not filename:
        filename = "untitled"
    
    return filename

def rename_new_file(file_path):
    """Rename the new file with the title from content.json"""
    try:
        title = get_title_from_json()
        if not title:
            logger.error("Could not get title from JSON, keeping original filename")
            return file_path
            
        # Sanitize the title for use as filename
        safe_title = sanitize_filename(title)
        logger.info(f"Original title: {title}")
        logger.info(f"Sanitized title: {safe_title}")
            
        # Get file extension from original file
        _, ext = os.path.splitext(file_path)
        
        # Create new filename with sanitized title
        new_filename = f"{safe_title}{ext}"
        new_path = os.path.join(CELEBRITY_VO_PATH, new_filename)
        
        # If file with same name exists, add number
        counter = 1
        while os.path.exists(new_path):
            new_filename = f"{safe_title} ({counter}){ext}"
            new_path = os.path.join(CELEBRITY_VO_PATH, new_filename)
            counter += 1
            
        # Rename the file
        logger.info(f"Attempting to rename: {file_path} -> {new_path}")
        os.rename(file_path, new_path)
        logger.info(f"Successfully renamed file to: {new_filename}")
        return new_path
    except Exception as e:
        logger.error(f"Error renaming file: {e}")
        logger.error(f"Original path: {file_path}")
        return file_path

def wait_for_export_complete():
    """Wait for new export file to appear in the directory"""
    try:
        logger.info(f"Waiting up to {EXPORT_TIMEOUT} seconds for export to complete...")
        start_time = time.time()
        initial_files = get_audio_files()
        logger.info(f"Starting with {len(initial_files)} files")

        while time.time() - start_time < EXPORT_TIMEOUT:
            # Check directory directly
            current_files = get_audio_files()
            new_files = current_files - initial_files
            
            if new_files:
                newest_file = max(new_files, key=lambda f: os.path.getctime(os.path.join(CELEBRITY_VO_PATH, f)))
                newest_path = os.path.join(CELEBRITY_VO_PATH, newest_file)
                
                # Verify file is complete
                if os.path.exists(newest_path) and os.path.getsize(newest_path) > 0:
                    logger.info(f"Export completed successfully: {newest_file}")
                    # Rename the file with title from JSON
                    renamed_path = rename_new_file(newest_path)
                    return True
            
            # Check queue for watchdog events as backup
            try:
                new_file = new_file_queue.get_nowait()
                if new_file and os.path.exists(new_file) and os.path.getsize(new_file) > 0:
                    logger.info(f"Export completed successfully (via watchdog): {new_file}")
                    # Rename the file with title from JSON
                    renamed_path = rename_new_file(new_file)
                    return True
            except Empty:
                pass  # No watchdog events
                
            # Log progress every 5 seconds
            remaining = int(EXPORT_TIMEOUT - (time.time() - start_time))
            if remaining % 5 == 0:
                current_count = len(current_files)
                logger.info(f"Still waiting for export... {remaining}s remaining (Files: {current_count}, New: {len(new_files)})")
            
            time.sleep(1)  # Short sleep between checks
                
        logger.error("Export timeout reached - no file detected")
        return False
        
    except Exception as e:
        logger.error(f"Error waiting for export: {e}")
        return False

def get_audio_files():
    """Get current list of audio files in the directory"""
    try:
        if os.path.exists(CELEBRITY_VO_PATH):
            return set(f for f in os.listdir(CELEBRITY_VO_PATH) 
                      if f.endswith(('.wav', '.mp3')))
        return set()
    except Exception as e:
        logger.error(f"Error getting audio files: {e}")
        return set()

def export_audio(driver):
    """Main export function that coordinates the export process"""
    try:
        logger.info("Starting export process...")
        
        # Get initial file list before anything else
        initial_files = get_audio_files()
        logger.info(f"Initial file count: {len(initial_files)}")
        
        # Set up file monitoring
        observer = setup_watchdog()
        if not observer:
            logger.error("Failed to set up file monitoring")
            return False
            
        try:
            # Try to click export
            if try_export(driver):
                # First check if a new file appeared immediately
                time.sleep(2)  # Short wait for file to be created
                current_files = get_audio_files()
                new_files = current_files - initial_files
                
                if new_files:
                    new_file = max(new_files, key=lambda f: os.path.getctime(os.path.join(CELEBRITY_VO_PATH, f)))
                    logger.info(f"New file detected immediately: {new_file}")
                    renamed_path = rename_new_file(os.path.join(CELEBRITY_VO_PATH, new_file))
                    
                    # Update Notion checkboxes after successful export
                    try:
                        # Get record ID from content.json
                        with open(CONTENT_JSON_PATH, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            record_id = data['records'][0]['id']
                            
                        # Import notion handler here to avoid circular imports
                        from notion import TargetNotionHandler, NOTION_TOKEN, NOTION_DATABASE_ID
                        notion_handler = TargetNotionHandler(NOTION_TOKEN, NOTION_DATABASE_ID)
                        
                        # Update checkboxes
                        notion_handler.update_notion_checkboxes(record_id, voiceover=True, ready_to_be_edited=True)
                        logger.info("Updated Notion checkboxes after successful export")
                    except Exception as e:
                        logger.error(f"Failed to update Notion checkboxes: {e}")
                    
                    driver.quit()
                    return True
                    
                # If no immediate file, wait for the watchdog to detect it
                if wait_for_export_complete():
                    logger.info("Export process completed successfully")
                    
                    # Update Notion checkboxes after successful export
                    try:
                        # Get record ID from content.json
                        with open(CONTENT_JSON_PATH, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            record_id = data['records'][0]['id']
                            
                        # Import notion handler here to avoid circular imports
                        from notion import TargetNotionHandler, NOTION_TOKEN, NOTION_DATABASE_ID
                        notion_handler = TargetNotionHandler(NOTION_TOKEN, NOTION_DATABASE_ID)
                        
                        # Update checkboxes
                        notion_handler.update_notion_checkboxes(record_id, voiceover=True, ready_to_be_edited=True)
                        logger.info("Updated Notion checkboxes after successful export")
                    except Exception as e:
                        logger.error(f"Failed to update Notion checkboxes: {e}")
                    
                    driver.quit()
                    return True
                else:
                    # One final check before giving up
                    current_files = get_audio_files()
                    new_files = current_files - initial_files
                    if new_files:
                        new_file = max(new_files, key=lambda f: os.path.getctime(os.path.join(CELEBRITY_VO_PATH, f)))
                        logger.info(f"New file detected in final check: {new_file}")
                        renamed_path = rename_new_file(os.path.join(CELEBRITY_VO_PATH, new_file))
                        
                        # Update Notion checkboxes after successful export
                        try:
                            # Get record ID from content.json
                            with open(CONTENT_JSON_PATH, 'r', encoding='utf-8') as f:
                                data = json.load(f)
                                record_id = data['records'][0]['id']
                                
                            # Import notion handler here to avoid circular imports
                            from notion import TargetNotionHandler, NOTION_TOKEN, NOTION_DATABASE_ID
                            notion_handler = TargetNotionHandler(NOTION_TOKEN, NOTION_DATABASE_ID)
                            
                            # Update checkboxes
                            notion_handler.update_notion_checkboxes(record_id, voiceover=True, ready_to_be_edited=True)
                            logger.info("Updated Notion checkboxes after successful export")
                        except Exception as e:
                            logger.error(f"Failed to update Notion checkboxes: {e}")
                        
                        driver.quit()
                        return True
                    
                    logger.error("Export file not detected")
                    driver.quit()
                    return False
            else:
                logger.error("Failed to click export button")
                driver.quit()
                return False
                
        finally:
            # Always stop the observer
            if observer:
                logger.info("Stopping watchdog observer...")
                observer.stop()
                observer.join()
                logger.info("Watchdog observer stopped")
                
    except Exception as e:
        logger.error(f"Error in export process: {e}")
        driver.quit()
        return False
