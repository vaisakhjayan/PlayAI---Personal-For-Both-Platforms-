import os
import time
import logging
import psutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import warnings
from notion import Colors, log
from platformconfig import get_chrome_profile_path

# Suppress unnecessary warnings
warnings.filterwarnings('ignore', category=DeprecationWarning)
logging.getLogger('WDM').setLevel(logging.CRITICAL)

# Configuration
PROFILE_DIR = get_chrome_profile_path()
PLAYHT_URL = "https://app.play.ht/studio/file/EvQQeB7ebXYukIkPClh2?voice=s3://voice-cloning-zero-shot/541946ca-d3c9-49c7-975b-09a4e42a991f/original/manifest.json"

def kill_chrome_instances():
    """Kill any Chrome instances that are using our profile directory."""
    try:
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                # Check if it's a Chrome process
                if proc.info['name'] and 'chrome' in proc.info['name'].lower():
                    cmdline = proc.info['cmdline']
                    if cmdline and any(PROFILE_DIR in arg for arg in cmdline):
                        log(f"Killing Chrome process {proc.info['pid']} using our profile", "info")
                        proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        time.sleep(2)  # Give processes time to close
    except Exception as e:
        log(f"Error killing Chrome instances: {str(e)}", "error")

def setup_chrome():
    """Setup Chrome with specific profile and configuration."""
    try:
        # Kill any existing Chrome instances using our profile
        kill_chrome_instances()
        
        # Create Chrome options
        chrome_options = Options()
        
        # Set up profile directory
        if not os.path.exists(PROFILE_DIR):
            os.makedirs(PROFILE_DIR)
            log("Created new Chrome profile directory", "info")
        
        # Add profile directory to options
        chrome_options.add_argument(f'--user-data-dir={PROFILE_DIR}')
        
        # Add other necessary options
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Set download preferences and authentication settings
        chrome_options.add_experimental_option('prefs', {
            'profile.default_content_setting_values.notifications': 2,
            'credentials_enable_service': True,
            'profile.password_manager_enabled': True,
            'profile.default_content_settings.popups': 0,
            'profile.managed_default_content_settings.images': 1,
            'profile.managed_default_content_settings.javascript': 1
        })
        
        # Initialize the driver silently
        service = Service(ChromeDriverManager().install())
        service.log_path = os.devnull
        driver = webdriver.Chrome(
            service=service,
            options=chrome_options
        )
        
        # Modify navigator.webdriver flag to prevent detection
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        # Navigate to PlayHT
        log("Navigating to Play.ht...", "info")
        driver.get(PLAYHT_URL)
        
        # Give time for the page to load
        time.sleep(5)
        
        log("Chrome setup complete ✨", "success")
        return driver
        
    except Exception as e:
        log(f"Error setting up Chrome: {str(e)}", "error")
        return None

def cleanup_chrome(driver):
    """Safely quit the Chrome browser."""
    try:
        if driver:
            driver.quit()
    except Exception:
        pass

def monitor_chrome():
    """Main function to monitor Chrome operations."""
    driver = None
    try:
        driver = setup_chrome()
        if not driver:
            log("Failed to initialize Chrome", "error")
            return None
            
        # Keep the browser open
        while True:
            time.sleep(1)
            
    except KeyboardInterrupt:
        cleanup_chrome(driver)
    except Exception as e:
        log(f"Chrome monitor error: {str(e)}", "error")
        cleanup_chrome(driver)
    
    return driver

if __name__ == "__main__":
    # Clear screen
    print("\033[H\033[J", end="")
    log("✨ CHROME MONITOR", "header")
    monitor_chrome()
