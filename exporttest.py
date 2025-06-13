from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import os
import time
import logging
from export import export_audio

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Disable other loggers
for logger_name in logging.root.manager.loggerDict:
    if logger_name != "__main__" and not logger_name.startswith("export"):
        logging.getLogger(logger_name).setLevel(logging.WARNING)

# Set up Chrome options
chrome_options = Options()

# Configure download behavior
chrome_options.add_experimental_option('prefs', {
    'download.default_directory': "E:\\Celebrity Voice Overs",
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True
})

# Add other necessary Chrome options
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

# Specify the user data directory for Chrome profiles
user_data_dir = os.path.expanduser("~/Desktop/Chrome Profile 1")
chrome_options.add_argument(f"user-data-dir={user_data_dir}")

try:
    logging.info("Starting export test...")
    # Initialize Chrome driver with the specified options
    driver = webdriver.Chrome(options=chrome_options)

    # Navigate to the Play.ht URL
    url = "https://app.play.ht/studio/file/5awrBo1FeAUVjaSGUk0i?voice=s3://voice-cloning-zero-shot/541946ca-d3c9-49c7-975b-09a4e42a991f/original/manifest.json"
    logging.info(f"Navigating to URL: {url}")
    driver.get(url)

    # Wait for the page to load
    logging.info("Waiting for page to load...")
    time.sleep(5)  # Give some time for the page to load

    # Try to export
    logging.info("Starting export process...")
    result = export_audio(driver)
    if result:
        logging.info("Export completed successfully")
    else:
        logging.error("Export failed")

    # Keep the browser window open until manually closed
    input("Press Enter to close the browser...")

except Exception as e:
    logging.error(f"Test failed with error: {str(e)}")

finally:
    # Clean up
    try:
        if 'driver' in locals():
            logging.info("Closing browser...")
            driver.quit()
    except Exception as e:
        logging.error(f"Error closing browser: {str(e)}")
