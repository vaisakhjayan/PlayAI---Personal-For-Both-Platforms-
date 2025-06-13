from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import os
import time
from contentpaster import start_content_pasting
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Set up Chrome options
chrome_options = Options()
# Specify the user data directory for Chrome profiles
user_data_dir = os.path.expanduser("~/Desktop/Chrome Profile 1")
chrome_options.add_argument(f"user-data-dir={user_data_dir}")

try:
    # Initialize Chrome driver with the specified options
    driver = webdriver.Chrome(options=chrome_options)

    # Navigate to the Play.ht URL
    url = "https://app.play.ht/studio/file/bblKsDe7AaaeE5P74h7t?voice=s3://voice-cloning-zero-shot/541946ca-d3c9-49c7-975b-09a4e42a991f/original/manifest.json"
    driver.get(url)

    # Wait a bit for the page to load
    time.sleep(5)

    # Start content pasting process
    logging.info("Starting content pasting process...")
    success = start_content_pasting(driver)
    
    if success:
        logging.info("Content pasted successfully!")
    else:
        logging.error("Failed to paste content")

    # Keep the browser window open until manually closed
    input("Press Enter to close the browser...")

except Exception as e:
    logging.error(f"An error occurred: {str(e)}")

finally:
    # Clean up
    try:
        if driver:
            driver.quit()
    except:
        pass
