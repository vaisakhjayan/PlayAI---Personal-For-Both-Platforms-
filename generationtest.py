from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import os
import time
from generationlogic import verify_and_generate

# Set up Chrome options
chrome_options = Options()

# Specify the user data directory for Chrome profiles
user_data_dir = os.path.expanduser("~/Desktop/Chrome Profile 1")
chrome_options.add_argument(f"user-data-dir={user_data_dir}")

try:
    # Initialize Chrome driver with the specified options
    driver = webdriver.Chrome(options=chrome_options)

    # Navigate to the Play.ht URL
    url = "https://app.play.ht/studio/file/5awrBo1FeAUVjaSGUk0i?voice=s3://voice-cloning-zero-shot/541946ca-d3c9-49c7-975b-09a4e42a991f/original/manifest.json"
    driver.get(url)

    # Wait for the page to load
    time.sleep(5)  # Give some time for the page to load

    # Try to generate and reload
    verify_and_generate(driver)

    # Keep the browser window open until manually closed
    input("Press Enter to close the browser...")

except Exception as e:
    pass

finally:
    # Clean up
    try:
        if 'driver' in locals():
            driver.quit()
    except:
        pass
