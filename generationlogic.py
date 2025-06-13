from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import json
import logging
import time
import sys

# Configurable delays (in seconds)
DELAY_BEFORE_GENERATE = 5  # Delay before clicking Generate All
DELAY_BEFORE_RELOAD = 60    # Delay before reloading the page after clicking Generate
DELAY_AFTER_RELOAD = 2      # Delay after reloading the page before clicking generate again

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
                    time.sleep(0.5)
                except:
                    pass
    except Exception as e:
        logging.debug(f"No error dialogs found or error handling them: {e}")

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

def verify_and_generate(driver):
    """
    Clicks generate, and if successful:
    - Waits exactly 15 seconds
    - Reloads the page
    - Waits 2 seconds
    - Checks for and clicks generate again if needed
    
    Args:
        driver: Selenium WebDriver instance
    
    Returns:
        bool: True if generation was triggered, False otherwise
    """
    try:
        # First check if driver session is still alive
        if not is_driver_alive(driver):
            logging.warning("Driver session is dead in verify_and_generate")
            return False
            
        # Try first generate click
        first_generate = try_generate(driver)
        
        # Only proceed with reload if generate was successful
        if first_generate:
            logging.info("Successfully clicked generate button first time")
            
            # Wait exactly 15 seconds from successful click
            logging.info(f"Waiting {DELAY_BEFORE_RELOAD} seconds before reload...")
            time.sleep(DELAY_BEFORE_RELOAD)
            
            # Reload the page
            driver.refresh()
            logging.info("Page reloaded")
            
            # Wait 2 seconds after reload
            time.sleep(DELAY_AFTER_RELOAD)
            logging.info(f"Waited {DELAY_AFTER_RELOAD} seconds after reload")
            
            # Try second generate click
            try:
                second_generate = try_generate(driver)
                if second_generate:
                    logging.info("Successfully clicked generate button second time")
                else:
                    logging.info("Generate button not needed after reload - text likely already processed")
            except Exception as e:
                logging.info("Generate button not available after reload")
                
            return True  # Return true if first generate was successful
        else:
            logging.warning("Failed to click generate button first time, skipping reload sequence")
            return False
        
    except Exception as e:
        error_msg = str(e).lower()
        if any(error_text in error_msg for error_text in ["connection", "session", "refused", "10061"]):
            logging.error(f"Connection error in generate process: {e}")
            logging.error("This indicates the WebDriver session has been lost")
        else:
            logging.debug(f"Generate button not clickable: {e}")
        return False

def reload_page(driver):
    """
    Reloads the page and waits for specified delay.
    Returns True if successful, False otherwise.
    """
    try:
        # First check if driver session is still alive
        if not is_driver_alive(driver):
            logging.warning("Driver session is dead in reload_page")
            return False
            
        logging.info(f"Waiting {DELAY_BEFORE_RELOAD} seconds before reload...")
        time.sleep(DELAY_BEFORE_RELOAD)
        
        # Reload the page
        driver.refresh()
        logging.info("Page reloaded")
        
        # Wait after reload
        logging.info(f"Waiting {DELAY_AFTER_RELOAD} seconds after reload...")
        time.sleep(DELAY_AFTER_RELOAD)
        
        return True
        
    except Exception as e:
        logging.error(f"Error reloading page: {e}")
        return False

def try_generate(driver):
    """Try to click the Generate button with multiple fallback methods"""
    try:
        # First check if driver session is still alive
        if not is_driver_alive(driver):
            logging.warning("Driver session is dead in try_generate")
            return False
            
        # Handle any error dialogs before clicking generate
        handle_error_dialogs(driver)
            
        # Try multiple methods to find and click the Generate button
        try:
            # Try first with exact text match
            generate_button = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//button[text()='Generate']"))
            )
            logging.info("Found Generate button with exact text match")
        except:
            try:
                # Try with contains() if exact match fails
                generate_button = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Generate')]"))
                )
                logging.info("Found Generate button with contains() match")
            except:
                # Try with case-insensitive match as last resort
                try:
                    generate_button = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'GENERATE', 'generate'), 'generate')]"))
                    )
                    logging.info("Found Generate button with case-insensitive match")
                except:
                    logging.info("Generate button not found - may not be needed")
                    return False
        
        # Wait a moment for any animations to complete
        time.sleep(1)
        
        # Ensure button is visible and clickable
        if generate_button.is_displayed() and generate_button.is_enabled():
            # Try multiple click methods
            try:
                # Try regular click first
                generate_button.click()
                logging.info("Clicked Generate button with regular click")
            except:
                try:
                    # Try JavaScript click if regular click fails
                    driver.execute_script("arguments[0].click();", generate_button)
                    logging.info("Clicked Generate button with JavaScript")
                except:
                    # Try Actions chain as last resort
                    ActionChains(driver).move_to_element(generate_button).click().perform()
                    logging.info("Clicked Generate button with Action Chains")
            
            # Handle any error dialogs after clicking
            handle_error_dialogs(driver)
            return True
        else:
            logging.info("Generate button found but not clickable - may not be needed")
            return False
            
    except Exception as e:
        error_msg = str(e).lower()
        if any(error_text in error_msg for error_text in ["connection", "session", "refused", "10061"]):
            logging.error(f"Connection error in generate process: {e}")
            logging.error("This indicates the WebDriver session has been lost")
        else:
            logging.info("Generate button not available - may not be needed")
        return False
