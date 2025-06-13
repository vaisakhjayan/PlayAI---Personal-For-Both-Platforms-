import json
import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import sys

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_content_from_json(json_file_path):
    """Load content from JSON file"""
    try:
        with open(json_file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            if 'records' in data and len(data['records']) > 0:
                return data['records'][0].get('content', '')
            return None
    except Exception as e:
        logging.error(f"Error loading JSON file: {e}")
        return None

def split_into_chunks(text, max_words=150):
    """Split text into chunks of approximately max_words"""
    # Clean the text first
    text = ' '.join(text.split())
    
    # Split into sentences
    sentences = []
    current = []
    
    for word in text.split():
        current.append(word)
        if word.endswith('.') or word.endswith('!') or word.endswith('?'):
            sentences.append(' '.join(current))
            current = []
    
    # Add any remaining words as the last sentence
    if current:
        sentences.append(' '.join(current))
    
    # Group sentences into chunks
    chunks = []
    current_chunk = []
    word_count = 0
    
    for sentence in sentences:
        sentence_words = len(sentence.split())
        
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
    
    return chunks

def paste_content_to_editor(driver):
    """Paste content from JSON into the editor when it's ready"""
    try:
        # Wait for editor to be fully loaded and interactive
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
        
        # Load content from JSON
        content = load_content_from_json('JSON Files/content.json')
        if not content:
            logging.error("No content found in JSON file")
            return False
        
        # Split content into chunks
        chunks = split_into_chunks(content)
        if not chunks:
            logging.error("No chunks created from content")
            return False
        
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
        
        logging.info("Successfully pasted all content to editor")
        return True
        
    except Exception as e:
        logging.error(f"Error in paste_content_to_editor: {e}")
        return False

# This function can be called from your main module when the website is loaded
def start_content_pasting(driver):
    """Main function to start the content pasting process"""
    return paste_content_to_editor(driver)
