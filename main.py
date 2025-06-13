import time
import logging
from notion import monitor_notion_database, TargetNotionHandler, NOTION_TOKEN, NOTION_DATABASE_ID, log
from chrome import setup_chrome, cleanup_chrome
from contentpaster import start_content_pasting
from generationlogic import verify_and_generate
from export import export_audio
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
from platformconfig import get_platform

# Set console title based on platform
if get_platform() == 'Windows':
    import ctypes
    ctypes.windll.kernel32.SetConsoleTitleW("PlayAI - PlayAI API")
elif get_platform() == 'macOS':
    import subprocess
    subprocess.run(['echo', '-n', '\033]0;PlayAI - PlayAI API\007'])

def main():
    try:
        # Clear screen and print header
        print("\033[H\033[J", end="")
        log("✨ VOICEOVER AUTOMATION", "header")
        
        # Initialize Notion handler
        notion_handler = TargetNotionHandler(NOTION_TOKEN, NOTION_DATABASE_ID)
        driver = None
        
        while True:
            try:
                # 1. Check Notion for new content
                log("Checking Notion for new content...", "wait")
                records = notion_handler.get_records_for_voiceover()
                
                if records:
                    log(f"Found {len(records)} records to process", "success")
                    
                    # 2. Setup Chrome if not already running
                    if not driver:
                        log("Setting up Chrome...", "wait")
                        driver = setup_chrome()
                        if not driver:
                            log("Failed to initialize Chrome", "error")
                            time.sleep(30)
                            continue
                    
                    # 3. For each record, process the content
                    for record in records:
                        try:
                            # Get title for logging
                            title = record['properties'].get('New Title', {}).get('title', [{}])[0].get('text', {}).get('content', 'Untitled')
                            log(f"Processing: {title}", "info")
                            
                            # 3. Paste content
                            log("Pasting content...", "wait")
                            if start_content_pasting(driver):
                                log("Content pasted successfully", "success")
                                
                                # 4. Verify and generate
                                log("Verifying content and generating...", "wait")
                                if verify_and_generate(driver): 
                                    log("Generation started successfully", "success")
                                    
                                    # 5. Export audio
                                    log("Exporting audio...", "wait")
                                    if export_audio(driver):
                                        log("Audio exported successfully", "success")
                                        # Set driver to None since export closes it
                                        driver = None
                                        # Break out of record processing loop to wait for next cycle
                                        break
                                    else:
                                        log("Audio export failed", "error")
                                        # Set driver to None since export closes it
                                        driver = None
                                        # Break out of record processing loop to wait for next cycle
                                        break
                                else:
                                    log("Generation failed", "error")
                            else:
                                log("Content pasting failed", "error")
                                
                        except Exception as e:
                            log(f"Error processing record: {str(e)}", "error")
                            continue
                            
                else:
                    log("No new content to process", "info")
                
                # Wait before next check
                next_check = time.strftime("%I:%M:%S %p", time.localtime(time.time() + 30))
                log(f"Next check at {next_check}", "wait")
                time.sleep(30)
                
            except Exception as e:
                log(f"Error in main loop: {str(e)}", "error")
                if driver:
                    cleanup_chrome(driver)
                    driver = None
                time.sleep(30)
                
    except KeyboardInterrupt:
        log("Stopping automation...", "info")
        if driver:
            cleanup_chrome(driver)
        log("Goodbye! ✌️", "header")
        
    except Exception as e:
        log(f"Fatal error: {str(e)}", "error")
        if driver:
            cleanup_chrome(driver)

if __name__ == "__main__":
    main()
