import pyautogui
import keyboard
import time
import psutil
import subprocess

# Pre-defined text to type (no API calls needed)
sample_content = [
    {"section": "Abstract", "content": "This is a test abstract for the Earned Value Management paper. This test bypasses the OpenAI API calls to test typing functionality only."},
    {"section": "Introduction", "content": "This is a test introduction for the Earned Value Management paper. The purpose of this test is to verify that the typing functionality works correctly."}
]

def is_word_running():
    """Check if Microsoft Word is running"""
    try:
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == 'WINWORD.EXE':
                return True
        return False
    except:
        return False

def setup_word():
    # Get the path to Microsoft Word executable
    word_path = r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
    
    # Check if Word is already running
    if not is_word_running():
        # Launch Microsoft Word
        subprocess.Popen([word_path])
        print("Launching Microsoft Word...")
        time.sleep(5)  # Wait for Word to open

    # Handle the New document screen
    # Look for and click on the Blank document option
    try:
        # Wait a moment for the new document screen to appear
        time.sleep(2)
        
        # Click on the Blank document
        # These coordinates are an estimate and may need adjustment
        blank_doc_x, blank_doc_y = 287, 220  # Adjust these coordinates if needed
        pyautogui.click(blank_doc_x, blank_doc_y)
        time.sleep(1)
    except Exception as e:
        # Fallback method - use keyboard shortcut if clicking fails
        print(f"Falling back to keyboard shortcut: {str(e)}")
        pyautogui.hotkey('ctrl', 'n')
        time.sleep(1)

def type_text(text, typing_mode='fast'):
    """Type text into Word with selected typing speed"""
    if typing_mode == 'fast':
        # Fast typing - 0.01s per character
        for char in text:
            keyboard.write(char)
            time.sleep(0.01)
    else:  # instant mode
        keyboard.write(text)

def press_enter(times=1):
    """Press Enter key multiple times"""
    for _ in range(times):
        keyboard.press_and_release('enter')
        time.sleep(0.1)

def type_heading(text):
    """Type and format a heading"""
    type_text(text)
    press_enter()
    # Select the line we just typed
    keyboard.press('shift+up')
    time.sleep(0.1)
    keyboard.release('shift+up')
    # Apply heading style
    keyboard.press('ctrl+alt+1')  # Heading 1 style
    time.sleep(0.1)
    keyboard.release('ctrl+alt+1')
    press_enter()

def test_typing():
    print("=== Typing Test (No API Calls) ===")
    print("Setting up Microsoft Word...")
    setup_word()
    
    print("\nStarting to type test content...")
    # Type title
    type_text("Earned Value Management: Test Document")
    press_enter(3)
    
    # Type each section
    for section in sample_content:
        print(f"Typing section: {section['section']}")
        type_heading(section["section"])
        type_text(section["content"])
        press_enter(2)
        
    print("\nTyping test completed!")

if __name__ == "__main__":
    test_typing()
