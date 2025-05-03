import pyautogui
import time
import psutil
import subprocess

def is_word_running():
    """Check if Microsoft Word is running"""
    try:
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == 'WINWORD.EXE':
                return True
        return False
    except Exception as e:
        print(f"Error checking if Word is running: {str(e)}")
        return False

def launch_word():
    """Launch Microsoft Word"""
    try:
        word_path = r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
        subprocess.Popen([word_path])
        print("Launching Microsoft Word...")
        time.sleep(5)  # Wait for Word to open
        return True
    except Exception as e:
        print(f"Error launching Word: {str(e)}")
        return False

def test_word_interaction():
    print("=== Word Interaction Test ===")
    
    # Check if Word is running
    if is_word_running():
        print("Microsoft Word is already running")
    else:
        print("Microsoft Word is not running")
        success = launch_word()
        if not success:
            print("Failed to launch Microsoft Word")
            return
    
    # Try to interact with Word
    print("\nAttempting to type into Word...")
    print("Will try to create a new document if needed")
    
    try:
        # Try to create a new document with Ctrl+N
        print("Pressing Ctrl+N to create a new document...")
        pyautogui.hotkey('ctrl', 'n')
        time.sleep(2)
        
        # Type some text
        print("Typing test text...")
        pyautogui.write("This is a test of typing into Microsoft Word.")
        print("\nWord interaction test completed!")
    except Exception as e:
        print(f"Error during Word interaction: {str(e)}")

if __name__ == "__main__":
    test_word_interaction()
