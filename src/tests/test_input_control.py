import pyautogui
import keyboard
import time

def test_pyautogui():
    print("Testing PyAutoGUI...")
    print(f"Current mouse position: {pyautogui.position()}")
    print("Moving mouse in 3 seconds...")
    time.sleep(3)
    # Move mouse to center of screen
    screen_width, screen_height = pyautogui.size()
    pyautogui.moveTo(screen_width // 2, screen_height // 2, duration=1)
    print(f"New mouse position: {pyautogui.position()}")
    print("PyAutoGUI test completed")

def test_keyboard():
    print("Testing Keyboard...")
    print("Press 'q' key to exit keyboard test")
    print("Typing 'test' in 3 seconds...")
    time.sleep(3)
    # Type 'test'
    keyboard.write('test')
    print("Keyboard test completed")

if __name__ == "__main__":
    print("=== Input Control Test Script ===")
    print("This script will test if pyautogui and keyboard can control your input devices.")
    print("\nStep 1: PyAutoGUI mouse control test")
    test_pyautogui()
    
    print("\nStep 2: Keyboard control test")
    print("Open a text editor or document before continuing")
    input("Press Enter when ready...")
    test_keyboard()
    
    print("\nTest completed. If no errors appeared and the tests worked as expected,")
    print("then both packages have the necessary permissions to control your input devices.")
