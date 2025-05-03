import keyboard
import time

print("=== Keyboard Control Test Script ===")
print("This script will test if the keyboard module can type text")
print("Please open a text editor (like Notepad) and place your cursor there")
print("The script will wait 5 seconds and then type some text")

# Wait for user to position cursor
for i in range(5, 0, -1):
    print(f"Typing will begin in {i} seconds...")
    time.sleep(1)

print("Now typing...")
try:
    # Try to type some text
    keyboard.write("This is a test of keyboard control.")
    print("\nKeyboard typing test completed successfully!")
except Exception as e:
    print(f"\nError during keyboard test: {str(e)}")
