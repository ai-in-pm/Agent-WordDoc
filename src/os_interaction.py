"""
OS Interaction for AI Agent

This module provides real-time interaction capabilities with the operating system,
including keyboard control, mouse control, application launching, and command-line interaction.
"""

import os
import time
import random
import string
import subprocess
import psutil
from typing import Dict, List, Any, Optional, Union, Tuple
import pyautogui
import keyboard
from enum import Enum

# Import cursor managers if on Windows
if os.name == 'nt':
    try:
        from src.cursor_manager import CursorManager
        ADVANCED_CURSOR = True
    except ImportError:
        from src.simple_cursor import SimpleCursorManager
        ADVANCED_CURSOR = False
        print("Using simplified cursor manager (no custom icon)")

    # Import AutoIt integration
    try:
        from src.autoit_integration import AutoItIntegration, AUTOIT_INSTALLED
        AUTOIT_AVAILABLE = True
    except ImportError:
        AUTOIT_AVAILABLE = False
        print("AutoIt integration not available")

    # Import control UI
    try:
        from src.utils.control_ui import AgentControlUI
        CONTROL_UI_AVAILABLE = True
    except ImportError:
        CONTROL_UI_AVAILABLE = False
        print("Control UI not available")

# Configure PyAutoGUI settings
pyautogui.FAILSAFE = True  # Move mouse to upper-left corner to abort
pyautogui.PAUSE = 0.1  # Add small pause between PyAutoGUI commands


class TypingMode(Enum):
    """Typing modes for keyboard interaction"""
    FAST = "fast"
    REALISTIC = "realistic"
    SLOW = "slow"


class OSInteraction:
    """Provides real-time interaction with the operating system"""

    def __init__(self, typing_mode: str = 'realistic', verbose: bool = False, show_robot_cursor: bool = True, use_autoit: bool = True, show_control_ui: bool = True):
        """
        Initialize the OS interaction module

        Args:
            typing_mode: Typing behavior mode ('fast', 'realistic', 'slow')
            verbose: Whether to print debug information
            show_robot_cursor: Whether to show a robot cursor during AI control
            use_autoit: Whether to use AutoIt for advanced Windows automation
            show_control_ui: Whether to show the control UI with Pause/Stop/Continue buttons
        """
        self.verbose = verbose
        self.typing_mode = typing_mode
        self.show_robot_cursor = show_robot_cursor
        self.use_autoit = use_autoit
        self.show_control_ui = show_control_ui
        self.paused = False
        self.stopped = False

        # Initialize typing behavior
        self._initialize_typing_behavior()

        # Track active applications
        self.active_apps = {}

        # Initialize cursor manager on Windows
        self.cursor_manager = None
        if os.name == 'nt' and show_robot_cursor:
            try:
                if ADVANCED_CURSOR:
                    # Use the advanced cursor manager with custom icon
                    self.cursor_manager = CursorManager(use_robot_cursor=True, size="standard", verbose=verbose)
                    if self.verbose:
                        print("[OS] Advanced cursor manager initialized")
                else:
                    # Use the simplified cursor manager
                    self.cursor_manager = SimpleCursorManager(verbose=verbose)
                    if self.verbose:
                        print("[OS] Simple cursor manager initialized")
            except Exception as e:
                if self.verbose:
                    print(f"[ERROR] Failed to initialize cursor manager: {str(e)}")

        # Initialize AutoIt integration on Windows
        self.autoit = None
        if os.name == 'nt' and use_autoit and 'AUTOIT_AVAILABLE' in globals() and AUTOIT_AVAILABLE:
            try:
                self.autoit = AutoItIntegration(verbose=verbose)
                if self.verbose:
                    print("[OS] AutoIt integration initialized")
            except Exception as e:
                if self.verbose:
                    print(f"[ERROR] Failed to initialize AutoIt integration: {str(e)}")
                    
        # Initialize Control UI
        self.control_ui = None
        if os.name == 'nt' and show_control_ui and 'CONTROL_UI_AVAILABLE' in globals() and CONTROL_UI_AVAILABLE:
            try:
                self.control_ui = AgentControlUI(
                    on_pause=self._handle_pause,
                    on_stop=self._handle_stop,
                    on_continue=self._handle_continue
                )
                if self.verbose:
                    print("[OS] Control UI initialized")
                # Show the control UI
                self.control_ui.show()
            except Exception as e:
                if self.verbose:
                    print(f"[ERROR] Failed to initialize control UI: {str(e)}")

        if self.verbose:
            print(f"[OS] Initialized OS interaction with typing mode: {typing_mode}")

    def _initialize_typing_behavior(self):
        """Initialize typing behavior based on mode"""
        # Base typing speeds (seconds per character)
        typing_speeds = {
            'fast': 0.01,
            'realistic': 0.05,
            'slow': 0.1
        }

        # Typing variations (random delay range to add)
        typing_variations = {
            'fast': (0.005, 0.02),
            'realistic': (0.01, 0.05),
            'slow': (0.05, 0.15)
        }

        # Error rates for different typing modes
        error_rates = {
            'fast': 0.01,  # 1% chance of error
            'realistic': 0.03,  # 3% chance of error
            'slow': 0.005  # 0.5% chance of error
        }

        # Set current typing behavior
        self.current_typing_speed = typing_speeds.get(self.typing_mode, typing_speeds['realistic'])
        self.current_typing_variation = typing_variations.get(self.typing_mode, typing_variations['realistic'])
        self.current_error_rate = error_rates.get(self.typing_mode, error_rates['realistic'])

        # Character-specific behavior
        self.character_behaviors = {
            'letters': {
                'speed_factor': 1.0,
                'error_rate': self.current_error_rate,
                'pause_after': 0.05  # 5% chance of pause after letter
            },
            'digits': {
                'speed_factor': 1.2,  # Slightly slower for digits
                'error_rate': self.current_error_rate * 1.5,  # Higher error rate for digits
                'pause_after': 0.1  # 10% chance of pause after digit
            },
            'punctuation': {
                'speed_factor': 1.3,  # Slower for punctuation
                'error_rate': self.current_error_rate * 0.5,  # Lower error rate for punctuation
                'pause_after': 0.3  # 30% chance of pause after punctuation
            },
            'spaces': {
                'speed_factor': 0.8,  # Faster for spaces
                'error_rate': 0.0,  # No errors for spaces
                'pause_after': 0.0  # No pauses after spaces
            },
            'special': {
                'speed_factor': 1.5,  # Slower for special characters
                'error_rate': self.current_error_rate * 2.0,  # Higher error rate for special chars
                'pause_after': 0.2  # 20% chance of pause after special char
            }
        }

    def activate_robot_cursor(self, blinking=True) -> bool:
        """
        Activate the robot cursor to indicate AI control

        Args:
            blinking: Whether the cursor should blink

        Returns:
            True if successful, False otherwise
        """
        if not self.show_robot_cursor or not self.cursor_manager:
            return False

        return self.cursor_manager.activate_robot_cursor(blinking=blinking)

    def deactivate_robot_cursor(self) -> bool:
        """
        Deactivate the robot cursor and restore the original cursor

        Returns:
            True if successful, False otherwise
        """
        if not self.show_robot_cursor or not self.cursor_manager:
            return False

        return self.cursor_manager.deactivate_robot_cursor()

    def cleanup(self) -> bool:
        """
        Clean up resources

        Returns:
            True if successful, False otherwise
        """
        success = True

        # Clean up cursor manager
        if self.cursor_manager:
            success = self.cursor_manager.cleanup()

        return success

    # AutoIt integration methods

    def word_type_text_autoit(self, text: str) -> bool:
        """
        Type text in Microsoft Word using AutoIt

        Args:
            text: Text to type

        Returns:
            True if successful, False otherwise
        """
        if not self.autoit:
            if self.verbose:
                print("[ERROR] AutoIt integration not available")
            return False

        # Activate robot cursor to indicate AI control
        if self.show_robot_cursor and self.cursor_manager:
            self.activate_robot_cursor()

        try:
            result = self.autoit.word_type_text(text, self.typing_mode)

            # Deactivate robot cursor when done
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            return result
        except Exception as e:
            # Ensure cursor is restored even on error
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            if self.verbose:
                print(f"[ERROR] Failed to type text with AutoIt: {str(e)}")
            return False

    def word_format_text_autoit(self, format_type: str, start_select: bool = True) -> bool:
        """
        Format selected text in Microsoft Word using AutoIt

        Args:
            format_type: Type of formatting ('bold', 'italic', 'underline', 'heading1', etc.)
            start_select: Whether to start by selecting text

        Returns:
            True if successful, False otherwise
        """
        if not self.autoit:
            if self.verbose:
                print("[ERROR] AutoIt integration not available")
            return False

        # Activate robot cursor to indicate AI control
        if self.show_robot_cursor and self.cursor_manager:
            self.activate_robot_cursor()

        try:
            result = self.autoit.word_format_text(format_type, start_select)

            # Deactivate robot cursor when done
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            return result
        except Exception as e:
            # Ensure cursor is restored even on error
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            if self.verbose:
                print(f"[ERROR] Failed to format text with AutoIt: {str(e)}")
            return False

    def word_insert_table_autoit(self, rows: int, columns: int) -> bool:
        """
        Insert a table in Microsoft Word using AutoIt

        Args:
            rows: Number of rows
            columns: Number of columns

        Returns:
            True if successful, False otherwise
        """
        if not self.autoit:
            if self.verbose:
                print("[ERROR] AutoIt integration not available")
            return False

        # Activate robot cursor to indicate AI control
        if self.show_robot_cursor and self.cursor_manager:
            self.activate_robot_cursor()

        try:
            result = self.autoit.word_insert_table(rows, columns)

            # Deactivate robot cursor when done
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            return result
        except Exception as e:
            # Ensure cursor is restored even on error
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            if self.verbose:
                print(f"[ERROR] Failed to insert table with AutoIt: {str(e)}")
            return False

    def word_navigate_autoit(self, location: str) -> bool:
        """
        Navigate to a specific location in Microsoft Word using AutoIt

        Args:
            location: Where to navigate ('start', 'end', 'next_page', 'previous_page')

        Returns:
            True if successful, False otherwise
        """
        if not self.autoit:
            if self.verbose:
                print("[ERROR] AutoIt integration not available")
            return False

        # Activate robot cursor to indicate AI control
        if self.show_robot_cursor and self.cursor_manager:
            self.activate_robot_cursor()

        try:
            result = self.autoit.word_navigate(location)

            # Deactivate robot cursor when done
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            return result
        except Exception as e:
            # Ensure cursor is restored even on error
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            if self.verbose:
                print(f"[ERROR] Failed to navigate with AutoIt: {str(e)}")
            return False

    def word_save_document_autoit(self, file_path: Optional[str] = None) -> bool:
        """
        Save the current Microsoft Word document using AutoIt

        Args:
            file_path: Path to save the document (if None, uses Save, otherwise uses Save As)

        Returns:
            True if successful, False otherwise
        """
        if not self.autoit:
            if self.verbose:
                print("[ERROR] AutoIt integration not available")
            return False

        # Activate robot cursor to indicate AI control
        if self.show_robot_cursor and self.cursor_manager:
            self.activate_robot_cursor()

        try:
            result = self.autoit.word_save_document(file_path)

            # Deactivate robot cursor when done
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            return result
        except Exception as e:
            # Ensure cursor is restored even on error
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            if self.verbose:
                print(f"[ERROR] Failed to save document with AutoIt: {str(e)}")
            return False

    def type_text(self, text: str, focus_window: bool = True, use_autoit: bool = None) -> bool:
        """
        Type text with realistic behavior

        Args:
            text: Text to type
            focus_window: Whether to ensure the window is focused first
            use_autoit: Whether to use AutoIt (if None, uses the instance setting)

        Returns:
            True if successful, False otherwise
        """
        # Determine whether to use AutoIt
        if use_autoit is None:
            use_autoit = self.use_autoit

        # Try to use AutoIt if available and requested
        if use_autoit and self.autoit and os.name == 'nt':
            # Check if we're typing in Word
            try:
                import win32gui
                foreground_window = win32gui.GetForegroundWindow()
                window_title = win32gui.GetWindowText(foreground_window)

                if "Word" in window_title or "Document" in window_title:
                    if self.verbose:
                        print(f"[OS] Using AutoIt to type in Word: {text[:30]}..." if len(text) > 30 else f"[OS] Using AutoIt to type in Word: {text}")
                    return self.word_type_text_autoit(text)
            except:
                # Fall back to standard typing if we can't check the window
                pass

        # Standard typing behavior
        try:
            # Activate robot cursor to indicate AI control
            if self.show_robot_cursor and self.cursor_manager:
                self.activate_robot_cursor()

            if self.verbose:
                print(f"[OS] Typing: {text[:30]}..." if len(text) > 30 else f"[OS] Typing: {text}")

            # Split text into words for better readability
            words = text.split()

            for word in words:
                # Add small pause between words
                time.sleep(random.uniform(0.1, 0.3))

                for char in word:
                    # Determine character type
                    if char in string.ascii_letters:
                        char_type = 'letters'
                    elif char in string.digits:
                        char_type = 'digits'
                    elif char in string.punctuation:
                        char_type = 'punctuation'
                    elif char in string.whitespace:
                        char_type = 'spaces'
                    else:
                        char_type = 'special'

                    # Get character behavior
                    behavior = self.character_behaviors[char_type]

                    # Calculate typing speed with variation
                    base_speed = self.current_typing_speed * behavior['speed_factor']
                    delay = base_speed + random.uniform(
                        *self.current_typing_variation
                    )

                    # Simulate typing errors
                    if random.random() < behavior['error_rate']:
                        # Type a random error character
                        error_char = random.choice(string.ascii_letters)
                        keyboard.write(error_char)
                        time.sleep(0.05)
                        keyboard.press_and_release('backspace')
                        time.sleep(0.02)

                    # Type the actual character
                    keyboard.write(char)
                    time.sleep(delay)

                    # Add small pause after certain characters
                    if random.random() < behavior['pause_after']:
                        time.sleep(random.uniform(0.01, 0.03))

                # Add space after word
                keyboard.write(' ')
                time.sleep(random.uniform(0.05, 0.1))

            # Deactivate robot cursor when done
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            return True
        except Exception as e:
            # Ensure cursor is restored even on error
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            if self.verbose:
                print(f"[ERROR] Failed to type text: {str(e)}")
            return False

    def type_command(self, command: str, execute: bool = True) -> bool:
        """
        Type a command with realistic behavior and optionally execute it

        Args:
            command: Command to type
            execute: Whether to press Enter to execute the command

        Returns:
            True if successful, False otherwise
        """
        try:
            # Type the command
            success = self.type_text(command)

            if not success:
                return False

            # Execute the command if requested
            if execute:
                time.sleep(random.uniform(0.2, 0.5))  # Pause before pressing Enter
                keyboard.press_and_release('enter')

                if self.verbose:
                    print(f"[OS] Executed command: {command}")

            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to type command: {str(e)}")
            return False

    def press_key(self, key: str, times: int = 1) -> bool:
        """
        Press a key multiple times

        Args:
            key: Key to press
            times: Number of times to press the key

        Returns:
            True if successful, False otherwise
        """
        try:
            for _ in range(times):
                keyboard.press_and_release(key)
                time.sleep(random.uniform(0.1, 0.2))

            if self.verbose:
                print(f"[OS] Pressed key '{key}' {times} times")

            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to press key '{key}': {str(e)}")
            return False

    def press_hotkey(self, *keys) -> bool:
        """
        Press a hotkey combination

        Args:
            *keys: Keys to press simultaneously

        Returns:
            True if successful, False otherwise
        """
        try:
            keyboard.press_and_release('+'.join(keys))

            if self.verbose:
                print(f"[OS] Pressed hotkey: {'+'.join(keys)}")

            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to press hotkey: {str(e)}")
            return False

    def move_mouse(self, x: int, y: int, duration: float = 0.5) -> bool:
        """
        Move the mouse to a specific position

        Args:
            x: X coordinate
            y: Y coordinate
            duration: Duration of the movement in seconds

        Returns:
            True if successful, False otherwise
        """
        try:
            # Activate robot cursor to indicate AI control
            if self.show_robot_cursor and self.cursor_manager:
                self.activate_robot_cursor()

            # Add some human-like randomness to the movement
            pyautogui.moveTo(
                x + random.randint(-5, 5),
                y + random.randint(-5, 5),
                duration=duration
            )

            if self.verbose:
                print(f"[OS] Moved mouse to ({x}, {y})")

            return True
        except Exception as e:
            # Ensure cursor is restored even on error
            if self.show_robot_cursor and self.cursor_manager:
                self.deactivate_robot_cursor()

            if self.verbose:
                print(f"[ERROR] Failed to move mouse: {str(e)}")
            return False

    def click_mouse(self, x: Optional[int] = None, y: Optional[int] = None, button: str = 'left') -> bool:
        """
        Click the mouse at the current position or a specific position

        Args:
            x: X coordinate (optional)
            y: Y coordinate (optional)
            button: Mouse button to click ('left', 'right', 'middle')

        Returns:
            True if successful, False otherwise
        """
        try:
            if x is not None and y is not None:
                # Move to position first
                self.move_mouse(x, y)

            # Add a small delay before clicking
            time.sleep(random.uniform(0.1, 0.2))

            # Click with the specified button
            pyautogui.click(button=button)

            if self.verbose:
                position = pyautogui.position()
                print(f"[OS] Clicked {button} button at ({position.x}, {position.y})")

            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to click mouse: {str(e)}")
            return False

    def double_click(self, x: Optional[int] = None, y: Optional[int] = None) -> bool:
        """
        Double-click the mouse at the current position or a specific position

        Args:
            x: X coordinate (optional)
            y: Y coordinate (optional)

        Returns:
            True if successful, False otherwise
        """
        try:
            if x is not None and y is not None:
                # Move to position first
                self.move_mouse(x, y)

            # Add a small delay before clicking
            time.sleep(random.uniform(0.1, 0.2))

            # Double-click
            pyautogui.doubleClick()

            if self.verbose:
                position = pyautogui.position()
                print(f"[OS] Double-clicked at ({position.x}, {position.y})")

            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to double-click: {str(e)}")
            return False

    def scroll(self, clicks: int) -> bool:
        """
        Scroll the mouse wheel

        Args:
            clicks: Number of clicks to scroll (positive for up, negative for down)

        Returns:
            True if successful, False otherwise
        """
        try:
            pyautogui.scroll(clicks)

            if self.verbose:
                direction = "up" if clicks > 0 else "down"
                print(f"[OS] Scrolled {direction} {abs(clicks)} clicks")

            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to scroll: {str(e)}")
            return False

    def launch_application(self, app_path: str) -> Optional[int]:
        """
        Launch an application

        Args:
            app_path: Path to the application executable

        Returns:
            Process ID if successful, None otherwise
        """
        try:
            # Launch the application
            process = subprocess.Popen([app_path])
            pid = process.pid

            # Store in active applications
            self.active_apps[pid] = {
                'path': app_path,
                'process': process,
                'launched_at': time.time()
            }

            if self.verbose:
                print(f"[OS] Launched application: {app_path} (PID: {pid})")

            # Give the application time to start
            time.sleep(2)

            return pid
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to launch application: {str(e)}")
            return None

    def is_application_running(self, app_name: str) -> bool:
        """
        Check if an application is running

        Args:
            app_name: Name of the application process

        Returns:
            True if running, False otherwise
        """
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].lower() == app_name.lower():
                    return True
            return False
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to check if application is running: {str(e)}")
            return False

    def focus_window(self, window_title: str) -> bool:
        """
        Focus a window by title

        Args:
            window_title: Title of the window to focus

        Returns:
            True if successful, False otherwise
        """
        try:
            # This is platform-specific and may need adjustment
            if os.name == 'nt':  # Windows
                import win32gui

                def callback(hwnd, windows):
                    if win32gui.IsWindowVisible(hwnd) and window_title.lower() in win32gui.GetWindowText(hwnd).lower():
                        windows.append(hwnd)
                    return True

                windows = []
                win32gui.EnumWindows(callback, windows)

                if windows:
                    win32gui.SetForegroundWindow(windows[0])

                    if self.verbose:
                        print(f"[OS] Focused window: {window_title}")

                    return True
                else:
                    if self.verbose:
                        print(f"[OS] Window not found: {window_title}")
                    return False
            else:
                if self.verbose:
                    print("[OS] Window focusing not implemented for this platform")
                return False
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to focus window: {str(e)}")
            return False

    def execute_command_line(self, command: str, real_typing: bool = True) -> Optional[str]:
        """
        Execute a command in the command line with realistic typing

        Args:
            command: Command to execute
            real_typing: Whether to type the command realistically or just execute it

        Returns:
            Command output if successful, None otherwise
        """
        try:
            if real_typing:
                # Launch command prompt or terminal
                if os.name == 'nt':  # Windows
                    self.launch_application('cmd.exe')
                    time.sleep(1)
                else:  # Unix-like
                    self.launch_application('xterm')
                    time.sleep(1)

                # Type and execute the command
                self.type_command(command)

                # Wait for command to complete
                time.sleep(2)

                # Return a placeholder since we can't capture output when typing
                return "[Command executed through realistic typing]"
            else:
                # Execute command directly and capture output
                result = subprocess.run(
                    command,
                    shell=True,
                    capture_output=True,
                    text=True
                )

                if self.verbose:
                    print(f"[OS] Executed command: {command}")
                    print(f"[OS] Command output: {result.stdout}")

                return result.stdout
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to execute command: {str(e)}")
            return None

    def run_python_module(self, module_path: str, args: str = "", real_typing: bool = True) -> bool:
        """
        Run a Python module with realistic typing

        Args:
            module_path: Path to the Python module
            args: Command-line arguments
            real_typing: Whether to type the command realistically

        Returns:
            True if successful, False otherwise
        """
        try:
            command = f"python -m {module_path}"
            if args:
                command += f" {args}"

            if real_typing:
                # Launch command prompt or terminal
                if os.name == 'nt':  # Windows
                    self.launch_application('cmd.exe')
                    time.sleep(1)
                else:  # Unix-like
                    self.launch_application('xterm')
                    time.sleep(1)

                # Type and execute the command
                self.type_command(command)

                if self.verbose:
                    print(f"[OS] Running Python module: {module_path} with args: {args}")

                return True
            else:
                # Execute directly
                subprocess.Popen(command, shell=True)

                if self.verbose:
                    print(f"[OS] Running Python module: {module_path} with args: {args}")

                return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to run Python module: {str(e)}")
            return False

    def take_screenshot(self, save_path: Optional[str] = None) -> Optional[str]:
        """
        Take a screenshot

        Args:
            save_path: Path to save the screenshot (optional)

        Returns:
            Path to the screenshot if saved, None otherwise
        """
        try:
            if save_path:
                # Save screenshot to file
                pyautogui.screenshot(save_path)

                if self.verbose:
                    print(f"[OS] Saved screenshot to: {save_path}")

                return save_path
            else:
                # Generate a filename based on timestamp
                timestamp = int(time.time())
                filename = f"screenshot_{timestamp}.png"

                # Save to current directory
                pyautogui.screenshot(filename)

                if self.verbose:
                    print(f"[OS] Saved screenshot to: {filename}")

                return filename
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to take screenshot: {str(e)}")
            return None

    def find_image_on_screen(self, image_path: str, confidence: float = 0.9) -> Optional[Tuple[int, int]]:
        """
        Find an image on the screen

        Args:
            image_path: Path to the image file
            confidence: Confidence level for the match (0.0 to 1.0)

        Returns:
            (x, y) coordinates of the center of the image if found, None otherwise
        """
        try:
            # Find the image on screen
            location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)

            if location:
                if self.verbose:
                    print(f"[OS] Found image at: ({location.x}, {location.y})")

                return (location.x, location.y)
            else:
                if self.verbose:
                    print(f"[OS] Image not found: {image_path}")

                return None
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to find image: {str(e)}")
            return None

    def click_image(self, image_path: str, confidence: float = 0.9) -> bool:
        """
        Find and click an image on the screen

        Args:
            image_path: Path to the image file
            confidence: Confidence level for the match (0.0 to 1.0)

        Returns:
            True if successful, False otherwise
        """
        try:
            # Find the image
            location = self.find_image_on_screen(image_path, confidence)

            if location:
                # Click the image
                self.click_mouse(location[0], location[1])
                return True
            else:
                return False
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to click image: {str(e)}")
            return False

    def _handle_pause(self):
        """Handle pause request from Control UI"""
        self.paused = True
        if self.verbose:
            print("[OS] AI Agent paused by user")
        
        # Additional pause behavior can be added here
        # For example, temporarily disable keyboard and mouse control
        
    def _handle_stop(self):
        """Handle stop request from Control UI"""
        self.stopped = True
        self.paused = False
        if self.verbose:
            print("[OS] AI Agent stopped by user")
        
        # Additional stop behavior can be added here
        # For example, end all ongoing operations and close applications
        
    def _handle_continue(self):
        """Handle continue request from Control UI"""
        self.paused = False
        if self.verbose:
            print("[OS] AI Agent continuing operation")
        
        # Additional continue behavior can be added here
        
    def _check_paused(self):
        """Check if agent is paused and wait if necessary"""
        while self.paused and not self.stopped:
            time.sleep(0.1)  # Small delay to prevent CPU hogging
        return self.stopped
    
    def type_text(self, text: str, interval: Optional[float] = None) -> bool:
        """
        Type text using the keyboard
        
        Args:
            text: Text to type
            interval: Override the typing interval
            
        Returns:
            Success status
        """
        if self._check_paused():
            return False
            
        # Show robot cursor during typing if enabled
        if self.cursor_manager:
            self.cursor_manager.start_robot_control()
            
        try:
            if interval is None:
                interval = self.get_typing_interval()
                
            # Type each character with realistic timing
            for char in text:
                if self._check_paused():
                    return False
                    
                pyautogui.write(char)
                
                # Add variable delay between keystrokes for realism
                if interval > 0:
                    actual_interval = interval
                    if self.typing_mode != TypingMode.FAST.value:
                        # Add slight randomness to timing for more realism
                        actual_interval = interval * random.uniform(0.8, 1.2)
                    time.sleep(actual_interval)
                    
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to type text: {str(e)}")
            return False
        finally:
            # Hide robot cursor after typing
            if self.cursor_manager:
                self.cursor_manager.end_robot_control()

    def cleanup(self):
        """Clean up resources"""
        # Close any active applications
        for app_name, app_info in list(self.active_apps.items()):
            if app_info['process'] and psutil.pid_exists(app_info['process'].pid):
                try:
                    app_info['process'].terminate()
                    if self.verbose:
                        print(f"[OS] Terminated {app_name}")
                except Exception:
                    pass
                    
        # Close control UI if open
        if self.control_ui:
            try:
                self.control_ui.hide()
                if self.verbose:
                    print("[OS] Closed control UI")
            except Exception:
                pass
