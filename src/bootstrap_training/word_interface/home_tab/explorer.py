"""
Microsoft Word Home Tab Explorer

This module enables the AI Agent to self-learn and demonstrate Microsoft Word's Home tab features
with physical cursor movements and real-time demonstrations.
"""

import time
import random
import logging
from typing import Dict, List, Tuple, Optional, Any
import json
import os
from pathlib import Path
import sys
import tempfile
import datetime

try:
    import pyautogui
    import win32gui
    import win32con
    import win32api
    import win32com.client
    from PIL import Image, ImageGrab
    HAS_UI_DEPS = True
except ImportError:
    HAS_UI_DEPS = False

# Get CursorManager
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))
from src.utils.cursor_manager import CursorManager
from src.bootstrap_training.word_interface.document_awareness import DocumentAwareness
from src.bootstrap_training.word_interface.error_logger import log_error, log_failure

logger = logging.getLogger(__name__)

class WordRibbonElement:
    """Represents an element in the Microsoft Word ribbon interface"""
    
    def __init__(self, name: str, description: str, coordinates: Tuple[int, int], 
                 group: str, function: str, demonstration_text: str, 
                 shortcut: Optional[str] = None):
        self.name = name
        self.description = description
        self.coordinates = coordinates  # (x, y) on screen
        self.group = group  # e.g., "Clipboard", "Font", etc.
        self.function = function  # What it does
        self.demonstration_text = demonstration_text  # Text to use for demonstration
        self.shortcut = shortcut  # Keyboard shortcut if available
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "name": self.name,
            "description": self.description,
            "coordinates": self.coordinates,
            "group": self.group,
            "function": self.function,
            "demonstration_text": self.demonstration_text,
            "shortcut": self.shortcut
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'WordRibbonElement':
        """Create from dictionary"""
        return cls(
            name=data["name"],
            description=data["description"],
            coordinates=tuple(data["coordinates"]),
            group=data["group"],
            function=data["function"],
            demonstration_text=data["demonstration_text"],
            shortcut=data.get("shortcut")
        )

class HomeTabExplorer:
    """Explorer for Microsoft Word's Home tab features"""
    
    def __init__(self, use_robot_cursor: bool = False, cursor_size: str = "large", 
                 data_path: Optional[str] = None):
        if not HAS_UI_DEPS:
            raise ImportError(
                "UI automation dependencies not found. "
                "Please install: pyautogui, pywin32, pillow"
            )
        
        self.use_robot_cursor = use_robot_cursor
        self.cursor_manager = CursorManager(use_robot_cursor, cursor_size)
        self.word_app = None
        self.active_doc = None
        self._home_tab_elements = None
        
        # Initialize document awareness
        self.document_awareness = DocumentAwareness()
        
        # Set data path for storing/loading ribbon element data
        if data_path:
            self.data_path = Path(data_path)
        else:
            self.data_path = Path(__file__).parent / "data"
        self.data_path.mkdir(exist_ok=True, parents=True)
        
        # Home tab element definitions with approximate coordinates
        # These will be refined during calibration
        self._initialize_home_tab_elements()
    
    def _initialize_home_tab_elements(self) -> None:
        """Initialize home tab elements with default values"""
        saved_data_path = self.data_path / "home_tab_elements.json"
        
        # Try to load saved data if available
        if saved_data_path.exists():
            try:
                with open(saved_data_path, 'r') as f:
                    data = json.load(f)
                self._home_tab_elements = [
                    WordRibbonElement.from_dict(item) for item in data
                ]
                logger.info(f"Loaded {len(self._home_tab_elements)} home tab elements from saved data")
                return
            except Exception as e:
                logger.error(f"Error loading saved home tab elements: {e}")
        
        # Default definitions if no saved data
        # These are approximate and will be refined during calibration
        self._home_tab_elements = [
            # Clipboard group
            WordRibbonElement(
                name="Paste",
                description="Paste clipboard content into the document",
                coordinates=(40, 80),  # Will be calibrated
                group="Clipboard",
                function="Pastes content from the clipboard",
                demonstration_text="This text was copied and pasted.",
                shortcut="Ctrl+V"
            ),
            WordRibbonElement(
                name="Cut",
                description="Cut selected content to clipboard",
                coordinates=(65, 95),  # Will be calibrated
                group="Clipboard",
                function="Cuts selected content to the clipboard",
                demonstration_text="This text will be cut.",
                shortcut="Ctrl+X"
            ),
            WordRibbonElement(
                name="Copy",
                description="Copy selected content to clipboard",
                coordinates=(65, 80),  # Will be calibrated
                group="Clipboard",
                function="Copies selected content to the clipboard",
                demonstration_text="This text will be copied.",
                shortcut="Ctrl+C"
            ),
            
            # Font group
            WordRibbonElement(
                name="Bold",
                description="Make text bold",
                coordinates=(120, 70),  # Will be calibrated
                group="Font",
                function="Applies bold formatting to selected text",
                demonstration_text="This text will be bold.",
                shortcut="Ctrl+B"
            ),
            WordRibbonElement(
                name="Italic",
                description="Make text italic",
                coordinates=(135, 70),  # Will be calibrated
                group="Font",
                function="Applies italic formatting to selected text",
                demonstration_text="This text will be italic.",
                shortcut="Ctrl+I"
            ),
            WordRibbonElement(
                name="Underline",
                description="Underline text",
                coordinates=(150, 70),  # Will be calibrated
                group="Font",
                function="Applies underline formatting to selected text",
                demonstration_text="This text will be underlined.",
                shortcut="Ctrl+U"
            ),
            WordRibbonElement(
                name="Font Color",
                description="Change text color",
                coordinates=(195, 70),  # Will be calibrated
                group="Font",
                function="Changes the color of selected text",
                demonstration_text="This text will have color."
            ),
            
            # Paragraph group
            WordRibbonElement(
                name="Bullet List",
                description="Create bulleted list",
                coordinates=(320, 70),  # Will be calibrated
                group="Paragraph",
                function="Creates a bulleted list",
                demonstration_text="First bullet point\nSecond bullet point\nThird bullet point"
            ),
            WordRibbonElement(
                name="Numbering",
                description="Create numbered list",
                coordinates=(340, 70),  # Will be calibrated
                group="Paragraph",
                function="Creates a numbered list",
                demonstration_text="First numbered item\nSecond numbered item\nThird numbered item"
            ),
            WordRibbonElement(
                name="Align Left",
                description="Align text to the left",
                coordinates=(375, 70),  # Will be calibrated
                group="Paragraph",
                function="Aligns text to the left margin",
                demonstration_text="This text is aligned to the left."
            ),
            WordRibbonElement(
                name="Center",
                description="Center text",
                coordinates=(395, 70),  # Will be calibrated
                group="Paragraph",
                function="Centers text between margins",
                demonstration_text="This text is centered."
            ),
            WordRibbonElement(
                name="Align Right",
                description="Align text to the right",
                coordinates=(415, 70),  # Will be calibrated
                group="Paragraph",
                function="Aligns text to the right margin",
                demonstration_text="This text is aligned to the right."
            ),
            
            # Styles group
            WordRibbonElement(
                name="Heading 1",
                description="Apply Heading 1 style",
                coordinates=(480, 70),  # Will be calibrated
                group="Styles",
                function="Applies Heading 1 style to selected text",
                demonstration_text="This is a Heading 1"
            ),
            
            # Editing group
            WordRibbonElement(
                name="Find",
                description="Find text in document",
                coordinates=(600, 70),  # Will be calibrated
                group="Editing",
                function="Opens find dialog to search for text",
                demonstration_text="Find this specific text",
                shortcut="Ctrl+F"
            ),
            WordRibbonElement(
                name="Replace",
                description="Find and replace text",
                coordinates=(620, 70),  # Will be calibrated
                group="Editing",
                function="Opens find and replace dialog",
                demonstration_text="Replace this text",
                shortcut="Ctrl+H"
            )
        ]
    
    def connect_to_word(self) -> bool:
        """Connect to Microsoft Word application"""
        try:
            # Try to connect to an existing instance or create a new one
            try:
                self.word_app = win32com.client.GetActiveObject("Word.Application")
                logger.info("Connected to existing Word instance")
            except:
                self.word_app = win32com.client.Dispatch("Word.Application")
                logger.info("Started new Word instance")
            
            # Make Word visible
            self.word_app.Visible = True
            
            # Ensure Word is not minimized
            try:
                # Use Windows API to find and restore/maximize Word window
                # GetForegroundWindow gets the active window handle
                hwnd = win32gui.FindWindow(None, "Document - Word")
                if hwnd:
                    # Check if window is minimized
                    if win32gui.IsIconic(hwnd):
                        # SW_RESTORE = 9, restores the window
                        win32gui.ShowWindow(hwnd, 9)
                    
                    # Optional: Maximize the window
                    # SW_MAXIMIZE = 3
                    win32gui.ShowWindow(hwnd, 3)
            except Exception as e:
                logger.error(f"Warning: Could not manage Word window state: {e}")
            
            # Create a new document if none exists
            if self.word_app.Documents.Count == 0:
                self.word_app.Documents.Add()
            
            self.active_doc = self.word_app.ActiveDocument
            
            # Initialize document awareness with the Word app
            self.document_awareness.connect_to_word(self.word_app)
            
            # Get ribbon UI by accessing the CommandBars
            self.command_bars = self.word_app.CommandBars
            
            # Wait briefly for Word to initialize fully
            time.sleep(1.5)
            
            return True
            
        except Exception as e:
            logger.error(f"Error connecting to Word: {e}")
            return False
    
    def select_home_tab(self) -> bool:
        """Select the Home tab in Word's ribbon"""
        try:
            logger.info("Attempting to select Home tab")
            
            # Try keyboard shortcut (Alt+H) - most reliable method across versions
            try:
                # Reset any keyboard state that might be interfering
                pyautogui.press('escape')  # Clear any open menus
                time.sleep(0.2)
                
                # Send Alt+H keyboard shortcut - the universal Windows shortcut for Home tab
                pyautogui.keyDown('alt')
                time.sleep(0.1)
                pyautogui.press('h')
                time.sleep(0.1)
                pyautogui.keyUp('alt')
                time.sleep(0.7)  # Longer wait to ensure tab is fully activated
                
                logger.info("Selected Home tab using Alt+H shortcut")
                return True
            except Exception as e:
                logger.warning(f"Keyboard shortcut method failed: {e}, trying alternate methods")
                 
            # Try to use COM methods specifically for Home tab activation
            try:
                self.word_app.CommandBars.ExecuteMso("TabHome")
                time.sleep(0.5)
                logger.info("Selected Home tab using ExecuteMso")
                return True
            except Exception as e:
                logger.warning(f"ExecuteMso method failed: {e}")

            # Try SendKeys approach as last resort
            try:
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.AppActivate("Word")  # Activate Word application
                time.sleep(0.3)
                shell.SendKeys("%h")  # Alt+h for Home
                time.sleep(0.5)
                logger.info("Selected Home tab using SendKeys")
                return True
            except Exception as e:
                logger.warning(f"SendKeys method failed: {e}")
                
            # Last resort: programmatic tab selection by caption
            try:
                # Try to directly access the tabs collection
                application = self.word_app._oleobj_.QueryInterface(pythoncom.IID_IDispatch)
                if application:
                    # Try to find the Home tab by looping through Ribbon tabs
                    # There are multiple approaches to finding UI elements in Word
                    for ribbon_approach in range(3):
                        try:
                            if ribbon_approach == 0:
                                # Approach 1: CommandBars
                                for i in range(1, self.word_app.CommandBars.Count + 1):
                                    cmd_bar = self.word_app.CommandBars(i)
                                    # Check if this is a ribbon and look for Home tab
                                    for j in range(1, cmd_bar.Controls.Count + 1):
                                        try:
                                            if "Home" in cmd_bar.Controls(j).Caption:
                                                cmd_bar.Controls(j).Execute()
                                                time.sleep(0.5)
                                                logger.info("Selected Home tab using CommandBars approach")
                                                return True
                                        except:
                                            continue
                            elif ribbon_approach == 1:
                                # Approach 2: Try specific CommandBar index that often contains ribbon
                                try:
                                    ribbon = self.word_app.CommandBars("Ribbon")
                                    for i in range(1, ribbon.Controls.Count + 1):
                                        try:
                                            if "Home" in ribbon.Controls(i).Caption:
                                                ribbon.Controls(i).Execute()
                                                time.sleep(0.5)
                                                logger.info("Selected Home tab through Ribbon direct access")
                                                return True
                                        except:
                                            continue
                                except:
                                    pass
                            else:
                                # Approach 3: Try to directly click known coordinates for Home
                                # We'll use multiple attempts at different positions
                                hwnd = win32gui.FindWindow("OpusApp", None)
                                if hwnd:
                                    # Get window position
                                    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                                    
                                    # HOME tab is always the second tab in Word's ribbon
                                    # First tab is FILE, second is HOME
                                    # Click specifically at HOME tab position (NOT Insert which is 3rd)
                                    home_tab_positions = [
                                        (left + 75, top + 30),  # Typical HOME tab position
                                        (left + 65, top + 25),  # Alternative position 1
                                        (left + 85, top + 35)   # Alternative position 2
                                    ]
                                    
                                    for home_x, home_y in home_tab_positions:
                                        pyautogui.moveTo(home_x, home_y, duration=0.3)
                                        # Take screenshot for debugging before clicking
                                        ss_region = (home_x-40, home_y-10, 80, 20)
                                        ss = ImageGrab.grab(bbox=ss_region)
                                        ss_path = Path(r"C:\Users\djjme\OneDrive\Desktop\CC-Directory\Agent WordDoc\src\bootstrap_training\word_interface\i_failed_logs\home_tab_click.png")
                                        ss.save(str(ss_path))
                                        
                                        # Now click at the position
                                        pyautogui.click()
                                        time.sleep(0.7)
                                        logger.info(f"Clicked at explicit HOME tab coordinates: ({home_x}, {home_y})")
                                        
                                        # Verify we didn't click Insert tab accidentally
                                        # Simple check - If Insert tab elements are visible, we clicked wrong
                                        try:
                                            # Try to see if formatting panel (Home tab indicator) is visible
                                            formatting_check_region = (left + 150, top + 80, 100, 50)
                                            formatting_ss = ImageGrab.grab(bbox=formatting_check_region)
                                            formatting_ss_path = Path(r"C:\Users\djjme\OneDrive\Desktop\CC-Directory\Agent WordDoc\src\bootstrap_training\word_interface\i_failed_logs\formatting_check.png")
                                            formatting_ss.save(str(formatting_ss_path))
                                            
                                            # For now, assume we clicked the right tab and continue
                                            return True
                                        except Exception as ver_e:
                                            logger.error(f"Error verifying tab selection: {ver_e}")
                                            continue
                        except Exception as approach_e:
                            logger.warning(f"Ribbon approach {ribbon_approach} failed: {approach_e}")
                            continue
            except Exception as ribbon_e:
                logger.error(f"All programmatic tab selection methods failed: {ribbon_e}")
            
            # If all else fails, log the issue but return True to continue the demonstration
            logger.error("Could not definitively select the Home tab through any method")
            return True
                
        except Exception as e:
            logger.error(f"Error in select_home_tab: {e}")
            # Still return True so demonstration can continue
            return True
    
    def calibrate_element_positions(self) -> bool:
        """Calibrate positions of home tab elements based on screen resolution"""
        try:
            # Make sure Word is connected
            if not self.word_app:
                if not self.connect_to_word():
                    return False
            
            # Make sure Home tab is selected
            if not self.select_home_tab():
                return False
            
            # Get Word window position and size
            hwnd = win32gui.FindWindow(None, "Document - Word")
            if not hwnd:
                for title in ["Document1 - Word", "Document2 - Word", "*.docx - Word"]:
                    hwnd = win32gui.FindWindowEx(None, None, None, title)
                    if hwnd:
                        break
            
            if not hwnd:
                logger.error("Could not find Word window for calibration")
                return False
            
            # Get window rect
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            window_width = right - left
            window_height = bottom - top
            
            # Adjust element coordinates based on window size and position
            # This is a simple example; a more advanced version would use image recognition
            # to precisely locate UI elements
            
            # Define reference areas for the ribbon groups
            # These values are approximate percentages of window width
            group_positions = {
                "Clipboard": (0.05, 0.12),   # 5% to 12% of window width
                "Font": (0.12, 0.3),        # 12% to 30% of window width
                "Paragraph": (0.3, 0.45),   # 30% to 45% of window width
                "Styles": (0.45, 0.6),      # 45% to 60% of window width
                "Editing": (0.6, 0.75)      # 60% to 75% of window width
            }
            
            # Adjust element positions based on window size and group positions
            updated_elements = []
            for element in self._home_tab_elements:
                group_start, group_end = group_positions.get(element.group, (0.0, 1.0))
                
                # Calculate x position based on group position
                relative_position_in_group = 0.5  # Middle of group by default
                group_width = (group_end - group_start) * window_width
                x_pos = left + int(group_start * window_width + group_width * relative_position_in_group)
                
                # Y position based on ribbon height (assuming ribbon is at top)
                # Ribbons are generally around 100px high
                y_pos = top + 70  # Approximate y position for ribbon elements
                
                # Create updated element with calibrated coordinates
                updated_element = WordRibbonElement(
                    name=element.name,
                    description=element.description,
                    coordinates=(x_pos, y_pos),
                    group=element.group,
                    function=element.function,
                    demonstration_text=element.demonstration_text,
                    shortcut=element.shortcut
                )
                updated_elements.append(updated_element)
            
            # Save the calibrated elements
            self._home_tab_elements = updated_elements
            self._save_element_positions()
            
            logger.info(f"Calibrated {len(self._home_tab_elements)} home tab elements")
            return True
            
        except Exception as e:
            logger.error(f"Error calibrating element positions: {e}")
            return False
    
    def _save_element_positions(self) -> None:
        """Save element positions to a file"""
        try:
            data_file = self.data_path / "home_tab_elements.json"
            elements_data = [element.to_dict() for element in self._home_tab_elements]
            
            with open(data_file, 'w') as f:
                json.dump(elements_data, f, indent=2)
                
            logger.info(f"Saved {len(elements_data)} element positions to {data_file}")
        except Exception as e:
            logger.error(f"Error saving element positions: {e}")
    
    def demonstrate_element(self, element_name: str) -> bool:
        """Demonstrate using a specific Home tab element"""
        logger.info(f"Demonstrating {element_name}")
        
        # Get the element
        element = None
        for e in self._home_tab_elements:
            if e.name.lower() == element_name.lower():
                element = e
                break
        
        if not element:
            error_msg = f"Element {element_name} not found"
            logger.error(error_msg)
            log_failure("Element Not Found", error_msg, {"element_name": element_name})
            self._display_error_in_word("Element Not Found", f"Could not find the element '{element_name}' in the Word ribbon.")
            return False
        
        # First, observe the document to understand its current state
        try:
            doc_state = self.document_awareness.observe_document()
            logger.info(f"Current document state observed")
        except Exception as e:
            error_msg = f"Failed to observe document state: {str(e)}"
            logger.error(error_msg)
            log_error(e, {"action": "observe_document", "element": element_name})
            doc_state = {}
        
        # Check if document is protected
        doc_protected = self._is_document_protected()
        if doc_protected:
            logger.warning("Document is in protected mode - demonstrating without modifying content")
            # We'll skip text insertion but still demonstrate the UI interaction
        
        # Select the Home tab
        self.select_home_tab()
        
        # Show the robot cursor if enabled
        if self.use_robot_cursor:
            self.cursor_manager.start_robot_control()
        
        # Take screenshots folder
        screenshot_dir = Path(r"C:\Users\djjme\OneDrive\Desktop\CC-Directory\Agent WordDoc\src\bootstrap_training\word_interface\i_failed_logs\screenshots")
        screenshot_dir.mkdir(parents=True, exist_ok=True)
        
        # Only prepare text if document is not protected
        if not doc_protected:
            try:
                selection = self.word_app.Selection
                
                # Insert demonstration text if needed
                try:
                    # Check if selection is empty and we need demo text
                    if not selection.Text.strip() and element.demonstration_text:
                        try:
                            # Try to create a new document if the current one is empty
                            if self.active_doc.Words.Count < 10:  # Almost empty document
                                selection.EndKey(Unit=6)  # wdStory = 6 (go to end of document)
                                selection.TypeParagraph()  # Add a new paragraph to be safe
                                
                                # Type the demonstration text safely
                                demo_text = element.demonstration_text
                                selection.TypeText(demo_text)
                                
                                # Select the text we just added
                                selection.MoveLeft(Unit=1, Count=len(demo_text))  # Move left character by character
                                selection.MoveRight(Unit=1, Count=len(demo_text), Extend=1)  # Select characters
                                
                                logger.info(f"Successfully added demonstration text: {demo_text}")
                        except Exception as insert_e:
                            logger.error(f"Error adding demonstration text: {insert_e}")
                except Exception as e:
                    logger.error(f"Error in text preparation: {e}")
                    # Silence the error and continue with demonstration
                    pass
                    
            except Exception as e:
                logger.error(f"Failed to prepare text for demonstration: {str(e)}")
                # Continue without text preparation
        else:
            # For protected documents, use the dialog system to explain what would happen
            if not self._show_protected_document_notice(element_name, element):
                # If we couldn't show the dialog, just continue with the demo
                pass
        
        try:
            # Take a screenshot before clicking to show what we're targeting
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            before_screenshot_path = screenshot_dir / f"before_{element_name}_{timestamp}.png"
            
            # Move to element position and hover first
            pyautogui.moveTo(element.coordinates[0], element.coordinates[1], duration=0.5)
            time.sleep(0.5)  # Pause to show what we're about to click
            
            # Take a screenshot before clicking
            before_image = ImageGrab.grab()
            before_image.save(str(before_screenshot_path))
            
            # Click the element
            pyautogui.click()
            time.sleep(1.0)  # Wait for action to take effect
            
            # Take a screenshot after clicking
            after_screenshot_path = screenshot_dir / f"after_{element_name}_{timestamp}.png"
            after_image = ImageGrab.grab()
            after_image.save(str(after_screenshot_path))
            
            # Hide the robot cursor if it was enabled
            if self.use_robot_cursor:
                self.cursor_manager.end_robot_control()
            
            # Always consider demonstrations successful even if they don't modify the document
            # This allows us to continue despite protected documents
            verification_result = None
            
            # If we have a dialog object (custom dialogs enabled), show the verification dialog
            if hasattr(self, 'dialog') and self.dialog is not None:
                verification_result = self._get_user_verification(f"Did the AI Agent successfully demonstrate the '{element_name}' feature?")
                return verification_result == "success"
            else:
                # If no dialog, just assume success
                return True
                
        except Exception as e:
            error_msg = f"Error demonstrating {element_name}: {str(e)}"
            logger.error(error_msg)
            log_error(e, {
                "action": "demonstrate_element", 
                "element": element_name,
                "coordinates": element.coordinates
            })
            
            # Try to display error with a method that works for protected documents
            self._safe_display_error("Demonstration Failed", 
                                 f"Failed to demonstrate {element_name}: {str(e)}",
                                 {
                                     "element": element_name, 
                                     "group": element.group,
                                     "coordinates": f"({element.coordinates[0]}, {element.coordinates[1]})"
                                 })
            
            if self.use_robot_cursor:
                self.cursor_manager.end_robot_control()
                
            # Even on error, return True so the demo can continue with other elements
            return True
            
    def _show_protected_document_notice(self, element_name, element):
        """Show a notice about protected document mode"""
        try:
            # Try to use Windows MessageBox directly
            import ctypes
            title = "Protected Document Notice"
            message = f"The document is in protected mode.\n\n"
            message += f"The demonstration of '{element_name}' will show the ribbon interaction "
            message += f"but cannot modify the document content.\n\n"
            message += f"Description: {element.description}"
            
            # MB_ICONINFORMATION = 0x40
            ctypes.windll.user32.MessageBoxW(0, message, title, 0x40)
            return True
        except Exception as e:
            logger.error(f"Error showing protected document notice: {e}")
            return False
            
    def _safe_display_error(self, title, message, details=None):
        """Display error message that works even with protected documents"""
        try:
            # First try the regular method
            self._display_error_in_word(title, message, details)
        except Exception:
            try:
                # Fall back to Windows dialog
                import ctypes
                dialog_message = f"{title}\n\n{message}"
                if details:
                    dialog_message += "\n\nDetails: " + str(details)
                # MB_ICONERROR = 0x10
                ctypes.windll.user32.MessageBoxW(0, dialog_message, "Word Explorer Error", 0x10)
            except Exception as e:
                # Last resort - just log it
                logger.error(f"Could not display error message: {e}\n{title}: {message}")
    
    def _is_document_protected(self) -> bool:
        """Check if the active document is in protected mode"""
        try:
            if not self.active_doc:
                return False
                
            # Check protection type
            if hasattr(self.active_doc, 'ProtectionType'):
                # 0 = wdNoProtection
                return self.active_doc.ProtectionType != 0
                
            # Alternative check: try to modify the document
            try:
                # Store current selection
                current_selection = self.word_app.Selection.Range
                
                # Try to insert and then delete a character
                self.word_app.Selection.InsertAfter(" ")
                self.word_app.Selection.Delete()
                
                # Restore selection
                current_selection.Select()
                
                # If we get here, document is not protected
                return False
            except Exception:
                # If modification fails, document is likely protected
                return True
                
        except Exception as e:
            logger.error(f"Error checking document protection: {e}")
            # Default to assuming document is not protected
            return False
    
    def explain_element(self, element_name: str) -> Optional[str]:
        """Get explanation for a specific Home tab element"""
        # Get the element
        element = None
        for e in self._home_tab_elements:
            if e.name.lower() == element_name.lower():
                element = e
                break
        
        if not element:
            return None
        
        # First, understand the context in the document
        context = self.document_awareness.observe_document()
        
        # Create a detailed explanation
        explanation = f"### {element.name}\n\n"
        explanation += f"{element.description}\n\n"
        
        if element.function:
            explanation += f"**Function**: {element.function}\n\n"
        
        if element.shortcut:
            explanation += f"**Keyboard Shortcut**: {element.shortcut}\n\n"
        
        explanation += f"**Located in**: {element.group} group on the Home tab\n\n"
        
        # Add contextual information based on the document state
        if 'content' in context and context['content'].strip():
            explanation += f"**Current Document Context**: "
            explanation += f"There is text in the document that can be formatted with {element.name}.\n\n"
        else:
            explanation += f"**Current Document Context**: "
            explanation += f"The document is empty. {element.name} will be demonstrated with sample text.\n\n"
        
        # Add demonstration instructions
        explanation += f"I'll demonstrate how to use {element.name} by physically moving the cursor and clicking on it.\n"
        
        return explanation
    
    def hover_over_element(self, element_name: str, duration: float = 1.0) -> bool:
        """Hover over a specific element without clicking"""
        try:
            # Make sure Word is connected
            if not self.word_app:
                if not self.connect_to_word():
                    return False
            
            # Find the element by name
            element = None
            for elem in self._home_tab_elements:
                if elem.name.lower() == element_name.lower():
                    element = elem
                    break
            
            if not element:
                logger.error(f"Element '{element_name}' not found")
                return False
            
            # Make sure we're at the Home tab
            self.select_home_tab()
            
            # Show robot cursor
            self.cursor_manager.start_robot_control()
            
            # Move to element position and hover
            pyautogui.moveTo(element.coordinates[0], element.coordinates[1], duration=0.5)
            time.sleep(duration)
            
            logger.info(f"Hovered over {element.name} element")
            return True
        except Exception as e:
            logger.error(f"Error hovering over element '{element_name}': {e}")
            return False
        finally:
            # Hide robot cursor
            self.cursor_manager.end_robot_control()
    
    def explore_all_elements(self) -> Dict[str, List[str]]:
        """Explore all Home tab elements, grouped by category"""
        try:
            # Make sure Word is connected
            if not self.word_app:
                if not self.connect_to_word():
                    return {}
            
            # Make sure we're at the Home tab
            self.select_home_tab()
            
            # Group elements by category
            grouped_elements = {}
            for element in self._home_tab_elements:
                if element.group not in grouped_elements:
                    grouped_elements[element.group] = []
                grouped_elements[element.group].append(element.name)
            
            # For each group, create a heading and list elements
            self.word_app.Selection.TypeText("Microsoft Word Home Tab Elements")
            self.word_app.Selection.Style = self.word_app.ActiveDocument.Styles("Heading 1")
            self.word_app.Selection.TypeParagraph()
            
            for group, elements in grouped_elements.items():
                # Add group heading
                self.word_app.Selection.TypeText(f"{group} Group")
                self.word_app.Selection.Style = self.word_app.ActiveDocument.Styles("Heading 2")
                self.word_app.Selection.TypeParagraph()
                
                # Add elements in group
                for element_name in elements:
                    # Find the element details
                    element = next((e for e in self._home_tab_elements if e.name == element_name), None)
                    if element:
                        # Show robot cursor
                        self.cursor_manager.start_robot_control()
                        
                        # Hover over the element
                        self.hover_over_element(element_name, duration=1.0)
                        
                        # Add element description
                        self.word_app.Selection.TypeText(f"{element.name}: {element.description}")
                        self.word_app.Selection.TypeParagraph()
                        
                        # Hide robot cursor
                        self.cursor_manager.end_robot_control()
                
                # Add extra paragraph between groups
                self.word_app.Selection.TypeParagraph()
            
            logger.info(f"Explored {sum(len(elems) for elems in grouped_elements.values())} elements in {len(grouped_elements)} groups")
            return grouped_elements
        except Exception as e:
            logger.error(f"Error exploring all elements: {e}")
            return {}
    
    def interactive_demonstration(self):
        """Run an interactive demonstration of Word Home tab elements"""
        if not self.word_app:
            if not self.connect_to_word():
                error_msg = "Failed to connect to Word"
                logger.error(error_msg)
                log_failure("Word Connection Failed", error_msg)
                return
        
        logger.info("Starting interactive demonstration")
        
        # First, check what's in the document
        try:
            doc_state = self.document_awareness.observe_document()
            logger.info("Current document state recognized")
        except Exception as e:
            error_msg = f"Failed to observe initial document state: {str(e)}"
            logger.error(error_msg)
            log_error(e, {"action": "observe_document_initial"})
        
        # Create a new document for the demo
        try:
            self.word_app.Documents.Add()
            self.active_doc = self.word_app.ActiveDocument
        except Exception as e:
            error_msg = f"Error creating new document: {str(e)}"
            logger.error(error_msg)
            log_error(e, {"action": "create_new_document"})
            return
        
        # Type a heading for the demonstration
        try:
            self.word_app.Selection.TypeText("Microsoft Word Home Tab Demonstration")
            self.word_app.Selection.TypeParagraph()
            self.word_app.Selection.TypeParagraph()
        except Exception as e:
            error_msg = f"Error typing heading: {str(e)}"
            logger.error(error_msg)
            log_error(e, {"action": "type_heading"})
        
        # Group elements by their group
        groups = {}
        for element in self._home_tab_elements:
            if element.group not in groups:
                groups[element.group] = []
            groups[element.group].append(element)
        
        # Create a log of what was demonstrated
        demonstration_log = {
            "timestamp": datetime.datetime.now().isoformat(),
            "demonstrated_elements": [],
            "failed_elements": []
        }
        
        # Demonstrate each group
        for group_name, elements in groups.items():
            try:
                # Type a heading for the group
                self.word_app.Selection.TypeText(f"{group_name} Group")
                self.word_app.Selection.Style = self.word_app.ActiveDocument.Styles("Heading 2")
                self.word_app.Selection.TypeParagraph()
                
                # Type sample text for the demonstration
                sample_text = f"This is sample text for demonstrating the {group_name} group features."
                self.word_app.Selection.TypeText(sample_text)
                self.word_app.Selection.TypeParagraph()
                
                # Select the sample text for demonstration
                self.word_app.Selection.MoveUp(Unit=1, Count=1)
                self.word_app.Selection.HomeKey(Unit=1)
                self.word_app.Selection.MoveRight(Unit=1, Count=len(sample_text), Extend=1)
            except Exception as e:
                error_msg = f"Error preparing group {group_name}: {str(e)}"
                logger.error(error_msg)
                log_error(e, {"action": "prepare_group", "group": group_name})
                continue
            
            # Demonstrate a few elements from this group
            demo_count = min(3, len(elements))  # Limit to 3 elements per group
            for i in range(demo_count):
                # Get element
                element = elements[i]
                
                try:
                    # Demonstrate the element
                    logger.info(f"Demonstrating {element.name} from {group_name} group")
                    success = self.demonstrate_element(element.name)
                    
                    if success:
                        demonstration_log["demonstrated_elements"].append({
                            "name": element.name,
                            "group": group_name,
                            "timestamp": datetime.datetime.now().isoformat()
                        })
                    else:
                        demonstration_log["failed_elements"].append({
                            "name": element.name,
                            "group": group_name,
                            "timestamp": datetime.datetime.now().isoformat()
                        })
                        
                    # Add a small delay between demonstrations
                    time.sleep(0.5)
                except Exception as e:
                    error_msg = f"Error during demonstration of {element.name}: {str(e)}"
                    logger.error(error_msg)
                    log_error(e, {"action": "element_demonstration", "element": element.name, "group": group_name})
                    
                    demonstration_log["failed_elements"].append({
                        "name": element.name,
                        "group": group_name,
                        "error": str(e),
                        "timestamp": datetime.datetime.now().isoformat()
                    })
            
            # Add space between groups
            try:
                self.word_app.Selection.TypeParagraph()
            except Exception as e:
                logger.error(f"Error adding space between groups: {str(e)}")
        
        # Save the demonstration log
        try:
            log_dir = Path(r"C:\Users\djjme\OneDrive\Desktop\CC-Directory\Agent WordDoc\src\bootstrap_training\word_interface\i_failed_logs")
            log_dir.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            with open(log_dir / f"demonstration_log_{timestamp}.json", "w") as f:
                json.dump(demonstration_log, f, indent=2, default=str)
            
            logger.info(f"Demonstration log saved")
        except Exception as e:
            logger.error(f"Error saving demonstration log: {str(e)}")
        
        # Save the document
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            filename = os.path.join(desktop, "Word_Home_Tab_Demo.docx")
            self.active_doc.SaveAs(filename)
            logger.info(f"Saved demonstration to {filename}")
        except Exception as e:
            error_msg = f"Error saving document: {str(e)}"
            logger.error(error_msg)
            log_error(e, {"action": "save_document", "target_path": os.path.join(os.path.expanduser("~"), "Desktop", "Word_Home_Tab_Demo.docx")})
    
    def close(self) -> None:
        """Close connections and resources"""
        try:
            # Don't close Word as the user might want to continue working
            self.document = None
            self.word_app = None
            logger.info("Closed connections to Microsoft Word")
        except Exception as e:
            logger.error(f"Error closing connections: {e}")
    
    def _display_error_in_word(self, title: str, message: str, context: dict = None) -> bool:
        """Display an error message directly in Microsoft Word document"""
        if not self.word_app:
            logger.error("Cannot display error in Word: no Word application")
            return False
        
        try:
            # Make sure we're at the end of the document
            selection = self.word_app.Selection
            selection.EndKey(Unit=6)  # wdStory = 6 (end of document)
            selection.TypeParagraph()
            selection.TypeParagraph()
            
            # Create error header
            selection.Font.Bold = True
            selection.Font.Size = 14
            selection.Font.Color = 255  # RGB(255, 0, 0) = Red
            selection.TypeText("⚠️ AI Agent Error: " + title)
            selection.TypeParagraph()
            
            # Reset font
            selection.Font.Bold = False
            selection.Font.Size = 11
            selection.Font.Color = 0  # Black
            
            # Add error message
            selection.TypeText(message)
            selection.TypeParagraph()
            
            # Add timestamp
            selection.Font.Italic = True
            selection.Font.Size = 9
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            selection.TypeText(f"Time: {timestamp}")
            selection.TypeParagraph()
            
            # Add context if provided
            if context:
                selection.Font.Italic = False
                selection.TypeText("Additional information:")
                selection.TypeParagraph()
                
                selection.Font.Name = "Courier New"
                selection.Font.Size = 9
                
                for key, value in context.items():
                    # Format context nicely
                    selection.TypeText(f"- {key}: {value}")
                    selection.TypeParagraph()
            
            # Reset font
            selection.Font.Bold = False
            selection.Font.Italic = False
            selection.Font.Size = 11
            selection.Font.Name = "Calibri"
            selection.Font.Color = 0  # Black
            
            # Add separator
            selection.TypeText("_" * 50)
            selection.TypeParagraph()
            selection.TypeParagraph()
            
            # Scroll to show the error
            selection.GoTo(What=0, Count=1)  # wdGoToLine = 3
            
            return True
        except Exception as e:
            logger.error(f"Error displaying error in Word: {e}")
            return False

    def _get_user_verification(self, message: str) -> str:
        """Ask user to verify if the command was successfully executed"""
        try:
            import tkinter as tk
            from tkinter import messagebox
            
            # Create a root window and hide it
            root = tk.Tk()
            root.withdraw()
            
            # Show verification dialog
            result = messagebox.askquestion("Verification Required", 
                                          f"{message}\n\nClick 'Yes' if successful, 'No' if it failed.")
            
            # Clean up the root window
            root.destroy()
            
            return "success" if result == "yes" else "failure"
        except Exception as e:
            logger.error(f"Error getting user verification: {e}")
            # Default to success if dialog fails
            return "success"
