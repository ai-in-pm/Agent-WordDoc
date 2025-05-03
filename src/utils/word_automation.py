"""
Microsoft Word automation utilities for the Word AI Agent
"""

import os
import time
import logging
import subprocess
from pathlib import Path
import ctypes
from typing import Tuple, Optional

# Windows API constants
SW_MAXIMIZE = 3
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
VK_CONTROL = 0x11
VK_SHIFT = 0x10
VK_MENU = 0x12  # Alt key

# Import platform-specific modules
try:
    import win32com.client
    import win32gui
    import win32con
    import win32api
    import pyautogui
    HAS_AUTOMATION_DEPS = True
except ImportError:
    HAS_AUTOMATION_DEPS = False

from src.core.logging import get_logger
from src.utils.cursor_manager import CursorManager

logger = get_logger(__name__)

class WordAutomation:
    """Microsoft Word automation class"""
    
    def __init__(self, use_robot_cursor: bool = True, cursor_size: str = "standard"):
        """Initialize Word automation"""
        if not HAS_AUTOMATION_DEPS:
            raise ImportError(
                "Word automation dependencies not found. "
                "Please install pywin32, pyautogui: pip install pywin32 pyautogui"
            )
        
        self.word_app = None
        self.document = None
        self._is_connected = False
        
        # Initialize cursor manager for visual robot cursor
        self.cursor_manager = CursorManager(use_robot_cursor, cursor_size)
    
    def open_word(self) -> bool:
        """Open Microsoft Word application"""
        try:
            logger.info("Opening Microsoft Word...")
            
            # Show robot cursor to indicate AI control
            self.cursor_manager.start_robot_control()
            
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = True
            time.sleep(1)  # Wait for Word to fully initialize
            
            # Maximize the Word window
            hwnd = win32gui.FindWindow(None, "Document - Word")
            if hwnd:
                win32gui.ShowWindow(hwnd, SW_MAXIMIZE)
            
            self._is_connected = True
            logger.info("Microsoft Word opened successfully")
            return True
        except Exception as e:
            logger.error(f"Failed to open Microsoft Word: {str(e)}")
            # Hide robot cursor if there was an error
            self.cursor_manager.end_robot_control()
            return False
    
    def open_document(self, path: Optional[str] = None) -> bool:
        """Open a document in Word or create a new one"""
        try:
            if not self._is_connected and not self.open_word():
                return False
                
            if path and os.path.exists(path):
                logger.info(f"Opening document: {path}")
                self.document = self.word_app.Documents.Open(path)
            else:
                logger.info("Creating new document")
                self.document = self.word_app.Documents.Add()
            
            return True
        except Exception as e:
            logger.error(f"Failed to open/create document: {str(e)}")
            return False
    
    def type_text(self, text: str, delay: float = 0.1) -> bool:
        """Type text into the active document with realistic typing delay"""
        try:
            if not self._is_connected and not self.open_word():
                return False
                
            if not self.document:
                self.open_document()
            
            logger.info(f"Typing text: {text[:50]}...")
            
            # Show robot cursor for typing
            self.cursor_manager.start_robot_control()
            
            # Set focus to Word
            hwnd = win32gui.FindWindow(None, "Document - Word")
            if hwnd:
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.5)  # Wait for focus
            
            # Use pyautogui for typing
            for char in text:
                pyautogui.write(char)
                time.sleep(delay)
            
            return True
        except Exception as e:
            logger.error(f"Error typing text: {str(e)}")
            return False
        finally:
            # Hide robot cursor when done typing
            self.cursor_manager.end_robot_control()
    
    def move_cursor(self, x: int, y: int) -> bool:
        """Move cursor to specific coordinates"""
        try:
            logger.info(f"Moving cursor to position ({x}, {y})")
            
            # Show robot cursor and move
            self.cursor_manager.start_robot_control()
            self.cursor_manager.move_cursor(x, y, duration=0.5)
            
            return True
        except Exception as e:
            logger.error(f"Error moving cursor: {str(e)}")
            return False
    
    def click(self, x: Optional[int] = None, y: Optional[int] = None, right_click: bool = False) -> bool:
        """Click at current position or specified coordinates"""
        try:
            # Show robot cursor
            self.cursor_manager.start_robot_control()
            
            if x is not None and y is not None:
                self.cursor_manager.move_cursor(x, y)
            
            if right_click:
                pyautogui.rightClick()
            else:
                pyautogui.click()
            
            return True
        except Exception as e:
            logger.error(f"Error clicking: {str(e)}")
            return False
        finally:
            # Hide robot cursor when done clicking
            self.cursor_manager.end_robot_control()
    
    def send_hotkey(self, *keys: str) -> bool:
        """Send keyboard hotkey (e.g., 'ctrl', 'b' for bold)"""
        try:
            # Show robot cursor for keyboard shortcuts
            self.cursor_manager.start_robot_control()
            
            pyautogui.hotkey(*keys)
            return True
        except Exception as e:
            logger.error(f"Error sending hotkey: {str(e)}")
            return False
        finally:
            # Hide robot cursor when done
            self.cursor_manager.end_robot_control()
    
    def save_document(self, path: Optional[str] = None) -> bool:
        """Save document to specified path"""
        try:
            if not self._is_connected or not self.document:
                logger.error("No active document to save")
                return False
            
            if path:
                logger.info(f"Saving document to: {path}")
                self.document.SaveAs(path)
            else:
                logger.info("Saving document")
                self.document.Save()
            
            return True
        except Exception as e:
            logger.error(f"Error saving document: {str(e)}")
            return False
    
    def close(self) -> bool:
        """Close Word application"""
        try:
            if self._is_connected:
                if self.document:
                    self.document.Close(SaveChanges=False)
                self.word_app.Quit()
                self._is_connected = False
                logger.info("Microsoft Word closed")
            return True
        except Exception as e:
            logger.error(f"Error closing Word: {str(e)}")
            return False
    
    def format_text(self, format_type: str, value: str = None) -> bool:
        """Apply formatting to selected text"""
        try:
            if not self._is_connected or not self.document:
                logger.error("No active document for formatting")
                return False
            
            selection = self.word_app.Selection
            
            if format_type.lower() == "bold":
                selection.Font.Bold = True
            elif format_type.lower() == "italic":
                selection.Font.Italic = True
            elif format_type.lower() == "underline":
                selection.Font.Underline = True
            elif format_type.lower() == "size" and value:
                selection.Font.Size = float(value)
            elif format_type.lower() == "font" and value:
                selection.Font.Name = value
            else:
                logger.warning(f"Unsupported format type: {format_type}")
                return False
            
            return True
        except Exception as e:
            logger.error(f"Error applying formatting: {str(e)}")
            return False
    
    def insert_template(self, template_id: str, title: str) -> bool:
        """Insert document template structure based on template ID"""
        try:
            if not self._is_connected:
                self.open_word()
            
            if not self.document:
                self.open_document()
            
            # Find template by ID using local import to avoid circular dependency
            from src.templates.template_manager import TemplateManager
            template_manager = TemplateManager()
            template = template_manager.get_template_by_id(template_id)
            
            if not template:
                logger.error(f"Template not found: {template_id}")
                return False
            
            logger.info(f"Inserting template: {template.name}")
            
            # Set title and create document structure
            self.word_app.Selection.Font.Size = 16
            self.word_app.Selection.Font.Bold = True
            self.word_app.Selection.TypeText(title.upper())
            self.word_app.Selection.TypeParagraph()
            self.word_app.Selection.TypeParagraph()
            
            # Add sections
            self.word_app.Selection.Font.Size = 12
            self.word_app.Selection.Font.Bold = False
            
            for section in template.sections:
                # Skip title as we've already added it
                if section.lower() != "title":
                    self.word_app.Selection.Font.Bold = True
                    self.word_app.Selection.TypeText(f"{section}")
                    self.word_app.Selection.TypeParagraph()
                    self.word_app.Selection.Font.Bold = False
                    self.word_app.Selection.TypeText(f"Content for {section}...")
                    self.word_app.Selection.TypeParagraph()
                    self.word_app.Selection.TypeParagraph()
            
            # Apply formatting settings from template
            if "font" in template.formatting:
                self.word_app.Selection.WholeStory()
                self.word_app.Selection.Font.Name = template.formatting["font"]
            
            return True
        except Exception as e:
            logger.error(f"Error inserting template: {str(e)}")
            return False
