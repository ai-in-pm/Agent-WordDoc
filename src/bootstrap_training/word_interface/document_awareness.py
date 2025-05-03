"""
Microsoft Word Document Awareness Module

Enables the AI Agent to observe and understand the content and state
of the Microsoft Word document in real-time.
"""

import time
import os
import re
import logging
from typing import Dict, List, Tuple, Optional, Any
import json
from pathlib import Path
import tempfile

# Configure logging
logger = logging.getLogger(__name__)

try:
    import win32com.client
    import win32gui
    import win32con
    import win32api
    import pyautogui
    from PIL import Image, ImageGrab
    HAS_DEPS = True
except ImportError:
    HAS_DEPS = False

# Optional OCR support
try:
    import pytesseract
    HAS_OCR = True
    
    # Set explicit path to Tesseract executable
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    logger.info(f"Set Tesseract path to: {pytesseract.pytesseract.tesseract_cmd}")
    
except ImportError:
    HAS_OCR = False

# OCR configuration
OCR_CONFIG = r'--oem 3 --psm 11'

# Check for Tesseract executable
tesseract_present = False
if HAS_OCR:
    try:
        # Use a very simple test to see if Tesseract works
        test_text = pytesseract.image_to_string(Image.new('RGB', (10, 10), color=(255, 255, 255)))
        tesseract_present = True
        logger.info("Tesseract OCR is properly installed and working")
    except Exception as e:
        logger.warning(f"Tesseract OCR is not properly installed: {e}")
        tesseract_present = False
        logger.warning("Please install Tesseract OCR from https://github.com/UB-Mannheim/tesseract/wiki")
        logger.warning("Make sure to add Tesseract to your PATH during installation")
        # Don't let this error propagate and cause the whole module to fail
        HAS_OCR = False  # Disable OCR functionality if Tesseract isn't working

# Ensure we have a proper status message for users
if not HAS_OCR:
    logger.info("OCR functionality is not available. The AI Agent will function with limited visual capabilities.")
else:
    if tesseract_present:
        logger.info("OCR functionality is available and working properly.")
    else:
        logger.info("OCR is disabled due to Tesseract issues. To fix, install Tesseract OCR and add it to PATH.")

class DocumentAwareness:
    """Provides awareness of Microsoft Word document content and state"""
    
    def __init__(self, word_app=None):
        """Initialize with an optional Word application instance"""
        if not HAS_DEPS:
            raise ImportError(
                "Document awareness dependencies not found. "
                "Please install: pywin32, pyautogui, pillow, pytesseract"
            )
        
        self.word_app = word_app
        self.last_content = ""
        self.last_selection = ""
        self.last_ui_elements = {}
        self.dialog_active = False
        self.dialog_text = ""
        self.ocr_available = HAS_OCR and tesseract_present
        
        # Try to set pytesseract path
        if os.name == 'nt' and HAS_OCR:
            # Try multiple possible paths
            tesseract_paths = [
                r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                r'C:\Tesseract-OCR\tesseract.exe'
            ]
            
            for path in tesseract_paths:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    logger.info(f"Found Tesseract at {path}")
                    self.ocr_available = True
                    break
    
    def connect_to_word(self, word_app=None):
        """Connect to a Word application instance"""
        if word_app:
            self.word_app = word_app
        else:
            try:
                # Try to connect to existing instance
                self.word_app = win32com.client.GetActiveObject("Word.Application")
            except:
                # Create new instance
                self.word_app = win32com.client.Dispatch("Word.Application")
                self.word_app.Visible = True
        
        return self.word_app is not None
    
    def get_document_content(self) -> str:
        """Get the text content of the active document"""
        if not self.word_app:
            return ""
        
        try:
            doc = self.word_app.ActiveDocument
            content = doc.Content.Text
            self.last_content = content
            return content
        except Exception as e:
            logger.error(f"Error getting document content: {e}")
            return ""
    
    def get_current_selection(self) -> str:
        """Get the currently selected text"""
        if not self.word_app:
            return ""
        
        try:
            selection = self.word_app.Selection.Text
            self.last_selection = selection
            return selection
        except Exception as e:
            logger.error(f"Error getting selection: {e}")
            return ""
    
    def get_document_properties(self) -> Dict[str, Any]:
        """Get properties of the active document"""
        if not self.word_app:
            return {}
        
        try:
            doc = self.word_app.ActiveDocument
            
            # Create constants locally to avoid dependency on win32com.client.constants
            # which can cause errors if not properly initialized
            WD_STAT_PAGES = 2      # wdStatisticPages
            WD_STAT_PARAGRAPHS = 4 # wdStatisticParagraphs
            WD_STAT_WORDS = 0      # wdStatisticWords
            WD_STAT_CHARS = 3      # wdStatisticCharacters
            WD_STAT_LINES = 5      # wdStatisticLines
            
            props = {
                "name": doc.Name,
                "path": doc.Path,
            }
            
            # Try to get statistics safely
            try:
                props["pages"] = doc.ComputeStatistics(WD_STAT_PAGES)
            except Exception:
                props["pages"] = "Unknown"
                
            try:
                props["paragraphs"] = doc.ComputeStatistics(WD_STAT_PARAGRAPHS)
            except Exception:
                props["paragraphs"] = "Unknown"
                
            try:
                props["words"] = doc.ComputeStatistics(WD_STAT_WORDS)
            except Exception:
                props["words"] = "Unknown"
                
            try:
                props["characters"] = doc.ComputeStatistics(WD_STAT_CHARS)
            except Exception:
                props["characters"] = "Unknown"
                
            try:
                props["lines"] = doc.ComputeStatistics(WD_STAT_LINES)
            except Exception:
                props["lines"] = "Unknown"
            
            return props
        except Exception as e:
            logger.error(f"Error getting document properties: {e}")
            return {"error": "Could not retrieve document properties"}
    
    def detect_active_dialog(self) -> Dict[str, Any]:
        """Detect if a dialog is active and extract its content"""
        if not self.word_app:
            return {"has_dialog": False}
        
        try:
            # First try COM method for getting dialog info (may not work for all dialogs)
            try:
                active_window = self.word_app.ActiveWindow
                if hasattr(active_window, "Type") and active_window.Type == 2:  # 2 = Dialog
                    return {
                        "has_dialog": True,
                        "method": "com",
                        "dialog_caption": self.word_app.Caption
                    }
            except Exception:
                pass  # Fallback to visual detection
            
            # Check if any dialog is visible by taking a screenshot of entire Word window
            hwnd = self._get_word_window_handle()
            if not hwnd:
                return {"has_dialog": False}
            
            # Get dialog detection screenshot
            screenshot = self._capture_full_window(hwnd)
            if screenshot is None:
                return {"has_dialog": False}
            
            # Use OCR to determine if there's a dialog (if OCR is available)
            dialog_present = False
            dialog_content = ""
            
            if HAS_OCR and tesseract_present:
                try:
                    # Extract text using OCR
                    dialog_content = pytesseract.image_to_string(screenshot, config=OCR_CONFIG)
                    
                    # Check for common dialog patterns (buttons, title patterns, etc.)
                    dialog_markers = ["OK", "Cancel", "Yes", "No", "Error", "Warning", "Alert"]
                    if any(marker in dialog_content for marker in dialog_markers):
                        dialog_present = True
                except Exception as e:
                    logger.error(f"Dialog OCR detection failed: {e}")
            else:
                # Without OCR, try a simple pixel analysis to detect dialog-like UI elements
                # This is a simplistic approach and won't be as reliable as OCR
                dialog_present = self._detect_dialog_by_pixels(screenshot)
            
            if dialog_present:
                return {
                    "has_dialog": True,
                    "method": "visual",
                    "dialog_content": dialog_content if HAS_OCR and tesseract_present else "OCR not available"
                }
            
            return {"has_dialog": False}
        except Exception as e:
            logger.error(f"Error detecting dialog: {e}")
            return {"has_dialog": False, "error": str(e)}
            
    def _detect_dialog_by_pixels(self, image: Image.Image) -> bool:
        """Attempt to detect a dialog by analyzing the image pixel patterns.
        This is a fallback when OCR is not available."""
        try:
            # Convert to grayscale for simpler processing
            gray_image = image.convert('L')
            
            # Check for regions with contrasting rectangles (potential dialog boxes)
            width, height = gray_image.size
            pixels = gray_image.load()
            
            # Look for rectangular shapes that might be dialogs (simplified approach)
            rect_count = 0
            edges = 0
            
            # Sample the image at a lower resolution for speed
            sample_step = 10
            for y in range(0, height - sample_step, sample_step):
                for x in range(0, width - sample_step, sample_step):
                    # Check for horizontal edges
                    if abs(pixels[x, y] - pixels[x + sample_step, y]) > 50:
                        edges += 1
                    # Check for vertical edges
                    if abs(pixels[x, y] - pixels[x, y + sample_step]) > 50:
                        edges += 1
            
            # If we have many edges, it might indicate UI elements like dialogs
            # This is a very simplistic heuristic and will need tuning
            dialog_threshold = (width * height) / (sample_step * sample_step * 50)
            
            return edges > dialog_threshold
        except Exception as e:
            logger.error(f"Error in dialog pixel detection: {e}")
            return False
    
    def get_ui_state(self) -> Dict[str, Any]:
        """Get the current state of the Word UI"""
        try:
            result = {}
            
            # Detect dialog first
            dialog_info = self.detect_active_dialog()
            result["dialog"] = dialog_info
            
            # If no dialog, get document info
            if not dialog_info.get("has_dialog", False):
                # Document content
                result["content"] = self.get_document_content()
                
                # Current selection
                result["selection"] = self.get_current_selection()
                
                # Document properties
                result["properties"] = self.get_document_properties()
                
                # Active tab (try to detect which ribbon tab is active)
                # This requires OCR since it's not directly accessible via COM
                if self.ocr_available:
                    try:
                        # Capture the ribbon area
                        ribbon_bbox = self._get_ribbon_bbox()
                        if ribbon_bbox:
                            ribbon_img = ImageGrab.grab(bbox=ribbon_bbox)
                            ribbon_text = pytesseract.image_to_string(ribbon_img, config=OCR_CONFIG)
                            
                            # Check for common tab names
                            tabs = ["Home", "Insert", "Design", "Layout", "References", 
                                    "Mailings", "Review", "View", "Developer"]
                            
                            found_tabs = [tab for tab in tabs if tab.lower() in ribbon_text.lower()]
                            result["active_tab"] = found_tabs[0] if found_tabs else "Unknown"
                    except Exception as e:
                        logger.error(f"Error detecting active tab: {e}")
                        result["active_tab"] = "Unknown"
                else:
                    # Without OCR, just assume Home tab
                    result["active_tab"] = "Home"
            
            return result
        
        except Exception as e:
            logger.error(f"Error getting UI state: {e}")
            return {"error": str(e)}
    
    def _get_ribbon_bbox(self) -> Tuple[int, int, int, int]:
        """Get the bounding box of the ribbon area"""
        try:
            # Find Word window
            hwnd = win32gui.FindWindow(None, "Document - Word")
            if not hwnd:
                # Try different window titles
                for title in ["Document1 - Word", "*.docx - Word"]:
                    hwnd = win32gui.FindWindowEx(None, None, None, title)
                    if hwnd:
                        break
            
            if hwnd:
                # Get window rect
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                
                # Ribbon is typically in the top ~100 pixels
                ribbon_height = 100
                return (left, top, right, top + ribbon_height)
            
            return None
        
        except Exception as e:
            logger.error(f"Error getting ribbon bbox: {e}")
            return None
    
    def capture_document_image(self) -> Optional[Image.Image]:
        """Capture an image of the document content area"""
        try:
            # Get the active window handle for Microsoft Word
            hwnd = win32gui.FindWindow("OpusApp", None)  # "OpusApp" is Word's window class
            
            if hwnd:
                # Get window position
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                win_width = right - left
                win_height = bottom - top
                
                # Document area is below ribbon (~100px) and above status bar (~30px)
                doc_top = top + 100
                doc_left = left + int(win_width * 0.05)    # Skip navigation pane if visible
                doc_right = right - int(win_width * 0.05)  # Skip scrollbar area
                doc_bottom = bottom - int(win_height * 0.05)  # Skip status bar
                
                # Capture just the document area
                doc_image = ImageGrab.grab(bbox=(doc_left, doc_top, doc_right, doc_bottom))
                return doc_image
            
            return None
        except Exception as e:
            logger.error(f"Error capturing document image: {e}")
            return None
    
    def observe_document(self) -> Dict[str, Any]:
        """Comprehensive observation of the Word document and UI state"""
        # Get UI state (includes dialog detection)
        ui_state = self.get_ui_state()
        
        # Capture document image for visual analysis
        doc_image = self.capture_document_image()
        if doc_image:
            # Save to temp file for potential further processing
            temp_img_path = os.path.join(tempfile.gettempdir(), "word_doc_capture.png")
            doc_image.save(temp_img_path)
            ui_state["document_image_path"] = temp_img_path
            
            # Extract text from the image as backup if COM fails
            if self.ocr_available:
                try:
                    image_text = pytesseract.image_to_string(doc_image, config=OCR_CONFIG)
                    ui_state["image_text"] = image_text
                except Exception as e:
                    logger.error(f"Error extracting text from image: {e}")
        
        # Return comprehensive document awareness data
        return ui_state
    
    def recognize_formatting(self, selected_text: str = None) -> Dict[str, Any]:
        """Recognize formatting attributes of the current selection or text"""
        if not self.word_app:
            return {}
        
        try:
            selection = self.word_app.Selection
            
            # If specific text provided, try to find and select it
            if selected_text:
                find_object = self.word_app.Selection.Find
                find_object.Text = selected_text
                find_object.Execute()
            
            # Get formatting attributes
            font = selection.Font
            paragraph = selection.ParagraphFormat
            
            formatting = {
                "font": {
                    "name": font.Name,
                    "size": font.Size,
                    "bold": font.Bold,
                    "italic": font.Italic,
                    "underline": font.Underline,
                    "color": f"RGB({font.Color & 0xFF}, {(font.Color >> 8) & 0xFF}, {(font.Color >> 16) & 0xFF})",
                    "highlight": bool(font.Highlight),
                    "superscript": font.Superscript,
                    "subscript": font.Subscript,
                    "strikethrough": font.StrikeThrough,
                },
                "paragraph": {
                    "alignment": self._get_alignment_name(paragraph.Alignment),
                    "left_indent": paragraph.LeftIndent,
                    "right_indent": paragraph.RightIndent,
                    "first_line_indent": paragraph.FirstLineIndent,
                    "space_before": paragraph.SpaceBefore,
                    "space_after": paragraph.SpaceAfter,
                    "line_spacing": paragraph.LineSpacing,
                }
            }
            
            return formatting
            
        except Exception as e:
            logger.error(f"Error recognizing formatting: {e}")
            return {}
    
    def _get_alignment_name(self, alignment_value):
        """Convert alignment value to name"""
        alignment_map = {
            0: "Left",
            1: "Center",
            2: "Right",
            3: "Justified"
        }
        return alignment_map.get(alignment_value, "Unknown")

    def capture_visual_content(self) -> Dict[str, Any]:
        """Capture visual content of the document window using screenshot and optional OCR"""
        if not self.word_app:
            return {}
        
        try:
            # Get the active window handle for Microsoft Word
            hwnd = self._get_word_window_handle()
            if not hwnd:
                logger.warning("Could not find Microsoft Word window")
                return {}
            
            # Take a screenshot of the document area
            screenshot = self._capture_document_area(hwnd)
            if screenshot is None:
                return {}
            
            result = {
                "has_screenshot": True,
                "screenshot_width": screenshot.width,
                "screenshot_height": screenshot.height,
            }
            
            # Perform OCR if available
            if HAS_OCR and tesseract_present:
                try:
                    text = pytesseract.image_to_string(screenshot, config=OCR_CONFIG)
                    result["ocr_text"] = text.strip()
                    result["has_ocr"] = True
                except Exception as e:
                    logger.error(f"OCR failed: {e}")
                    result["ocr_error"] = str(e)
                    result["has_ocr"] = False
            else:
                result["has_ocr"] = False
                result["ocr_status"] = "Tesseract OCR not available"
            
            return result
        except Exception as e:
            logger.error(f"Error capturing visual content: {e}")
            return {"error": str(e)}

    def _get_word_window_handle(self) -> int:
        """Get the handle for the Microsoft Word window"""
        try:
            if not self.word_app:
                return 0
                
            # Get the caption of the Word window
            caption = self.word_app.Caption
            
            # Helper function to find the Word window
            def enum_window_callback(hwnd, result):
                window_text = win32gui.GetWindowText(hwnd)
                # Look for Word's window containing the caption
                if caption in window_text and win32gui.IsWindowVisible(hwnd):
                    result.append(hwnd)
                return True
            
            result = []
            win32gui.EnumWindows(enum_window_callback, result)
            
            if result:
                return result[0]
            
            # Fallback: try to find Word's window by class name
            word_hwnd = win32gui.FindWindow("OpusApp", None)  # Word's main window class
            if word_hwnd:
                return word_hwnd
                
            return 0
        except Exception as e:
            logger.error(f"Error getting Word window handle: {e}")
            return 0
            
    def _capture_document_area(self, hwnd) -> Optional[Image.Image]:
        """Capture the document area of Word"""
        try:
            if not hwnd:
                return None
                
            # Get window position and dimensions
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            
            # Estimate the document area (excluding ribbon, status bar, etc.)
            # These values are approximate and may need adjustment
            doc_top = top + int(height * 0.15)     # Skip ribbon area
            doc_left = left + int(width * 0.05)    # Skip navigation pane if visible
            doc_right = right - int(width * 0.05)  # Skip scrollbar area
            doc_bottom = bottom - int(height * 0.05)  # Skip status bar
            
            # Capture the document area
            doc_area = (doc_left, doc_top, doc_right, doc_bottom)
            screenshot = ImageGrab.grab(bbox=doc_area)
            
            return screenshot
        except Exception as e:
            logger.error(f"Error capturing document area: {e}")
            return None
            
    def _capture_full_window(self, hwnd) -> Optional[Image.Image]:
        """Capture the entire Word window"""
        try:
            if not hwnd:
                return None
                
            # Get window position
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            
            # Capture the entire window
            screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
            
            return screenshot
        except Exception as e:
            logger.error(f"Error capturing full window: {e}")
            return None
