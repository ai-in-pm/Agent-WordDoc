"""
Cursor management for the Word AI Agent
"""

import os
import sys
import time
import logging
import ctypes
from pathlib import Path
from typing import Tuple, Optional
from dataclasses import dataclass

try:
    import pyautogui
    import win32api
    import win32con
    import win32gui
    from PIL import Image, ImageDraw, ImageTk
    HAS_CURSOR_DEPS = True
except ImportError:
    HAS_CURSOR_DEPS = False

from src.core.logging import get_logger

logger = get_logger(__name__)

@dataclass
class CursorSize:
    """Cursor size presets"""
    STANDARD = (32, 32)
    LARGE = (48, 48)
    EXTRA_LARGE = (64, 64)

class RobotCursor:
    """Robot cursor implementation for visual indication when AI is controlling"""
    
    def __init__(self, size: Tuple[int, int] = CursorSize.STANDARD, verbose: bool = False):
        """Initialize the robot cursor"""
        if not HAS_CURSOR_DEPS:
            raise ImportError(
                "Robot cursor dependencies not found. "
                "Please install pyautogui, pillow: pip install pyautogui pillow"
            )
        
        self.size = size
        self.is_visible = False
        self.cursor_image = None
        self.original_position = (0, 0)
        self.verbose = verbose
        self.cursor_assets_dir = Path(__file__).parent.parent / "assets" / "cursors"
        self._create_cursor_assets_dir()
        self._load_robot_cursor()
        
        # Create a more reliable cursor display method
        if os.name == 'nt':
            # Use win32api for Windows cursor handling
            self.cursor_handle = None
            self.cursor_info = win32gui.GetCursorInfo()
            self.original_cursor = self.cursor_info[1]  # Store original cursor handle
    
    def _create_cursor_assets_dir(self):
        """Create directory for cursor assets if it doesn't exist"""
        self.cursor_assets_dir.mkdir(parents=True, exist_ok=True)
    
    def _load_robot_cursor(self):
        """Load robot cursor image from file or generate if not exists"""
        # First, try to use the babyadam.png image from images directory
        babyadam_path = Path(__file__).parent.parent / "images" / "babyadam.png"
        
        if babyadam_path.exists():
            try:
                # Load and resize the babyadam image to the specified cursor size
                original_image = Image.open(babyadam_path)
                
                # Ensure image has alpha channel for transparency
                if original_image.mode != 'RGBA':
                    original_image = original_image.convert('RGBA')
                    
                resized_image = original_image.resize(self.size, Image.LANCZOS)
                
                # Save the resized image to cursor assets directory
                cursor_path = self.cursor_assets_dir / "robot_cursor.png"
                resized_image.save(cursor_path)
                
                if self.verbose:
                    logger.info(f"Using babyadam.png as robot cursor: {cursor_path}")
                
                self.cursor_image = resized_image
                
                # For Windows, create a cursor from the image
                if os.name == 'nt':
                    self._create_windows_cursor(cursor_path)
                    
                return cursor_path
                
            except Exception as e:
                logger.error(f"Failed to load babyadam.png: {str(e)}. Falling back to default cursor.")
                return self._generate_robot_cursor()
        else:
            # Fallback to generating the default robot cursor
            logger.warning(f"babyadam.png not found at {babyadam_path}, generating default cursor")
            return self._generate_robot_cursor()
    
    def _create_windows_cursor(self, cursor_path):
        """Create a Windows cursor from the image file"""
        try:
            # This is a more reliable way to create a custom cursor in Windows
            # Extract icon from image and convert to cursor
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hinst = win32api.GetModuleHandle(None)
            
            # We'll use a temporary icon file approach
            # Convert PNG to ICO format for Windows cursor
            ico_path = self.cursor_assets_dir / "robot_cursor.ico"
            
            # Save as ICO if needed
            if not ico_path.exists():
                img = Image.open(cursor_path)
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')
                img.save(ico_path)
                
            # Create cursor from icon file
            self.cursor_handle = win32gui.LoadImage(
                hinst, str(ico_path), win32con.IMAGE_ICON,
                self.size[0], self.size[1], icon_flags
            )
            
            if self.verbose:
                logger.info(f"Created Windows cursor from {ico_path}")
                
        except Exception as e:
            logger.error(f"Failed to create Windows cursor: {str(e)}")
    
    def _generate_robot_cursor(self):
        """Generate robot cursor image (fallback method)"""
        width, height = self.size
        cursor = Image.new('RGBA', (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(cursor)
        
        # Main robot head (rounded rectangle)
        head_color = (51, 153, 255, 240)  # Blue with transparency
        border_color = (0, 102, 204, 255)  # Darker blue for border
        
        # Calculate proportions based on size
        margin = max(2, width // 16)
        head_width = width - 2 * margin
        head_height = height - 2 * margin
        corner_radius = max(3, width // 10)
        
        # Draw rounded rectangle for head
        rect_bounds = [margin, margin, margin + head_width, margin + head_height]
        draw.rounded_rectangle(rect_bounds, fill=head_color, outline=border_color, radius=corner_radius, width=max(1, width // 32))
        
        # Draw eyes
        eye_radius = max(2, width // 10)
        left_eye_center = (width // 3, height // 3)
        right_eye_center = (width * 2 // 3, height // 3)
        
        draw.ellipse(
            (left_eye_center[0] - eye_radius, left_eye_center[1] - eye_radius, 
             left_eye_center[0] + eye_radius, left_eye_center[1] + eye_radius),
            fill=(255, 255, 255, 255)
        )
        draw.ellipse(
            (right_eye_center[0] - eye_radius, right_eye_center[1] - eye_radius, 
             right_eye_center[0] + eye_radius, right_eye_center[1] + eye_radius),
            fill=(255, 255, 255, 255)
        )
        
        # Draw mouth
        mouth_width = head_width // 2
        mouth_y = height * 2 // 3
        mouth_start = (width // 2 - mouth_width // 2, mouth_y)
        mouth_end = (width // 2 + mouth_width // 2, mouth_y)
        draw.line([mouth_start, mouth_end], fill=(255, 255, 255, 255), width=max(1, width // 32))
        
        # Draw antenna
        antenna_base = (width // 2, margin)
        antenna_top = (width // 2, max(0, margin - margin))
        draw.line([antenna_base, antenna_top], fill=border_color, width=max(1, width // 32))
        draw.ellipse(
            (antenna_top[0] - eye_radius // 2, antenna_top[1] - eye_radius // 2,
             antenna_top[0] + eye_radius // 2, antenna_top[1] + eye_radius // 2),
            fill=(255, 0, 0, 255)  # Red antenna tip
        )
        
        # Save cursor image
        cursor_path = self.cursor_assets_dir / "robot_cursor.png"
        cursor.save(cursor_path)
        logger.info(f"Generated default robot cursor image: {cursor_path}")
        
        self.cursor_image = cursor
        
        # For Windows, create a cursor from the image
        if os.name == 'nt':
            self._create_windows_cursor(cursor_path)
            
        return cursor_path
    
    def show(self):
        """Show the robot cursor"""
        if not HAS_CURSOR_DEPS or self.is_visible:
            return
        
        # Store original position
        self.original_position = pyautogui.position()
        
        # Set to visible
        self.is_visible = True
        
        # Use Windows API for setting cursor if available
        if os.name == 'nt' and self.cursor_handle:
            try:
                # Set the cursor using Windows API
                ctypes.windll.user32.SetCursor(self.cursor_handle)
                if self.verbose:
                    logger.info("Robot cursor activated using Windows API")
            except Exception as e:
                logger.error(f"Failed to set cursor using Windows API: {str(e)}")
        
        logger.info("Robot cursor activated")
    
    def hide(self):
        """Hide the robot cursor"""
        if not HAS_CURSOR_DEPS or not self.is_visible:
            return
        
        # Set to invisible
        self.is_visible = False
        
        # Restore original cursor in Windows
        if os.name == 'nt':
            try:
                # Restore the original cursor
                ctypes.windll.user32.SetCursor(self.original_cursor)
                if self.verbose:
                    logger.info("Restored original cursor using Windows API")
            except Exception as e:
                logger.error(f"Failed to restore original cursor: {str(e)}")
        
        logger.info("Robot cursor deactivated")
    
    def move(self, x: int, y: int, duration: float = 0.5):
        """Move the robot cursor"""
        if not HAS_CURSOR_DEPS or not self.is_visible:
            return
        
        # Move using pyautogui
        pyautogui.moveTo(x, y, duration=duration)
        
        # Re-show the cursor after moving as movement might reset it
        if os.name == 'nt' and self.cursor_handle:
            try:
                ctypes.windll.user32.SetCursor(self.cursor_handle)
            except Exception:
                pass

class CursorManager:
    """Manager for cursor operations"""
    
    def __init__(self, use_robot_cursor: bool = True, size: str = "standard", verbose: bool = False):
        """Initialize cursor manager"""
        self.use_robot_cursor = use_robot_cursor
        self.verbose = verbose
        
        # Set cursor size based on preference
        if size.lower() == "large":
            cursor_size = CursorSize.LARGE
        elif size.lower() == "extra_large":
            cursor_size = CursorSize.EXTRA_LARGE
        else:
            cursor_size = CursorSize.STANDARD
        
        # Create robot cursor if enabled
        self.robot_cursor = None
        if use_robot_cursor and HAS_CURSOR_DEPS:
            try:
                self.robot_cursor = RobotCursor(cursor_size, verbose=verbose)
                logger.info(f"Initialized robot cursor with size: {size}")
            except Exception as e:
                logger.error(f"Failed to initialize robot cursor: {str(e)}")
        
        # Setup periodic cursor refresh for Windows
        if os.name == 'nt' and use_robot_cursor and HAS_CURSOR_DEPS and self.robot_cursor:
            self._setup_cursor_refresh()
    
    def _setup_cursor_refresh(self):
        """Setup a background thread to refresh cursor periodically"""
        import threading
        
        def refresh_cursor():
            """Refresh the robot cursor if it's active"""
            while True:
                try:
                    if self.robot_cursor and self.robot_cursor.is_visible:
                        # Refresh cursor as it may have been reset by Windows
                        if self.robot_cursor.cursor_handle:
                            ctypes.windll.user32.SetCursor(self.robot_cursor.cursor_handle)
                except Exception:
                    pass  # Ignore errors in background thread
                
                time.sleep(0.1)  # Check every 100ms
        
        # Start the refresh thread
        refresh_thread = threading.Thread(target=refresh_cursor, daemon=True)
        refresh_thread.start()
        if self.verbose:
            logger.info("Started cursor refresh thread")
    
    def start_robot_control(self):
        """Show robot cursor to indicate AI control"""
        if self.robot_cursor and self.use_robot_cursor:
            self.robot_cursor.show()
            return True
        return False
    
    def end_robot_control(self):
        """Hide robot cursor when AI control ends"""
        if self.robot_cursor and self.use_robot_cursor:
            self.robot_cursor.hide()
            return True
        return False
    
    def move_cursor(self, x: int, y: int, duration: float = 0.5):
        """Move cursor to position"""
        if self.robot_cursor and self.use_robot_cursor and self.robot_cursor.is_visible:
            self.robot_cursor.move(x, y, duration)
            return True
        elif HAS_CURSOR_DEPS:
            pyautogui.moveTo(x, y, duration=duration)
            return True
        return False
