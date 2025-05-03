"""
Simple Cursor Manager for AI Agent

This module provides a simplified cursor management system that doesn't rely on
external dependencies like PIL. It changes the cursor appearance to indicate
when the AI Agent is in control.
"""

import os
import time
import ctypes
import threading

class SimpleCursorManager:
    """A simplified cursor manager that uses built-in Windows cursors"""
    
    # Windows cursor constants
    IDC_ARROW = 32512    # Standard arrow
    IDC_CROSS = 32515    # Crosshair
    IDC_WAIT = 32514     # Hourglass
    IDC_IBEAM = 32513    # Text I-beam
    IDC_HAND = 32649     # Hand pointer
    
    def __init__(self, verbose=False):
        """
        Initialize the simple cursor manager
        
        Args:
            verbose: Whether to print debug information
        """
        self.verbose = verbose
        self.original_cursor_handle = None
        self.is_active = False
        self.blink_thread = None
        self.stop_thread = False
        
        # Load user32.dll for cursor operations
        self.user32 = ctypes.windll.user32
        
        if self.verbose:
            print("[CURSOR] Simple cursor manager initialized")
    
    def _set_system_cursor(self, cursor_id):
        """
        Set the system cursor using a built-in Windows cursor
        
        Args:
            cursor_id: Windows cursor ID constant
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Load the cursor
            cursor_handle = self.user32.LoadCursorW(0, cursor_id)
            
            # Set as system cursor
            result = self.user32.SetCursor(cursor_handle)
            
            return result != 0
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to set system cursor: {str(e)}")
            return False
    
    def _save_original_cursor(self):
        """Save the original cursor to restore later"""
        try:
            # Get the current cursor
            self.original_cursor_handle = self.user32.GetCursor()
            
            if self.verbose:
                print("[CURSOR] Original cursor saved")
            
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to save original cursor: {str(e)}")
            return False
    
    def _restore_original_cursor(self):
        """Restore the original cursor"""
        try:
            if self.original_cursor_handle:
                self.user32.SetCursor(self.original_cursor_handle)
                
                if self.verbose:
                    print("[CURSOR] Original cursor restored")
                
                return True
            
            # Fallback to arrow cursor
            return self._set_system_cursor(self.IDC_ARROW)
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to restore original cursor: {str(e)}")
            
            # Try to set arrow cursor as fallback
            return self._set_system_cursor(self.IDC_ARROW)
    
    def _cursor_blink_thread(self, interval=0.5):
        """
        Thread function to make the cursor blink
        
        Args:
            interval: Blink interval in seconds
        """
        while not self.stop_thread:
            # Set wait cursor (hourglass)
            self._set_system_cursor(self.IDC_WAIT)
            time.sleep(interval)
            
            if self.stop_thread:
                break
            
            # Set hand cursor
            self._set_system_cursor(self.IDC_HAND)
            time.sleep(interval)
    
    def activate_robot_cursor(self, blinking=True):
        """
        Activate a special cursor to indicate AI control
        
        Args:
            blinking: Whether the cursor should blink
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if self.is_active:
                return True
            
            # Save the original cursor
            self._save_original_cursor()
            
            if blinking:
                # Start blinking thread
                self.stop_thread = False
                self.blink_thread = threading.Thread(
                    target=self._cursor_blink_thread
                )
                self.blink_thread.daemon = True
                self.blink_thread.start()
            else:
                # Just set the wait cursor (hourglass)
                self._set_system_cursor(self.IDC_WAIT)
            
            self.is_active = True
            
            if self.verbose:
                print("[CURSOR] AI cursor activated")
            
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to activate AI cursor: {str(e)}")
            return False
    
    def deactivate_robot_cursor(self):
        """
        Deactivate the special cursor and restore the original cursor
        
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.is_active:
                return True
            
            # Stop the blinking thread if it's running
            if self.blink_thread and self.blink_thread.is_alive():
                self.stop_thread = True
                self.blink_thread.join(1.0)  # Wait for thread to finish
            
            # Restore the original cursor
            success = self._restore_original_cursor()
            
            if success:
                self.is_active = False
                
                if self.verbose:
                    print("[CURSOR] AI cursor deactivated")
            
            return success
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to deactivate AI cursor: {str(e)}")
            return False
    
    def cleanup(self):
        """Clean up resources"""
        try:
            # Deactivate the special cursor
            self.deactivate_robot_cursor()
            
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to clean up cursor manager: {str(e)}")
            return False
