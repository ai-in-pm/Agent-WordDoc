"""
Control UI for Word AI Agent

This module provides a floating control panel with Pause, Stop, and Continue buttons
to allow the user to control the AI Agent at any time during execution.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from typing import Callable, Dict, Optional

class AgentControlUI:
    """Floating control panel for the AI Agent"""
    
    def __init__(self, 
                 on_pause: Optional[Callable] = None, 
                 on_stop: Optional[Callable] = None, 
                 on_continue: Optional[Callable] = None):
        """
        Initialize the control UI
        
        Args:
            on_pause: Callback function when Pause button is clicked
            on_stop: Callback function when Stop button is clicked
            on_continue: Callback function when Continue button is clicked
        """
        self.root = None
        self.window_shown = False
        self.is_paused = False
        
        # Store callback functions
        self.on_pause = on_pause
        self.on_stop = on_stop
        self.on_continue = on_continue
        
        # Initialize UI in a separate thread
        self.ui_thread = threading.Thread(target=self._initialize_ui)
        self.ui_thread.daemon = True  # Thread will close when main program exits
    
    def _initialize_ui(self):
        """Initialize the UI components"""
        self.root = tk.Tk()
        self.root.title("AI Agent Control")
        
        # Make window always stay on top with high topmost priority
        self.root.attributes("-topmost", True)
        
        # Set window size and position (bottom right corner)
        window_width = 200
        window_height = 100
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_position = screen_width - window_width - 20
        y_position = screen_height - window_height - 60
        
        # Configure window appearance to make it stand out
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        self.root.configure(bg='#E74C3C')  # Bright red background to make it more visible
        self.root.wm_attributes("-alpha", 0.9)  # Slight transparency
        
        # Keep the window on top consistently by periodically lifting it
        def keep_on_top():
            self.root.lift()
            self.root.after(1000, keep_on_top)  # Check every second
        
        keep_on_top()
        
        # Ensure the window cannot be minimized or hidden
        self.root.resizable(False, False)  # Prevent resizing
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)  # Prevent closing with X button
        
        # Create a frame with padding and distinctive background
        main_frame = ttk.Frame(self.root, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create title label with distinctive font and color
        title_label = tk.Label(main_frame, text="AI AGENT CONTROL", 
                              font=("Arial", 10, "bold"),
                              fg="white", bg="#2c3e50")
        title_label.pack(pady=(0, 10), fill=tk.X)
        
        # Create button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, expand=True)
        
        # Configure style for buttons with distinctive colors
        style = ttk.Style()
        style.configure('Pause.TButton', font=('Arial', 9, 'bold'))
        style.configure('Stop.TButton', font=('Arial', 9, 'bold'))
        style.configure('Continue.TButton', font=('Arial', 9, 'bold'))
        
        # Create buttons with distinctive styles
        self.pause_button = ttk.Button(button_frame, text="PAUSE", style='Pause.TButton', command=self._handle_pause)
        self.pause_button.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        self.stop_button = ttk.Button(button_frame, text="STOP", style='Stop.TButton', command=self._handle_stop)
        self.stop_button.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        self.continue_button = ttk.Button(button_frame, text="CONTINUE", style='Continue.TButton', command=self._handle_continue)
        self.continue_button.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        self.continue_button.configure(state="disabled")  # Initially disabled
        
        # Add a status label
        self.status_label = tk.Label(main_frame, text="AI Agent Active", 
                                   font=("Arial", 8), fg="white", bg="#27AE60")
        self.status_label.pack(pady=(5, 0), fill=tk.X)
        
        # Run the UI
        self.window_shown = True
        self.root.mainloop()
    
    def _handle_pause(self):
        """Handle Pause button click"""
        if not self.is_paused and self.on_pause:
            self.on_pause()
        
        self.is_paused = True
        self.pause_button.configure(state="disabled")
        self.continue_button.configure(state="normal")
        self.status_label.configure(text="AI Agent Paused", bg="#F39C12")  # Orange for paused
    
    def _handle_stop(self):
        """Handle Stop button click"""
        if self.on_stop:
            self.on_stop()
        
        # Reset UI state
        self.is_paused = False
        self.pause_button.configure(state="normal")
        self.continue_button.configure(state="disabled")
        self.status_label.configure(text="AI Agent Stopped", bg="#E74C3C")  # Red for stopped
    
    def _handle_continue(self):
        """Handle Continue button click"""
        if self.is_paused and self.on_continue:
            self.on_continue()
        
        self.is_paused = False
        self.pause_button.configure(state="normal")
        self.continue_button.configure(state="disabled")
        self.status_label.configure(text="AI Agent Active", bg="#27AE60")  # Green for active
    
    def show(self):
        """Show the control UI"""
        if not self.window_shown:
            self.ui_thread.start()
    
    def hide(self):
        """Hide the control UI"""
        if self.window_shown and self.root:
            try:
                self.root.destroy()
                self.window_shown = False
            except Exception:
                pass  # Ignore errors if window is already closing
    
    def update_status(self, status_text: str, status_color: str = "#27AE60"):
        """Update status text in UI"""
        if self.window_shown and self.root:
            try:
                self.status_label.configure(text=status_text, bg=status_color)
            except Exception:
                pass  # Ignore errors if window is closing


# Example usage
if __name__ == "__main__":
    # For testing purposes
    def on_pause():
        print("Agent paused")
    
    def on_stop():
        print("Agent stopped")
    
    def on_continue():
        print("Agent continuing")
    
    control_ui = AgentControlUI(
        on_pause=on_pause,
        on_stop=on_stop,
        on_continue=on_continue
    )
    
    control_ui.show()
    
    # Keep the main thread running
    try:
        while True:
            import time
            time.sleep(1)
    except KeyboardInterrupt:
        control_ui.hide()
