"""
GUI for Microsoft Word Interface Explorer

Provides a graphical interface for launching the AI Agent's self-learning
and demonstration of Microsoft Word's Home tab features.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import subprocess
from pathlib import Path

# Ensure the src directory is in the Python path
src_path = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(src_path.parent))

from src.bootstrap_training.word_interface.log_analyzer import log_analyzer

class WordExplorerGUI:
    """GUI for Word Interface Explorer"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Word Interface Explorer")
        self.root.geometry("600x450")
        self.root.resizable(True, True)
        
        # Set icon if available
        icon_path = Path(__file__).parent.parent.parent.parent / "assets" / "icons" / "word_icon.ico"
        if icon_path.exists():
            self.root.iconbitmap(str(icon_path))
        
        self.create_widgets()
    
    def create_widgets(self):
        """Create GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Microsoft Word Interface Explorer", 
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Description
        desc_text = (
            "This tool allows the AI Agent to explore and demonstrate "
            "Microsoft Word's Home tab features with physical cursor movements."
        )
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=550)
        desc_label.pack(pady=(0, 20))
        
        # Mode frame
        mode_frame = ttk.LabelFrame(main_frame, text="Exploration Mode", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Mode radio buttons
        self.mode_var = tk.StringVar(value="demonstrate")
        modes = [
            ("Full Interactive Demonstration", "demonstrate", "Complete demonstration of all Home tab elements"),
            ("Explore All Elements", "explore_all", "Document all Home tab elements"),
            ("Demonstrate Specific Element", "element", "Demonstrate a single element")
        ]
        
        for i, (text, value, tooltip) in enumerate(modes):
            rb = ttk.Radiobutton(
                mode_frame, 
                text=text, 
                value=value, 
                variable=self.mode_var,
                command=self.update_interface
            )
            rb.pack(anchor=tk.W, pady=2)
            
            # Add tooltip via label
            tooltip_label = ttk.Label(mode_frame, text=f"  {tooltip}", foreground="gray")
            tooltip_label.pack(anchor=tk.W, padx=(20, 0), pady=(0, 5))
        
        # Element selection (only visible for 'element' mode)
        self.element_frame = ttk.Frame(mode_frame, padding=(20, 5))
        
        element_label = ttk.Label(self.element_frame, text="Select element:")
        element_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.element_var = tk.StringVar(value="Bold")
        elements = ["Paste", "Cut", "Copy", "Bold", "Italic", "Underline", 
                   "Font Color", "Bullet List", "Numbering", "Align Left", 
                   "Center", "Align Right", "Heading 1", "Find", "Replace"]
        
        element_combo = ttk.Combobox(
            self.element_frame, 
            textvariable=self.element_var,
            values=elements,
            width=20
        )
        element_combo.pack(side=tk.LEFT)
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Robot cursor options frame
        robot_cursor_frame = ttk.LabelFrame(options_frame, text="Robot Cursor")
        robot_cursor_frame.pack(fill="x", padx=10, pady=10)
        
        self.use_robot_cursor_var = tk.BooleanVar(value=False)  # Default to False to disable robot cursor
        use_robot_cursor_check = ttk.Checkbutton(
            robot_cursor_frame,
            text="Use Robot Cursor",
            variable=self.use_robot_cursor_var
        )
        use_robot_cursor_check.pack(anchor="w", padx=10, pady=5)
        
        # Robot cursor size options
        cursor_size_frame = ttk.Frame(robot_cursor_frame)
        cursor_size_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(cursor_size_frame, text="Cursor Size:").pack(side="left")
        
        self.cursor_size_var = tk.StringVar(value="standard")
        for size in ["small", "standard", "large", "extra_large"]:
            ttk.Radiobutton(
                cursor_size_frame,
                text=size.replace("_", " ").title(),
                value=size,
                variable=self.cursor_size_var
            ).pack(side="left", padx=5)
        
        # Calibration option
        self.calibrate_var = tk.BooleanVar(value=False)
        calibrate_cb = ttk.Checkbutton(
            options_frame, 
            text="Force element position calibration", 
            variable=self.calibrate_var
        )
        calibrate_cb.pack(anchor=tk.W, pady=2)
        
        # Delay option
        delay_frame = ttk.Frame(options_frame)
        delay_frame.pack(fill=tk.X, pady=(5, 0))
        
        delay_label = ttk.Label(delay_frame, text="Startup delay (seconds):")
        delay_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.delay_var = tk.DoubleVar(value=3.0)
        delay_spin = ttk.Spinbox(
            delay_frame, 
            from_=1.0, 
            to=10.0, 
            increment=0.5, 
            textvariable=self.delay_var,
            width=5
        )
        delay_spin.pack(side=tk.LEFT)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Warning message
        warning_text = (
            "Note: This will physically move your cursor and interact with Microsoft Word. "
            "Please ensure Word is installed and ready."
        )
        warning_label = ttk.Label(
            main_frame, 
            text=warning_text, 
            foreground="red",
            wraplength=550
        )
        warning_label.pack(pady=(0, 20))
        
        # Start button
        start_button_frame = ttk.Frame(main_frame)
        start_button_frame.pack(pady=(0, 10))
        
        self.start_button = ttk.Button(
            start_button_frame, 
            text="Start Exploration", 
            command=self.start_exploration,
            style="Accent.TButton",
            padding=(20, 10)
        )
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Execute button for real-time demonstrations
        self.execute_button = ttk.Button(
            start_button_frame, 
            text="Execute Now", 
            command=self.execute_demonstration,
            style="Accent.TButton",
            padding=(20, 10)
        )
        self.execute_button.pack(side=tk.LEFT)
        
        # Apply style for the accent button if supported
        try:
            self.root.tk.call("source", "azure.tcl")
            self.root.tk.call("set_theme", "light")
        except tk.TclError:
            # Use standard styling if custom theme not available
            pass
        
        # Update interface based on initial values
        self.update_interface()
    
    def update_interface(self):
        """Update interface based on selected options"""
        # Show/hide element selection based on mode
        if self.mode_var.get() == "element":
            self.element_frame.pack(anchor=tk.W, pady=(5, 0))
        else:
            self.element_frame.pack_forget()
    
    def start_exploration(self):
        """Start the Word interface exploration"""
        # Build command
        script_path = Path(__file__).parent.parent / "run_explorer.py"
        cmd = [sys.executable, str(script_path)]
        
        # Add mode argument
        mode = self.mode_var.get()
        if mode == "demonstrate":
            cmd.append("--demonstrate")
        elif mode == "explore_all":
            cmd.append("--explore-all")
        elif mode == "element":
            cmd.extend(["--element", self.element_var.get()])
        
        # Add options
        if self.use_robot_cursor_var.get():
            cmd.append("--robot-cursor")
        else:
            cmd.append("--no-robot-cursor")
        
        cmd.extend(["--cursor-size", self.cursor_size_var.get()])
        
        if self.calibrate_var.get():
            cmd.append("--calibrate")
        
        cmd.extend(["--delay", str(self.delay_var.get())])
        
        # Display command
        command_str = " ".join(cmd)
        print(f"Running command: {command_str}")
        
        # Disable the start button during execution
        self.start_button.configure(state="disabled")
        self.start_button["text"] = "Exploration Running..."
        
        # Run the command in a separate thread
        def run_command():
            try:
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                stdout, stderr = process.communicate()
                
                # Log output
                if stdout:
                    print("Output:\n", stdout)
                if stderr:
                    print("Errors:\n", stderr)
                
                # Re-enable the start button
                self.root.after(0, self._enable_start_button)
                
                # Show completion message
                if process.returncode == 0:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Exploration Complete", 
                        "The Word interface exploration has completed successfully!"
                    ))
                else:
                    self.root.after(0, lambda: messagebox.showerror(
                        "Exploration Failed", 
                        f"The exploration failed with exit code {process.returncode}.\n\nError: {stderr}"
                    ))
                    
            except Exception as e:
                print(f"Error running explorer: {e}")
                self.root.after(0, self._enable_start_button)
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", 
                    f"Failed to start exploration: {str(e)}"
                ))
        
        # Start the thread
        threading.Thread(target=run_command, daemon=True).start()
    
    def _enable_start_button(self):
        """Re-enable the start button"""
        self.start_button.configure(state="normal")
        self.start_button["text"] = "Start Exploration"
        self.execute_button.configure(state="normal")
        self.execute_button["text"] = "Execute Now"
    
    def execute_demonstration(self):
        """Execute a real-time demonstration with the AI Agent"""
        # Disable buttons during execution
        self.start_button.configure(state="disabled")
        self.execute_button.configure(state="disabled")
        self.execute_button["text"] = "Executing..."
        
        # Get selected options
        mode = self.mode_var.get()
        element = self.element_var.get() if mode == "element" else None
        
        # Store current settings for potential retry
        self.current_settings = {
            'mode': mode,
            'element': element,
            'robot_cursor': self.use_robot_cursor_var.get(),
            'cursor_size': self.cursor_size_var.get(),
            'calibrate': self.calibrate_var.get()
        }
        
        # Execute with current settings
        self._execute_with_settings(self.current_settings)
    
    def _execute_with_settings(self, settings):
        """Execute a demonstration with the given settings"""
        # Build command with real-time flag
        script_path = Path(__file__).parent.parent / "run_explorer.py"
        cmd = [sys.executable, str(script_path)]
        
        # Add mode argument
        if settings['mode'] == "demonstrate":
            cmd.append("--demonstrate")
        elif settings['mode'] == "explore_all":
            cmd.append("--explore-all")
        elif settings['mode'] == "element":
            cmd.extend(["--element", settings['element']])
        
        # Add options
        if settings['robot_cursor']:
            cmd.append("--robot-cursor")
        else:
            cmd.append("--no-robot-cursor")
        
        cmd.extend(["--cursor-size", settings['cursor_size']])
        
        if settings['calibrate']:
            cmd.append("--calibrate")
        
        # No delay for immediate execution
        cmd.extend(["--delay", "0.5"])
        
        # Run the command immediately (minimize GUI window to avoid interfering)
        self.root.iconify()
        
        # Run the command in a separate thread
        def run_command():
            try:
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                stdout, stderr = process.communicate()
                
                # Log output
                if stdout:
                    print("Output:\n", stdout)
                if stderr:
                    print("Errors:\n", stderr)
                
                # Restore window and re-enable buttons
                self.root.after(0, self.root.deiconify)
                self.root.after(0, self._enable_start_button)
                
                # Show completion or error message
                if process.returncode == 0:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Demonstration Complete", 
                        "The AI Agent has completed the demonstration!"
                    ))
                else:
                    # Show error dialog with retry option
                    self.root.after(0, lambda: self._show_retry_dialog(
                        "Demonstration Failed", 
                        f"The demonstration failed with exit code {process.returncode}.\n\nError: {stderr}"
                    ))
                    
            except Exception as e:
                print(f"Error running demonstration: {e}")
                self.root.after(0, self.root.deiconify)
                self.root.after(0, self._enable_start_button)
                self.root.after(0, lambda: self._show_retry_dialog(
                    "Error", 
                    f"Failed to start demonstration: {str(e)}"
                ))
        
        # Start the thread
        threading.Thread(target=run_command, daemon=True).start()
    
    def _show_retry_dialog(self, title, message):
        """Show a custom dialog with Retry and End buttons"""
        # Create a custom dialog
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("500x350")
        dialog.transient(self.root)  # Make dialog modal
        dialog.grab_set()
        
        # Add icon
        try:
            dialog.iconbitmap(self.root.iconbitmap())
        except:
            pass  # Ignore if icon setting fails
        
        # Configure dialog appearance
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Error icon
        icon_frame = ttk.Frame(frame)
        icon_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Create a red X icon
        canvas = tk.Canvas(icon_frame, width=40, height=40, bg=dialog.cget('bg'), highlightthickness=0)
        canvas.pack(side=tk.LEFT, padx=(0, 15))
        
        # Draw red circle with X
        canvas.create_oval(5, 5, 35, 35, fill="#FF4444", outline="")
        canvas.create_line(12, 12, 28, 28, fill="white", width=3)
        canvas.create_line(28, 12, 12, 28, fill="white", width=3)
        
        # Error message
        error_title = ttk.Label(icon_frame, text=title, font=("Arial", 14, "bold"))
        error_title.pack(side=tk.LEFT, anchor=tk.W)
        
        # Scrollable error details
        details_frame = ttk.Frame(frame)
        details_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(details_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add text widget for error message
        text_widget = tk.Text(details_frame, wrap=tk.WORD, height=10, 
                              yscrollcommand=scrollbar.set)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)
        
        # Insert error message
        text_widget.insert(tk.END, message)
        text_widget.config(state=tk.DISABLED)  # Make read-only
        
        # Buttons frame
        buttons_frame = ttk.Frame(frame)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Retry button
        retry_button = ttk.Button(
            buttons_frame,
            text="Retry",
            style="Accent.TButton",
            command=lambda: self._retry_action(dialog)
        )
        retry_button.pack(side=tk.RIGHT, padx=5)
        
        # End button
        end_button = ttk.Button(
            buttons_frame,
            text="End",
            command=dialog.destroy
        )
        end_button.pack(side=tk.RIGHT, padx=5)
        
        # Center the dialog on the parent window
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Set retry button as default
        retry_button.focus_set()
        
        # Bind Enter key to retry
        dialog.bind('<Return>', lambda event: self._retry_action(dialog))
        
        # Bind Escape key to end
        dialog.bind('<Escape>', lambda event: dialog.destroy())
    
    def _retry_action(self, dialog):
        """Retry the current action with intelligent adaptation based on log analysis"""
        # Close the dialog
        dialog.destroy()
        
        # Create analyzing dialog
        self._show_analyzing_dialog()
        
        # Analyze logs in a separate thread to keep UI responsive
        def analyze_and_retry():
            try:
                # Analyze recent logs
                analysis_result = log_analyzer.analyze_logs(max_age_minutes=5)
                
                # Get retry recommendations
                recommendations = log_analyzer.generate_retry_recommendations(analysis_result)
                
                # Apply recommended settings
                modified_settings = self.current_settings.copy()
                modified_settings.update(recommendations.get("settings", {}))
                
                # Show analysis results
                self.root.after(0, lambda: self._close_analyzing_dialog())
                
                if recommendations["recommendation"] != "standard_retry":
                    # Show what was learned from the logs
                    self.root.after(0, lambda: self._show_learning_dialog(
                        recommendations["message"],
                        recommendations.get("solutions", [])
                    ))
                
                # Re-execute with potentially modified settings
                self.root.after(500, lambda: self._execute_with_settings(modified_settings))
                
            except Exception as e:
                # If analysis fails, fall back to standard retry
                print(f"Error during log analysis: {e}")
                self.root.after(0, lambda: self._close_analyzing_dialog())
                self.root.after(500, lambda: self._execute_with_settings(self.current_settings))
        
        # Start analysis thread
        threading.Thread(target=analyze_and_retry, daemon=True).start()
    
    def _show_analyzing_dialog(self):
        """Show a dialog that indicates log analysis is in progress"""
        self.analyzing_dialog = tk.Toplevel(self.root)
        self.analyzing_dialog.title("Analyzing Previous Attempts")
        self.analyzing_dialog.geometry("400x150")
        self.analyzing_dialog.transient(self.root)  # Make dialog modal
        self.analyzing_dialog.grab_set()
        
        # Configure dialog appearance
        frame = ttk.Frame(self.analyzing_dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Add spinner or progress indicator
        ttk.Label(
            frame, 
            text="AI Agent is analyzing failure logs...",
            font=("Arial", 12)
        ).pack(pady=(0, 10))
        
        progress = ttk.Progressbar(frame, mode='indeterminate')
        progress.pack(fill=tk.X, pady=10)
        progress.start(10)
        
        ttk.Label(
            frame, 
            text="Identifying issues and adjusting settings for retry",
            wraplength=350
        ).pack(pady=10)
        
        # Center the dialog on the parent window
        self.analyzing_dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - self.analyzing_dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - self.analyzing_dialog.winfo_height()) // 2
        self.analyzing_dialog.geometry(f"+{x}+{y}")
    
    def _close_analyzing_dialog(self):
        """Close the analyzing dialog if it exists"""
        if hasattr(self, 'analyzing_dialog') and self.analyzing_dialog.winfo_exists():
            self.analyzing_dialog.destroy()
    
    def _show_learning_dialog(self, message, solutions):
        """Show what the AI Agent learned from log analysis"""
        learning_dialog = tk.Toplevel(self.root)
        learning_dialog.title("AI Agent Learning")
        learning_dialog.geometry("500x300")
        learning_dialog.transient(self.root)  # Make dialog modal
        learning_dialog.grab_set()
        
        # Configure dialog appearance
        frame = ttk.Frame(learning_dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Add light bulb icon or similar for "learning"
        icon_frame = ttk.Frame(frame)
        icon_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Create a lightbulb icon
        canvas = tk.Canvas(icon_frame, width=40, height=40, bg=learning_dialog.cget('bg'), highlightthickness=0)
        canvas.pack(side=tk.LEFT, padx=(0, 15))
        
        # Draw lightbulb - yellow circle with radiating lines
        canvas.create_oval(10, 10, 30, 30, fill="#FFCC00", outline="")
        for angle in range(0, 360, 45):  # Draw 8 rays
            import math
            rad = math.radians(angle)
            x1, y1 = 20 + 15 * math.cos(rad), 20 + 15 * math.sin(rad)
            x2, y2 = 20 + 20 * math.cos(rad), 20 + 20 * math.sin(rad)
            canvas.create_line(x1, y1, x2, y2, fill="#FFCC00", width=2)
        
        # Learning title
        ttk.Label(
            icon_frame, 
            text="AI Agent Has Learned From Previous Attempts", 
            font=("Arial", 12, "bold")
        ).pack(side=tk.LEFT, anchor=tk.W)
        
        # Message about what was learned
        ttk.Label(
            frame, 
            text=message,
            wraplength=450,
            justify=tk.LEFT
        ).pack(fill=tk.X, pady=(0, 15))
        
        # Solutions frame with heading
        if solutions:
            solutions_frame = ttk.LabelFrame(frame, text="Solutions Applied", padding=10)
            solutions_frame.pack(fill=tk.BOTH, expand=True)
            
            for solution in solutions:
                ttk.Label(
                    solutions_frame,
                    text=f"• {solution}",
                    wraplength=430,
                    justify=tk.LEFT
                ).pack(anchor=tk.W, pady=2)
        
        # Continue button
        ttk.Button(
            frame,
            text="Continue with Retry",
            command=learning_dialog.destroy,
            style="Accent.TButton"
        ).pack(pady=(15, 0))
        
        # Center the dialog on the parent window
        learning_dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - learning_dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - learning_dialog.winfo_height()) // 2
        learning_dialog.geometry(f"+{x}+{y}")
        
        # Set focus to continue button
        learning_dialog.bind('<Return>', lambda event: learning_dialog.destroy())
        learning_dialog.bind('<Escape>', lambda event: learning_dialog.destroy())
    
def main():
    """Main function"""
    root = tk.Tk()
    app = WordExplorerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
