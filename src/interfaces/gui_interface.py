"""
GUI Interface for the Word AI Agent

Provides a graphical user interface for interaction with the agent.
"""

import os
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.scrolledtext as scrolledtext
import asyncio
import threading
import time
from typing import Dict, Any, List, Optional, Callable

from src.agents.word_agent import AgentFactory
from src.config.config import Config
from src.core.logging import get_logger, QueueHandler
from src.templates.template_manager import TemplateManager, PaperTemplate
from src.interfaces.template_dialog import TemplateSelectionDialog

logger = get_logger(__name__)

class AgentGUI(tk.Tk):
    """Graphical User Interface for the Word AI Agent"""
    
    def __init__(self):
        super().__init__()
        
        self.title("Word AI Agent")
        self.geometry("800x600")
        self.minsize(600, 400)
        
        self.agent_config = self._load_default_config()
        self.log_queue = queue.Queue()
        self.is_running = False
        self.agent_thread = None
        
        self._create_widgets()
        self._create_menu()
        self._setup_log_monitor()
    
    def _load_default_config(self) -> Config:
        """Load default configuration"""
        return Config(
            api_key="",
            typing_mode="realistic",
            verbose=True,
            iterative=True,
            self_improve=True,
            self_evolve=True,
            track_position=True,
            robot_cursor=True,
            use_autoit=True,
            delay=3.0,
            max_retries=3,
            retry_delay=1.0,
            log_level="INFO"
        )
    
    def _create_widgets(self):
        """Create GUI widgets"""
        # Main container with padding
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top frame for inputs
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Topic input
        ttk.Label(top_frame, text="Topic:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.topic_var = tk.StringVar(value="Earned Value Management")
        ttk.Entry(top_frame, textvariable=self.topic_var, width=50).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # API Key input
        ttk.Label(top_frame, text="API Key:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.api_key_var = tk.StringVar()
        api_key_entry = ttk.Entry(top_frame, textvariable=self.api_key_var, width=50, show="*")
        api_key_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Template selection
        ttk.Label(top_frame, text="Template:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.template_var = tk.StringVar()
        template_frame = ttk.Frame(top_frame)
        template_frame.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        self.template_manager = TemplateManager()
        self.selected_template = None
        
        # Create template dropdown
        template_names = self.template_manager.get_template_names()
        if template_names:
            self.template_var.set(template_names[0])
        
        template_dropdown = ttk.Combobox(template_frame, textvariable=self.template_var, 
                                       values=template_names, width=40, state="readonly")
        template_dropdown.pack(side=tk.LEFT)
        ttk.Button(template_frame, text="Details", command=self._select_template).pack(side=tk.LEFT, padx=5)
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options")
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Typing mode
        ttk.Label(options_frame, text="Typing Mode:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.typing_mode_var = tk.StringVar(value="realistic")
        ttk.Combobox(options_frame, textvariable=self.typing_mode_var, values=["fast", "realistic", "slow"], width=10, state="readonly").grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Checkboxes for options
        options_grid = ttk.Frame(options_frame)
        options_grid.grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Create checkboxes with variables
        self.verbose_var = tk.BooleanVar(value=True)
        self.iterative_var = tk.BooleanVar(value=True)
        self.self_improve_var = tk.BooleanVar(value=True)
        self.self_evolve_var = tk.BooleanVar(value=True)
        self.track_position_var = tk.BooleanVar(value=True)
        self.robot_cursor_var = tk.BooleanVar(value=True)
        self.use_autoit_var = tk.BooleanVar(value=True)
        
        checkboxes = [
            ("Verbose", self.verbose_var),
            ("Iterative Mode", self.iterative_var),
            ("Self-Improvement", self.self_improve_var),
            ("Self-Evolution", self.self_evolve_var),
            ("Position Tracking", self.track_position_var),
            ("Robot Cursor", self.robot_cursor_var),
            ("Use AutoIt", self.use_autoit_var)
        ]
        
        # Arrange checkboxes in grid
        for i, (text, var) in enumerate(checkboxes):
            row, col = divmod(i, 3)
            ttk.Checkbutton(options_grid, text=text, variable=var).grid(row=row, column=col, sticky=tk.W, padx=10, pady=2)
        
        # Delay slider
        delay_frame = ttk.Frame(options_frame)
        delay_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        ttk.Label(delay_frame, text="Start Delay:").pack(side=tk.LEFT, padx=(0, 5))
        self.delay_var = tk.DoubleVar(value=3.0)
        ttk.Scale(delay_frame, from_=0, to=10, variable=self.delay_var, orient=tk.HORIZONTAL, length=200).pack(side=tk.LEFT)
        ttk.Label(delay_frame, textvariable=tk.StringVar(value="")).pack(side=tk.LEFT, padx=(5, 0))
        self.delay_var.trace_add("write", lambda *args: delay_frame.children[list(delay_frame.children.keys())[-1]].configure(text=f"{self.delay_var.get():.1f} sec"))
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Run and Stop buttons
        self.run_button = ttk.Button(button_frame, text="Run Agent", command=self._run_agent)
        self.run_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Stop Agent", command=self._stop_agent, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar(value=0.0)
        ttk.Progressbar(button_frame, variable=self.progress_var, length=200, mode="determinate").pack(side=tk.LEFT, padx=5)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(button_frame, textvariable=self.status_var).pack(side=tk.LEFT, padx=5)
        
        # Log display
        log_frame = ttk.LabelFrame(main_frame, text="Log")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, width=80, height=15)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Set up tag for log levels
        self.log_text.tag_configure("INFO", foreground="black")
        self.log_text.tag_configure("DEBUG", foreground="gray")
        self.log_text.tag_configure("WARNING", foreground="orange")
        self.log_text.tag_configure("ERROR", foreground="red")
        self.log_text.tag_configure("CRITICAL", foreground="red", font=("TkDefaultFont", 10, "bold"))
        
        # Make log read-only
        self.log_text.config(state=tk.DISABLED)
    
    def _create_menu(self):
        """Create application menu"""
        menubar = tk.Menu(self)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Load Config", command=self._load_config)
        file_menu.add_command(label="Save Config", command=self._save_config)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_exit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Clear Log", command=self._clear_log)
        tools_menu.add_command(label="Plugin Manager", command=self._open_plugin_manager)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Documentation", command=self._show_documentation)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.config(menu=menubar)
    
    def _setup_log_monitor(self):
        """Setup log monitoring"""
        self.after(100, self._check_log_queue)
    
    def _check_log_queue(self):
        """Check for log messages in the queue"""
        try:
            while True:
                record = self.log_queue.get_nowait()
                self._display_log(record)
                self.log_queue.task_done()
        except queue.Empty:
            pass
        
        self.after(100, self._check_log_queue)
    
    def _display_log(self, record):
        """Display a log message in the log text widget"""
        level, message = record
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n", level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def _select_template(self):
        """Open template selection dialog"""
        def on_template_selected(template):
            """Callback for template selection"""
            if template:
                self.selected_template = template
                self.template_var.set(template.name)
        
        # Create and show dialog
        dialog = TemplateSelectionDialog(self, self.template_manager, on_template_selected)
    
    def _update_config_from_gui(self):
        """Update configuration from GUI values"""
        self.agent_config.api_key = self.api_key_var.get()
        self.agent_config.typing_mode = self.typing_mode_var.get()
        self.agent_config.verbose = self.verbose_var.get()
        self.agent_config.iterative = self.iterative_var.get()
        self.agent_config.self_improve = self.self_improve_var.get()
        self.agent_config.self_evolve = self.self_evolve_var.get()
        self.agent_config.track_position = self.track_position_var.get()
        self.agent_config.robot_cursor = self.robot_cursor_var.get()
        self.agent_config.use_autoit = self.use_autoit_var.get()
        self.agent_config.delay = self.delay_var.get()
    
    def _run_agent(self):
        """Run the agent in a separate thread"""
        if self.is_running:
            return
        
        # Update config from GUI
        self._update_config_from_gui()
        
        # Validate inputs
        if not self.agent_config.api_key:
            messagebox.showerror("Error", "API Key is required")
            return
        
        # Update UI state
        self.run_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_var.set("Running...")
        self.progress_var.set(0.0)
        self.is_running = True
        
        # Start agent thread
        self.agent_thread = threading.Thread(target=self._agent_thread_func)
        self.agent_thread.daemon = True
        self.agent_thread.start()
    
    def _agent_thread_func(self):
        """Function to run in agent thread"""
        try:
            # Log starting
            self.log_queue.put(("INFO", f"Starting agent with topic: {self.topic_var.get()}"))
            
            # Create event loop for asyncio
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
            # Create and run agent
            agent = AgentFactory.create_agent(self.agent_config)
            
            # Run agent
            loop.run_until_complete(self._run_agent_async(agent))
            
            # Update UI when done
            self.status_var.set("Completed")
            self.progress_var.set(100.0)
            
        except Exception as e:
            # Log error
            self.log_queue.put(("ERROR", f"Error running agent: {str(e)}"))
            self.status_var.set("Error")
        finally:
            # Reset UI state
            self.is_running = False
            self.run_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
    
    async def _run_agent_async(self, agent):
        """Run the agent asynchronously"""
        # Initialize agent
        self.log_queue.put(("INFO", "Initializing agent..."))
        await agent.initialize()
        self.progress_var.set(10.0)
        
        # Process text
        text = f"Writing paper on: {self.topic_var.get()}"
        self.log_queue.put(("INFO", f"Processing text: {text}"))
        await agent.process_text(text)
        self.progress_var.set(90.0)
        
        # Finalize agent
        self.log_queue.put(("INFO", "Finalizing agent..."))
        await agent.finalize()
        self.progress_var.set(100.0)
        
        self.log_queue.put(("INFO", "Agent completed successfully"))
    
    def _stop_agent(self):
        """Stop the running agent"""
        if not self.is_running:
            return
        
        # Set flag to stop agent
        self.is_running = False
        self.status_var.set("Stopping...")
        self.log_queue.put(("INFO", "Stopping agent..."))
    
    def _clear_log(self):
        """Clear the log display"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def _load_config(self):
        """Load configuration from file"""
        filename = filedialog.askopenfilename(
            title="Load Configuration",
            filetypes=[("YAML Files", "*.yaml"), ("All Files", "*.*")]
        )
        if not filename:
            return
        
        try:
            self.agent_config = Config.from_yaml(filename)
            self._update_gui_from_config()
            messagebox.showinfo("Success", "Configuration loaded successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Error loading configuration: {str(e)}")
    
    def _save_config(self):
        """Save configuration to file"""
        filename = filedialog.asksaveasfilename(
            title="Save Configuration",
            defaultextension=".yaml",
            filetypes=[("YAML Files", "*.yaml"), ("All Files", "*.*")]
        )
        if not filename:
            return
        
        try:
            self._update_config_from_gui()
            self.agent_config.to_yaml(filename)
            messagebox.showinfo("Success", "Configuration saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving configuration: {str(e)}")
    
    def _update_gui_from_config(self):
        """Update GUI values from configuration"""
        self.api_key_var.set(self.agent_config.api_key)
        self.typing_mode_var.set(self.agent_config.typing_mode)
        self.verbose_var.set(self.agent_config.verbose)
        self.iterative_var.set(self.agent_config.iterative)
        self.self_improve_var.set(self.agent_config.self_improve)
        self.self_evolve_var.set(self.agent_config.self_evolve)
        self.track_position_var.set(self.agent_config.track_position)
        self.robot_cursor_var.set(self.agent_config.robot_cursor)
        self.use_autoit_var.set(self.agent_config.use_autoit)
        self.delay_var.set(self.agent_config.delay)
    
    def _open_plugin_manager(self):
        """Open plugin manager dialog"""
        # This would be implemented in a separate class
        messagebox.showinfo("Plugin Manager", "Plugin Manager not yet implemented")
    
    def _show_documentation(self):
        """Show documentation"""
        # This would open a web browser or help window
        messagebox.showinfo("Documentation", "Documentation not yet implemented")
    
    def _show_about(self):
        """Show about dialog"""
        messagebox.showinfo(
            "About Word AI Agent",
            "Word AI Agent\nVersion 1.0.0\n\nAI-powered document generation tool"
        )
    
    def _on_exit(self):
        """Handle application exit"""
        if self.is_running:
            if not messagebox.askyesno("Exit", "Agent is running. Are you sure you want to exit?"):
                return
        
        self.quit()

def run_gui():
    """Run the GUI application"""
    app = AgentGUI()
    app.mainloop()
