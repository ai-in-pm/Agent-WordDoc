"""Template selection dialog for the Word AI Agent GUI"""

import tkinter as tk
from tkinter import ttk
import tkinter.scrolledtext as scrolledtext
from typing import Optional, Callable

from src.templates.template_manager import TemplateManager, PaperTemplate

class TemplateSelectionDialog(tk.Toplevel):
    """Dialog for selecting a paper template"""
    def __init__(self, parent, template_manager: TemplateManager, callback: Callable[[Optional[PaperTemplate]], None]):
        super().__init__(parent)
        
        self.title("Select Paper Template")
        self.geometry("600x500")
        self.minsize(500, 400)
        self.transient(parent)
        self.grab_set()
        
        self.template_manager = template_manager
        self.callback = callback
        self.selected_template = None
        
        self._create_widgets()
        self._center_window()
        self._populate_templates()
    
    def _create_widgets(self):
        """Create dialog widgets"""
        # Create frames
        self.top_frame = ttk.Frame(self)
        self.top_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.content_frame = ttk.Frame(self)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.button_frame = ttk.Frame(self)
        self.button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Template selection dropdown
        ttk.Label(self.top_frame, text="Select Template:").pack(side=tk.LEFT, padx=5)
        self.template_var = tk.StringVar()
        self.template_dropdown = ttk.Combobox(self.top_frame, textvariable=self.template_var, state="readonly", width=40)
        self.template_dropdown.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.template_dropdown.bind("<<ComboboxSelected>>", self._on_template_selected)
        
        # Template details frame
        self.details_frame = ttk.LabelFrame(self.content_frame, text="Template Details")
        self.details_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Description
        ttk.Label(self.details_frame, text="Description:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.description_var = tk.StringVar()
        description_entry = ttk.Entry(self.details_frame, textvariable=self.description_var, state="readonly", width=50)
        description_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        
        # Citation style
        ttk.Label(self.details_frame, text="Citation Style:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.citation_var = tk.StringVar()
        citation_entry = ttk.Entry(self.details_frame, textvariable=self.citation_var, state="readonly", width=20)
        citation_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Formatting
        ttk.Label(self.details_frame, text="Formatting:").grid(row=2, column=0, sticky=tk.NW, padx=5, pady=5)
        self.format_text = scrolledtext.ScrolledText(self.details_frame, width=40, height=4, wrap=tk.WORD)
        self.format_text.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
        self.format_text.config(state=tk.DISABLED)
        
        # Sections
        ttk.Label(self.details_frame, text="Sections:").grid(row=3, column=0, sticky=tk.NW, padx=5, pady=5)
        self.sections_text = scrolledtext.ScrolledText(self.details_frame, width=40, height=10, wrap=tk.WORD)
        self.sections_text.grid(row=3, column=1, sticky=tk.NSEW, padx=5, pady=5)
        self.sections_text.config(state=tk.DISABLED)
        
        # Configure grid weights
        self.details_frame.columnconfigure(1, weight=1)
        self.details_frame.rowconfigure(3, weight=1)
        
        # Buttons
        ttk.Button(self.button_frame, text="Select", command=self._on_select).pack(side=tk.RIGHT, padx=5)
        ttk.Button(self.button_frame, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
    
    def _center_window(self):
        """Center dialog on parent window"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
    
    def _populate_templates(self):
        """Populate template dropdown"""
        template_names = [t.name for t in self.template_manager.templates]
        self.template_dropdown['values'] = template_names
        
        if template_names:
            self.template_dropdown.current(0)
            self._on_template_selected(None)
    
    def _on_template_selected(self, event):
        """Handle template selection"""
        selected_name = self.template_var.get()
        
        # Find selected template
        self.selected_template = next((t for t in self.template_manager.templates if t.name == selected_name), None)
        
        if self.selected_template:
            # Update UI
            self.description_var.set(self.selected_template.description)
            self.citation_var.set(self.selected_template.citation_style)
            
            # Update formatting text
            self.format_text.config(state=tk.NORMAL)
            self.format_text.delete(1.0, tk.END)
            format_text = "\n".join([f"{k}: {v}" for k, v in self.selected_template.formatting.items()])
            self.format_text.insert(tk.END, format_text)
            self.format_text.config(state=tk.DISABLED)
            
            # Update sections text
            self.sections_text.config(state=tk.NORMAL)
            self.sections_text.delete(1.0, tk.END)
            sections_text = "\n".join([f"• {s}" for s in self.selected_template.sections])
            self.sections_text.insert(tk.END, sections_text)
            self.sections_text.config(state=tk.DISABLED)
    
    def _on_select(self):
        """Handle selection confirmation"""
        self.callback(self.selected_template)
        self.destroy()
    
    def _on_cancel(self):
        """Handle dialog cancellation"""
        self.callback(None)
        self.destroy()
