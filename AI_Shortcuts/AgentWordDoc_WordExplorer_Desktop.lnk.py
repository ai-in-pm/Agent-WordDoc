"""
Shortcut creator for Word Interface Explorer GUI

This script creates a shortcut that launches the Word Interface Explorer GUI
in the requested desktop folder.
"""

import os
import sys
import winshell
from pathlib import Path
import pythoncom
from win32com.client import Dispatch

# Get the root directory
root_dir = Path(__file__).parent.parent

# Path to the GUI script
target_script = root_dir / "src" / "bootstrap_training" / "word_interface" / "home_tab" / "gui.py"

# Path where the shortcut will be created
shortcut_dir = Path(r"C:\Users\djjme\OneDrive\Desktop\AI_Shortcuts\AgentWordDoc")
shortcut_path = shortcut_dir / "AgentWordDoc_WordExplorer.lnk"

# Create shortcut
def create_shortcut():
    pythoncom.CoInitialize()
    try:
        # Ensure the target directory exists
        os.makedirs(shortcut_dir, exist_ok=True)
        
        # Get python executable path
        python_exe = sys.executable
        
        # Create shell object
        shell = Dispatch('WScript.Shell')
        
        # Create shortcut
        shortcut = shell.CreateShortCut(str(shortcut_path))
        shortcut.Targetpath = python_exe
        shortcut.Arguments = f'"{target_script}"'
        shortcut.WorkingDirectory = str(root_dir)
        
        # Set icon - use a Word icon if available, otherwise use Python icon
        icon_path = root_dir / "assets" / "icons" / "word_icon.ico"
        if icon_path.exists():
            shortcut.IconLocation = str(icon_path)
        else:
            # Try to use the Microsoft Word icon if installed
            office_path = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
            if os.path.exists(office_path):
                shortcut.IconLocation = f"{office_path},0"
            else:
                # Fallback to Python icon
                shortcut.IconLocation = f"{python_exe},0"
        
        shortcut.Description = "Launch Word Interface Explorer GUI"
        shortcut.save()
        print(f"Shortcut created at {shortcut_path}")
        return True
    except Exception as e:
        print(f"Error creating shortcut: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    success = create_shortcut()
    sys.exit(0 if success else 1)
