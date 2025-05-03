"""
AutoIt Integration for AI Agent

This module provides integration with AutoIt for advanced Windows automation capabilities.
It allows the AI Agent to perform complex interactions with Microsoft Word and other applications.
"""

import os
import sys
import time
import tempfile
import subprocess
import random
from typing import Dict, List, Any, Optional, Union, Tuple

# Check if AutoIt is installed
AUTOIT_INSTALLED = False
try:
    # Try to import the AutoItX Python library if available
    import autoit
    AUTOIT_INSTALLED = True
except ImportError:
    # Check if the AutoIt executable is available
    try:
        autoit_path = r"C:\Program Files (x86)\AutoIt3\AutoIt3.exe"
        if os.path.exists(autoit_path):
            AUTOIT_INSTALLED = True
    except:
        pass

class AutoItIntegration:
    """Provides integration with AutoIt for advanced Windows automation"""
    
    def __init__(self, verbose: bool = False):
        """
        Initialize the AutoIt integration
        
        Args:
            verbose: Whether to print debug information
        """
        self.verbose = verbose
        self.script_dir = os.path.join(tempfile.gettempdir(), "ai_agent_autoit_scripts")
        
        # Create script directory if it doesn't exist
        if not os.path.exists(self.script_dir):
            os.makedirs(self.script_dir)
        
        # Check if AutoIt is installed
        if not AUTOIT_INSTALLED:
            if self.verbose:
                print("[AUTOIT] AutoIt is not installed. Some features may not work.")
        else:
            if self.verbose:
                print("[AUTOIT] AutoIt integration initialized")
    
    def create_script(self, script_content: str, script_name: Optional[str] = None) -> str:
        """
        Create an AutoIt script file
        
        Args:
            script_content: The content of the AutoIt script
            script_name: Optional name for the script file
            
        Returns:
            Path to the created script file
        """
        try:
            # Generate a script name if not provided
            if not script_name:
                timestamp = int(time.time())
                random_suffix = ''.join(random.choices('abcdefghijklmnopqrstuvwxyz', k=5))
                script_name = f"ai_agent_script_{timestamp}_{random_suffix}.au3"
            
            # Ensure the script has the .au3 extension
            if not script_name.endswith('.au3'):
                script_name += '.au3'
            
            # Create the script file
            script_path = os.path.join(self.script_dir, script_name)
            with open(script_path, 'w') as f:
                f.write(script_content)
            
            if self.verbose:
                print(f"[AUTOIT] Created script: {script_path}")
            
            return script_path
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to create AutoIt script: {str(e)}")
            return ""
    
    def run_script(self, script_path: str, wait: bool = True) -> bool:
        """
        Run an AutoIt script
        
        Args:
            script_path: Path to the AutoIt script
            wait: Whether to wait for the script to complete
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if not AUTOIT_INSTALLED:
                if self.verbose:
                    print("[ERROR] AutoIt is not installed")
                return False
            
            # Check if the script exists
            if not os.path.exists(script_path):
                if self.verbose:
                    print(f"[ERROR] Script not found: {script_path}")
                return False
            
            # Run the script
            autoit_exe = r"C:\Program Files (x86)\AutoIt3\AutoIt3.exe"
            
            if wait:
                # Run and wait for completion
                result = subprocess.run([autoit_exe, script_path], capture_output=True, text=True)
                
                if self.verbose:
                    print(f"[AUTOIT] Script executed: {script_path}")
                    if result.stdout:
                        print(f"[AUTOIT] Output: {result.stdout}")
                    if result.stderr:
                        print(f"[AUTOIT] Error: {result.stderr}")
                
                return result.returncode == 0
            else:
                # Run without waiting
                subprocess.Popen([autoit_exe, script_path])
                
                if self.verbose:
                    print(f"[AUTOIT] Script launched: {script_path}")
                
                return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to run AutoIt script: {str(e)}")
            return False
    
    def create_and_run_script(self, script_content: str, wait: bool = True) -> bool:
        """
        Create and run an AutoIt script
        
        Args:
            script_content: The content of the AutoIt script
            wait: Whether to wait for the script to complete
            
        Returns:
            True if successful, False otherwise
        """
        script_path = self.create_script(script_content)
        if script_path:
            return self.run_script(script_path, wait)
        return False
    
    def word_type_text(self, text: str, typing_speed: str = "realistic") -> bool:
        """
        Type text in Microsoft Word using AutoIt
        
        Args:
            text: Text to type
            typing_speed: Typing speed ('fast', 'realistic', 'slow')
            
        Returns:
            True if successful, False otherwise
        """
        # Define typing delays based on speed
        if typing_speed == "fast":
            delay_range = (10, 30)  # milliseconds
        elif typing_speed == "slow":
            delay_range = (100, 300)  # milliseconds
        else:  # realistic
            delay_range = (50, 150)  # milliseconds
        
        # Create AutoIt script for typing
        script_content = f"""
        ; AutoIt script for typing text in Microsoft Word
        #include <MsgBoxConstants.au3>
        
        ; Activate Microsoft Word
        If Not WinActive("Microsoft Word") Then
            WinActivate("Microsoft Word")
            If @error Then
                MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                Exit 1
            EndIf
            Sleep(500)
        EndIf
        
        ; Type the text with realistic timing
        Local $text = "{text}"
        Local $chars = StringSplit($text, "")
        
        For $i = 1 To $chars[0]
            ; Type the character
            Send($chars[$i])
            
            ; Random delay between keystrokes
            Sleep(Random({delay_range[0]}, {delay_range[1]}))
            
            ; Occasionally make a typing error and correct it (5% chance)
            If Random(1, 100) <= 5 Then
                ; Type a random wrong character
                Send(Chr(Random(97, 122)))
                Sleep(50)
                ; Delete it
                Send("{{BACKSPACE}}")
                Sleep(20)
                ; Retype the correct character
                Send($chars[$i])
                Sleep(Random({delay_range[0]}, {delay_range[1]}))
            EndIf
            
            ; Add longer pause after punctuation
            If StringInStr(".,:;!?", $chars[$i]) > 0 Then
                Sleep(Random(100, 300))
            EndIf
            
            ; Add pause after space (end of word)
            If $chars[$i] = " " Then
                Sleep(Random(80, 200))
            EndIf
        Next
        
        Exit 0
        """
        
        return self.create_and_run_script(script_content)
    
    def word_format_text(self, format_type: str, start_select: bool = True) -> bool:
        """
        Format selected text in Microsoft Word
        
        Args:
            format_type: Type of formatting ('bold', 'italic', 'underline', 'heading1', etc.)
            start_select: Whether to start by selecting text (if False, assumes text is already selected)
            
        Returns:
            True if successful, False otherwise
        """
        # Define formatting commands
        format_commands = {
            "bold": "^b",  # Ctrl+B
            "italic": "^i",  # Ctrl+I
            "underline": "^u",  # Ctrl+U
            "heading1": "^+1",  # Ctrl+Alt+1
            "heading2": "^+2",  # Ctrl+Alt+2
            "heading3": "^+3",  # Ctrl+Alt+3
            "center": "^e",  # Ctrl+E
            "left": "^l",  # Ctrl+L
            "right": "^r",  # Ctrl+R
            "justify": "^j",  # Ctrl+J
            "bullet": "^+l",  # Ctrl+Shift+L
            "numbering": "^+n",  # Ctrl+Shift+N
        }
        
        # Get the formatting command
        command = format_commands.get(format_type.lower())
        if not command:
            if self.verbose:
                print(f"[ERROR] Unknown formatting type: {format_type}")
            return False
        
        # Create AutoIt script for formatting
        script_content = f"""
        ; AutoIt script for formatting text in Microsoft Word
        #include <MsgBoxConstants.au3>
        
        ; Activate Microsoft Word
        If Not WinActive("Microsoft Word") Then
            WinActivate("Microsoft Word")
            If @error Then
                MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                Exit 1
            EndIf
            Sleep(500)
        EndIf
        
        ; Select text if needed
        {"If Not WinActive(\"Microsoft Word\") Then" if start_select else "; No selection needed"}
            {"Send(\"^a\")" if start_select else "; Skip selection"}
            {"Sleep(200)" if start_select else ""}
        {"EndIf" if start_select else ""}
        
        ; Apply formatting
        Send("{command}")
        Sleep(200)
        
        ; Deselect text
        Send("{{RIGHT}}")
        
        Exit 0
        """
        
        return self.create_and_run_script(script_content)
    
    def word_insert_table(self, rows: int, columns: int) -> bool:
        """
        Insert a table in Microsoft Word
        
        Args:
            rows: Number of rows
            columns: Number of columns
            
        Returns:
            True if successful, False otherwise
        """
        # Create AutoIt script for inserting a table
        script_content = f"""
        ; AutoIt script for inserting a table in Microsoft Word
        #include <MsgBoxConstants.au3>
        
        ; Activate Microsoft Word
        If Not WinActive("Microsoft Word") Then
            WinActivate("Microsoft Word")
            If @error Then
                MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                Exit 1
            EndIf
            Sleep(500)
        EndIf
        
        ; Open Insert Table dialog
        Send("!n")  ; Alt+N for Insert tab
        Sleep(200)
        Send("t")  ; T for Table
        Sleep(200)
        
        ; Enter table dimensions
        Send("{rows}")
        Sleep(100)
        Send("{{TAB}}")
        Sleep(100)
        Send("{columns}")
        Sleep(100)
        Send("{{ENTER}}")
        
        Exit 0
        """
        
        return self.create_and_run_script(script_content)
    
    def word_navigate(self, location: str) -> bool:
        """
        Navigate to a specific location in Microsoft Word
        
        Args:
            location: Where to navigate ('start', 'end', 'next_page', 'previous_page')
            
        Returns:
            True if successful, False otherwise
        """
        # Define navigation commands
        navigation_commands = {
            "start": "^{HOME}",  # Ctrl+Home
            "end": "^{END}",  # Ctrl+End
            "next_page": "^{PGDN}",  # Ctrl+Page Down
            "previous_page": "^{PGUP}",  # Ctrl+Page Up
        }
        
        # Get the navigation command
        command = navigation_commands.get(location.lower())
        if not command:
            if self.verbose:
                print(f"[ERROR] Unknown navigation location: {location}")
            return False
        
        # Create AutoIt script for navigation
        script_content = f"""
        ; AutoIt script for navigating in Microsoft Word
        #include <MsgBoxConstants.au3>
        
        ; Activate Microsoft Word
        If Not WinActive("Microsoft Word") Then
            WinActivate("Microsoft Word")
            If @error Then
                MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                Exit 1
            EndIf
            Sleep(500)
        EndIf
        
        ; Navigate
        Send("{command}")
        Sleep(200)
        
        Exit 0
        """
        
        return self.create_and_run_script(script_content)
    
    def word_save_document(self, file_path: Optional[str] = None) -> bool:
        """
        Save the current Microsoft Word document
        
        Args:
            file_path: Path to save the document (if None, uses Save, otherwise uses Save As)
            
        Returns:
            True if successful, False otherwise
        """
        # Create AutoIt script for saving
        if file_path:
            script_content = f"""
            ; AutoIt script for saving a Word document with a specific path
            #include <MsgBoxConstants.au3>
            
            ; Activate Microsoft Word
            If Not WinActive("Microsoft Word") Then
                WinActivate("Microsoft Word")
                If @error Then
                    MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                    Exit 1
                EndIf
                Sleep(500)
            EndIf
            
            ; Save As
            Send("^+s")  ; Ctrl+Shift+S for Save As
            Sleep(1000)
            
            ; Enter file path
            Send("{file_path}")
            Sleep(500)
            Send("{{ENTER}}")
            
            ; Handle potential overwrite confirmation
            If WinExists("Confirm Save As") Then
                Send("y")  ; Yes to overwrite
                Sleep(500)
            EndIf
            
            Exit 0
            """
        else:
            script_content = """
            ; AutoIt script for saving a Word document
            #include <MsgBoxConstants.au3>
            
            ; Activate Microsoft Word
            If Not WinActive("Microsoft Word") Then
                WinActivate("Microsoft Word")
                If @error Then
                    MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                    Exit 1
                EndIf
                Sleep(500)
            EndIf
            
            ; Save
            Send("^s")  ; Ctrl+S for Save
            Sleep(500)
            
            ; Handle potential save dialog if it's a new document
            If WinExists("Save As") Then
                Send("!s")  ; Alt+S for Save
                Sleep(500)
            EndIf
            
            Exit 0
            """
        
        return self.create_and_run_script(script_content)
    
    def word_find_text(self, text: str, select: bool = True) -> bool:
        """
        Find text in Microsoft Word
        
        Args:
            text: Text to find
            select: Whether to select the found text
            
        Returns:
            True if successful, False otherwise
        """
        # Create AutoIt script for finding text
        script_content = f"""
        ; AutoIt script for finding text in Microsoft Word
        #include <MsgBoxConstants.au3>
        
        ; Activate Microsoft Word
        If Not WinActive("Microsoft Word") Then
            WinActivate("Microsoft Word")
            If @error Then
                MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                Exit 1
            EndIf
            Sleep(500)
        EndIf
        
        ; Open Find dialog
        Send("^f")  ; Ctrl+F for Find
        Sleep(500)
        
        ; Enter text to find
        Send("{text}")
        Sleep(200)
        
        ; Find
        Send("{{ENTER}}")
        
        ; Close Find dialog if not selecting
        {"" if select else "Send(\"{{ESC}}\")\nSleep(200)"}
        
        Exit 0
        """
        
        return self.create_and_run_script(script_content)
    
    def word_replace_text(self, find_text: str, replace_text: str, replace_all: bool = False) -> bool:
        """
        Replace text in Microsoft Word
        
        Args:
            find_text: Text to find
            replace_text: Text to replace with
            replace_all: Whether to replace all occurrences
            
        Returns:
            True if successful, False otherwise
        """
        # Create AutoIt script for replacing text
        script_content = f"""
        ; AutoIt script for replacing text in Microsoft Word
        #include <MsgBoxConstants.au3>
        
        ; Activate Microsoft Word
        If Not WinActive("Microsoft Word") Then
            WinActivate("Microsoft Word")
            If @error Then
                MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                Exit 1
            EndIf
            Sleep(500)
        EndIf
        
        ; Open Replace dialog
        Send("^h")  ; Ctrl+H for Replace
        Sleep(500)
        
        ; Enter text to find
        Send("{find_text}")
        Sleep(200)
        Send("{{TAB}}")
        Sleep(200)
        
        ; Enter replacement text
        Send("{replace_text}")
        Sleep(200)
        
        ; Replace or Replace All
        {"Send(\"!a\")" if replace_all else "Send(\"!r\")"}  ; {"Alt+A for Replace All" if replace_all else "Alt+R for Replace"}
        Sleep(500)
        
        ; Close dialog when done
        Send("{{ESC}}")
        
        Exit 0
        """
        
        return self.create_and_run_script(script_content)
    
    def word_insert_image(self, image_path: str) -> bool:
        """
        Insert an image in Microsoft Word
        
        Args:
            image_path: Path to the image file
            
        Returns:
            True if successful, False otherwise
        """
        # Create AutoIt script for inserting an image
        script_content = f"""
        ; AutoIt script for inserting an image in Microsoft Word
        #include <MsgBoxConstants.au3>
        
        ; Activate Microsoft Word
        If Not WinActive("Microsoft Word") Then
            WinActivate("Microsoft Word")
            If @error Then
                MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                Exit 1
            EndIf
            Sleep(500)
        EndIf
        
        ; Open Insert Picture dialog
        Send("!n")  ; Alt+N for Insert tab
        Sleep(200)
        Send("p")  ; P for Picture
        Sleep(200)
        Send("f")  ; F for From File
        Sleep(1000)
        
        ; Enter image path
        Send("{image_path}")
        Sleep(500)
        Send("{{ENTER}}")
        
        Exit 0
        """
        
        return self.create_and_run_script(script_content)
    
    def word_create_new_document(self) -> bool:
        """
        Create a new Microsoft Word document
        
        Returns:
            True if successful, False otherwise
        """
        # Create AutoIt script for creating a new document
        script_content = """
        ; AutoIt script for creating a new Word document
        #include <MsgBoxConstants.au3>
        
        ; Check if Word is running
        If Not ProcessExists("WINWORD.EXE") Then
            ; Start Word
            Run("winword.exe")
            Sleep(2000)
        Else
            ; Activate Word
            WinActivate("Microsoft Word")
            Sleep(500)
        EndIf
        
        ; Create new document
        Send("^n")  ; Ctrl+N for New
        Sleep(1000)
        
        Exit 0
        """
        
        return self.create_and_run_script(script_content)
    
    def word_close_document(self, save: bool = True) -> bool:
        """
        Close the current Microsoft Word document
        
        Args:
            save: Whether to save before closing
            
        Returns:
            True if successful, False otherwise
        """
        # Create AutoIt script for closing a document
        script_content = f"""
        ; AutoIt script for closing a Word document
        #include <MsgBoxConstants.au3>
        
        ; Activate Microsoft Word
        If Not WinActive("Microsoft Word") Then
            WinActivate("Microsoft Word")
            If @error Then
                MsgBox($MB_ICONERROR, "Error", "Could not activate Microsoft Word")
                Exit 1
            EndIf
            Sleep(500)
        EndIf
        
        ; Save if requested
        {"Send(\"^s\")\nSleep(500)\n\n; Handle potential save dialog if it's a new document\nIf WinExists(\"Save As\") Then\n    Send(\"!s\")  ; Alt+S for Save\n    Sleep(500)\nEndIf" if save else "; Skipping save"}
        
        ; Close document
        Send("^w")  ; Ctrl+W for Close
        Sleep(500)
        
        ; Handle save prompt if it appears
        If WinExists("Microsoft Word", "Do you want to save") Then
            {"Send(\"!s\")" if save else "Send(\"!n\")"}  ; {"Alt+S for Save" if save else "Alt+N for Don't Save"}
            Sleep(500)
        EndIf
        
        Exit 0
        """
        
        return self.create_and_run_script(script_content)
    
    def custom_autoit_script(self, script_content: str, wait: bool = True) -> bool:
        """
        Run a custom AutoIt script
        
        Args:
            script_content: The content of the AutoIt script
            wait: Whether to wait for the script to complete
            
        Returns:
            True if successful, False otherwise
        """
        return self.create_and_run_script(script_content, wait)
