"""
Voice Command Tutorial for Agent WordDoc

This module provides an introduction to voice commands for new users
and helps users learn available commands through interactive guidance.
"""

import asyncio
import random
from typing import List, Dict, Any, Optional, Callable
import time
import threading

from src.core.logging import get_logger
from src.voice.elevenlabs_integration import ElevenLabsIntegration
from src.voice.command_handler import CommandHandler

# Initialize logger
logger = get_logger(__name__)

class VoiceCommandTutorial:
    """Interactive tutorial for voice commands"""
    
    def __init__(self, voice_integration: ElevenLabsIntegration, command_handler: CommandHandler):
        self.voice = voice_integration
        self.commands = command_handler
        self.tutorial_running = False
        self.tutorial_thread = None
        self.command_categories = self._build_command_categories()
        logger.info("Voice command tutorial initialized")
    
    def _build_command_categories(self) -> Dict[str, List[Dict[str, str]]]:
        """Organize commands into categories for the tutorial"""
        all_commands = self.commands.get_commands_list()
        
        # Define categories
        categories = {
            "basic": [],
            "document": [],
            "navigation": [],
            "editing": [],
            "formatting": [],
            "system": [],
        }
        
        # Map commands to categories based on keywords
        category_keywords = {
            "basic": ["start", "stop", "pause", "continue", "resume", "help"],
            "document": ["write about", "add section", "add paragraph", "add introduction", "add conclusion", "add references"],
            "navigation": ["go to", "navigate"],
            "editing": ["delete", "undo", "redo"],
            "formatting": ["format", "bold", "italic", "underlined"],
            "system": ["save", "exit", "quit", "enable", "disable", "set typing"],
        }
        
        # Assign commands to categories
        for cmd in all_commands:
            assigned = False
            for category, keywords in category_keywords.items():
                if any(keyword in cmd["command"].lower() for keyword in keywords):
                    categories[category].append(cmd)
                    assigned = True
                    break
            
            # If not assigned to any category, put in basic
            if not assigned:
                categories["basic"].append(cmd)
        
        return categories
    
    def start_tutorial(self) -> None:
        """Start the voice command tutorial in a separate thread"""
        if self.tutorial_running:
            logger.info("Tutorial already running")
            return
        
        self.tutorial_running = True
        self.tutorial_thread = threading.Thread(target=self._run_tutorial)
        self.tutorial_thread.daemon = True
        self.tutorial_thread.start()
        logger.info("Started voice command tutorial")
    
    def stop_tutorial(self) -> None:
        """Stop the running tutorial"""
        if not self.tutorial_running:
            return
        
        self.tutorial_running = False
        if self.tutorial_thread and self.tutorial_thread.is_alive():
            self.tutorial_thread.join(timeout=2.0)
        logger.info("Stopped voice command tutorial")
    
    def _run_tutorial(self) -> None:
        """Run the tutorial sequence"""
        try:
            # Introduction
            self._speak_with_pause(
                "Welcome to the Agent WordDoc voice command system. "
                "I'll help you learn how to control the agent using your voice. "
                "Let's go through some basic commands you can use.", 
                pause=1.0
            )
            
            # Basic controls
            self._speak_with_pause(
                "First, let's cover basic control commands. "
                "You can say 'start writing' to begin, 'pause' to pause the agent, "
                "and 'continue' or 'resume' to resume operation. "
                "To stop completely, just say 'stop writing'.", 
                pause=0.5
            )
            
            # Document creation
            self._speak_with_pause(
                "To create document content, try commands like: "
                "'write about artificial intelligence', "
                "'add section literature review', "
                "'add introduction', or 'add conclusion'.", 
                pause=0.5
            )
            
            # Navigation
            self._speak_with_pause(
                "To navigate in your document, you can say "
                "'go to top', 'go to bottom', or 'go to section introduction'.", 
                pause=0.5
            )
            
            # Editing
            self._speak_with_pause(
                "For editing, try commands like 'undo', 'redo', "
                "or 'delete last paragraph'.", 
                pause=0.5
            )
            
            # System commands
            self._speak_with_pause(
                "Finally, you can use system commands like 'save document' to save your work, "
                "'set typing speed to fast' to change typing behavior, "
                "or 'exit' to quit the application.", 
                pause=0.5
            )
            
            # Help system
            self._speak_with_pause(
                "If you forget any commands, just say 'help' or 'list commands' "
                "to see what's available. Now, what would you like to do first?", 
                pause=0.0
            )
            
            logger.info("Voice command tutorial completed")
            
        except Exception as e:
            logger.error(f"Error in tutorial: {e}")
        finally:
            self.tutorial_running = False
    
    def _speak_with_pause(self, text: str, pause: float = 0.5) -> None:
        """Speak text with a pause afterward"""
        if not self.tutorial_running:
            return
        
        success = self.voice.speak(text)
        if success and pause > 0 and self.tutorial_running:
            time.sleep(pause)
    
    def provide_command_help(self, category: Optional[str] = None) -> None:
        """Provide specific help for a command category"""
        if category and category in self.command_categories:
            commands = self.command_categories[category]
            category_name = category.title()
            
            if not commands:
                self.voice.speak(f"I don't have any {category} commands to show you.")
                return
            
            # Prepare speech text
            speech = f"Here are {category_name} commands you can use:\n"
            
            # Add example commands (limit to 5 for voice)
            for i, cmd in enumerate(commands[:5]):
                speech += f"{i+1}. {cmd['command']}: {cmd['description']}\n"
            
            self.voice.speak(speech)
            
        else:
            # If no category specified, list categories
            self.voice.speak(
                "I can help you with different types of commands. "
                "Say 'help with basic commands', 'help with document commands', "
                "'help with navigation commands', 'help with editing commands', "
                "or 'help with system commands' to learn more."
            )
    
    def ask_question(self, question_type: str) -> None:
        """Ask user a question to guide them through commands"""
        questions = {
            "getting_started": [
                "Would you like me to help you get started with voice commands?",
                "Do you want to learn how to control the agent with your voice?",
                "Should I show you how to use voice commands?"
            ],
            "document_creation": [
                "What topic would you like to write about?",
                "Would you like to add an introduction to your document?",
                "Should we add some sections to organize your document?"
            ],
            "navigation": [
                "Would you like to go to a specific part of the document?",
                "Do you want to move to the beginning or end of the document?",
                "Should I help you navigate through the document?"
            ],
            "editing": [
                "Would you like to make changes to what's already written?",
                "Do you need to undo or redo any actions?",
                "Should we edit any of the existing content?"
            ],
            "next_step": [
                "What would you like to do next?",
                "Is there anything specific you'd like me to help with?",
                "What command should we try now?"
            ],
        }
        
        if question_type in questions:
            question = random.choice(questions[question_type])
            self.voice.speak(question)
        else:
            self.voice.speak("What would you like to do next?")
