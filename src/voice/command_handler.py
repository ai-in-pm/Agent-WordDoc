"""
Voice Command Handler for Agent WordDoc

This module processes voice commands and maps them to agent actions.
"""

import os
import re
import asyncio
from typing import Dict, Any, Callable, List, Optional, Union
import logging

from src.core.logging import get_logger

# Initialize logger
logger = get_logger(__name__)

class CommandHandler:
    """Handle voice commands for the Word AI Agent"""
    
    def __init__(self):
        self.commands = {}
        self._setup_default_commands()
        logger.info("Voice command handler initialized")
    
    def _setup_default_commands(self) -> None:
        """Set up default command patterns and handlers"""
        # Format: (command_pattern, handler_function, command_description)
        default_commands = [
            # Basic control commands
            (r'start( writing)?', self._handle_start, "Start writing"),
            (r'stop( writing)?', self._handle_stop, "Stop writing"),
            (r'stop (the )?agent( from type| from typing| typing)? now', self._handle_emergency_stop, "Immediately stop the agent from typing"),
            (r'pause', self._handle_pause, "Pause the agent"),
            (r'continue|resume', self._handle_continue, "Continue writing"),
            
            # Topic-related commands
            (r'write about (.*)', self._handle_write_about, "Write about a specific topic"),
            (r'change topic to (.*)', self._handle_change_topic, "Change the current topic"),
            
            # Document structure commands
            (r'add (section|paragraph) (.*)', self._handle_add_section, "Add a new section or paragraph"),
            (r'add (introduction|conclusion)', self._handle_add_standard_section, "Add standard section"),
            (r'add (references|bibliography)', self._handle_add_references, "Add references section"),
            
            # Formatting commands
            (r'format as (.*)', self._handle_format, "Apply formatting"),
            (r'make (.*) (bold|italic|underlined)', self._handle_text_style, "Style text"),
            
            # Navigation commands
            (r'go to (top|bottom|start|end)', self._handle_navigation, "Navigate in document"),
            (r'go to section (.*)', self._handle_goto_section, "Go to specific section"),
            
            # Edit commands
            (r'delete (last|previous) (sentence|paragraph|section)', self._handle_delete, "Delete content"),
            (r'undo( last action)?', self._handle_undo, "Undo last action"),
            (r'redo', self._handle_redo, "Redo last undone action"),
            
            # Agent control commands
            (r'set (typing|writing) (speed|mode) to (.*)', self._handle_typing_speed, "Change typing speed"),
            (r'enable (.*)', self._handle_enable_feature, "Enable a feature"),
            (r'disable (.*)', self._handle_disable_feature, "Disable a feature"),
            
            # Help and meta commands
            (r'help( me)?', self._handle_help, "Show help information"),
            (r'list commands', self._handle_list_commands, "List available commands"),
            (r'explain (.*)', self._handle_explain, "Explain a feature or command"),
            
            # System commands
            (r'save( document)?', self._handle_save, "Save current document"),
            (r'exit|quit', self._handle_exit, "Exit the application"),
        ]
        
        # Add commands to the registry
        for pattern, handler, description in default_commands:
            self.register_command(pattern, handler, description)
    
    def register_command(self, pattern: str, handler: Callable, description: str = "") -> None:
        """Register a new command pattern and handler"""
        try:
            # Validate pattern by compiling it
            re.compile(pattern)
            self.commands[pattern] = {
                'handler': handler,
                'description': description
            }
            logger.debug(f"Registered command pattern: {pattern}")
        except re.error as e:
            logger.error(f"Invalid command pattern '{pattern}': {e}")
    
    def process_command(self, command_text: str, agent=None, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Process a voice command and execute appropriate action"""
        if context is None:
            context = {}
        
        command_text = command_text.lower().strip()
        logger.info(f"Processing command: {command_text}")
        
        # Store the original command text
        result = {
            'command': command_text,
            'matched': False,
            'response': '',
            'action': None,
            'success': False,
            'params': {}
        }
        
        # Try to match against registered patterns
        for pattern, command_info in self.commands.items():
            match = re.match(f'^{pattern}$', command_text, re.IGNORECASE)
            if match:
                logger.debug(f"Command matched pattern: {pattern}")
                result['matched'] = True
                handler = command_info['handler']
                
                # Extract any captured groups as parameters
                params = match.groups() if match.groups() else []
                result['params'] = {'match_groups': params, 'agent': agent, **context}
                
                try:
                    # Execute the command handler
                    handler_result = handler(*params, agent=agent, **context)
                    result.update(handler_result if isinstance(handler_result, dict) else 
                                  {'response': str(handler_result) if handler_result else 'Command executed',
                                   'success': True})
                except Exception as e:
                    logger.error(f"Error executing command handler: {e}")
                    result['response'] = f"Error executing command: {str(e)}"
                    result['success'] = False
                
                break
        
        # If no command matched
        if not result['matched']:
            similar_commands = self._find_similar_commands(command_text)
            suggestions = ", ".join(similar_commands[:3]) if similar_commands else ""
            
            if suggestions:
                result['response'] = f"I didn't understand that command. Did you mean: {suggestions}?"
            else:
                result['response'] = "Sorry, I didn't understand that command. Try saying 'help' for a list of commands."
        
        logger.info(f"Command result: {result['success']} - {result['response'][:50]}")
        return result
    
    def _find_similar_commands(self, command_text: str) -> List[str]:
        """Find similar commands to the given text for suggestions"""
        # Simple similarity-based suggestion algorithm
        similar = []
        for pattern, info in self.commands.items():
            # Convert regex pattern to a readable example
            example = self._pattern_to_example(pattern, info['description'])
            # Very basic similarity check - can be improved
            if any(word in example.lower() for word in command_text.lower().split()):
                similar.append(example)
        return similar
    
    def _pattern_to_example(self, pattern: str, description: str) -> str:
        """Convert a regex pattern to a readable example"""
        # This is a very simple conversion - could be improved
        example = pattern
        example = re.sub(r'\((.*?)\)', r'\1', example)  # Remove capturing groups
        example = re.sub(r'\|(.*?)\)', r')', example)    # Remove alternatives in groups
        example = re.sub(r'\|', ' or ', example)         # Convert remaining | to 'or'
        example = re.sub(r'\?', '', example)             # Remove optional markers
        example = re.sub(r'\\', '', example)             # Remove escapes
        example = re.sub(r'\(|\)', '', example)           # Remove any remaining parentheses
        
        return example.strip()
    
    def get_commands_list(self) -> List[Dict[str, str]]:
        """Get a list of available commands with descriptions"""
        commands_list = []
        for pattern, info in self.commands.items():
            example = self._pattern_to_example(pattern, info['description'])
            commands_list.append({
                'command': example,
                'description': info['description']
            })
        return commands_list
    
    # Command handlers - would be connected to the agent in a real implementation
    def _handle_start(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle start command"""
        return {'response': "Starting to write now.", 'action': 'start', 'success': True}
    
    def _handle_stop(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle stop command"""
        return {'response': "Stopping writing.", 'action': 'stop', 'success': True}
    
    def _handle_emergency_stop(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle emergency stop command - immediately stop typing"""
        if agent and hasattr(agent, 'state'):
            # Force typing state to false
            agent.state.is_typing = False
            # Try to stop any active typing
            if hasattr(agent, 'typing_simulator'):
                if hasattr(agent.typing_simulator, 'stop_typing'):
                    agent.typing_simulator.stop_typing()
            
            # If using Word, try to cancel current actions
            if agent.state.using_word and hasattr(agent, 'word_automation'):
                try:
                    # Send escape key to cancel any operations
                    agent.word_automation.send_keys('{ESC}')
                except:
                    pass
        
        return {'response': "Emergency stop activated. Stopping all typing immediately.", 'action': 'emergency_stop', 'success': True}
    
    def _handle_pause(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle pause command"""
        return {'response': "Pausing writing.", 'action': 'pause', 'success': True}
    
    def _handle_continue(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle continue command"""
        return {'response': "Continuing to write.", 'action': 'continue', 'success': True}
    
    def _handle_write_about(self, topic, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle write about command"""
        return {
            'response': f"I'll write about {topic}.", 
            'action': 'write_about', 
            'success': True,
            'topic': topic
        }
    
    def _handle_change_topic(self, topic, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle change topic command"""
        return {
            'response': f"Changing topic to {topic}.", 
            'action': 'change_topic', 
            'success': True,
            'topic': topic
        }
    
    def _handle_add_section(self, section_type, section_name, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle add section command"""
        return {
            'response': f"Adding {section_type} '{section_name}'.", 
            'action': 'add_section', 
            'success': True,
            'section_type': section_type,
            'section_name': section_name
        }
    
    def _handle_add_standard_section(self, section_name, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle add standard section command"""
        return {
            'response': f"Adding {section_name} section.", 
            'action': 'add_standard_section', 
            'success': True,
            'section_name': section_name
        }
    
    def _handle_add_references(self, section_name, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle add references command"""
        return {
            'response': f"Adding {section_name} section.", 
            'action': 'add_references', 
            'success': True,
            'section_name': section_name
        }
    
    def _handle_format(self, format_type, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle format command"""
        return {
            'response': f"Formatting as {format_type}.", 
            'action': 'format', 
            'success': True,
            'format_type': format_type
        }
    
    def _handle_text_style(self, text, style, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle text style command"""
        return {
            'response': f"Making '{text}' {style}.", 
            'action': 'text_style', 
            'success': True,
            'text': text,
            'style': style
        }
    
    def _handle_navigation(self, position, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle navigation command"""
        return {
            'response': f"Navigating to {position}.", 
            'action': 'navigation', 
            'success': True,
            'position': position
        }
    
    def _handle_goto_section(self, section_name, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle go to section command"""
        return {
            'response': f"Going to section '{section_name}'.", 
            'action': 'goto_section', 
            'success': True,
            'section_name': section_name
        }
    
    def _handle_delete(self, target_position, target_type, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle delete command"""
        return {
            'response': f"Deleting {target_position} {target_type}.", 
            'action': 'delete', 
            'success': True,
            'target_position': target_position,
            'target_type': target_type
        }
    
    def _handle_undo(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle undo command"""
        return {'response': "Undoing last action.", 'action': 'undo', 'success': True}
    
    def _handle_redo(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle redo command"""
        return {'response': "Redoing last undone action.", 'action': 'redo', 'success': True}
    
    def _handle_typing_speed(self, mode_type, change_type, speed, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle typing speed command"""
        return {
            'response': f"Setting {mode_type} {change_type} to {speed}.", 
            'action': 'typing_speed', 
            'success': True,
            'speed': speed
        }
    
    def _handle_enable_feature(self, feature, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle enable feature command"""
        return {
            'response': f"Enabling {feature}.", 
            'action': 'enable_feature', 
            'success': True,
            'feature': feature
        }
    
    def _handle_disable_feature(self, feature, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle disable feature command"""
        return {
            'response': f"Disabling {feature}.", 
            'action': 'disable_feature', 
            'success': True,
            'feature': feature
        }
    
    def _handle_help(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle help command"""
        commands = self.get_commands_list()
        help_text = "Available voice commands:\n\n"
        for cmd in commands[:10]:  # Limit to top 10 for voice response
            help_text += f"• {cmd['command']}: {cmd['description']}\n"
        help_text += "\nSay 'list commands' for more options."
        
        return {
            'response': help_text, 
            'action': 'help', 
            'success': True,
            'commands': commands
        }
    
    def _handle_list_commands(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle list commands command"""
        commands = self.get_commands_list()
        cmd_list = "All available voice commands:\n\n"
        for cmd in commands:
            cmd_list += f"• {cmd['command']}: {cmd['description']}\n"
        
        return {
            'response': cmd_list, 
            'action': 'list_commands', 
            'success': True,
            'commands': commands
        }
    
    def _handle_explain(self, feature, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle explain command"""
        explanations = {
            "typing speed": "You can set typing speed to fast, realistic, or slow.",
            "self improve": "Self-improvement allows the agent to learn from mistakes and perform better over time.",
            "self evolve": "Self-evolution enables the agent to develop new capabilities based on your usage patterns.",
            "templates": "Templates provide pre-defined structures for various document types.",
            "word automation": "Word automation controls Microsoft Word directly to create documents.",
        }
        
        explanation = explanations.get(
            feature.lower(), 
            f"I don't have specific information about '{feature}'. Try asking about typing speed, self improve, or templates."
        )
        
        return {
            'response': explanation, 
            'action': 'explain', 
            'success': True,
            'feature': feature
        }
    
    def _handle_save(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle save command"""
        return {'response': "Saving the document.", 'action': 'save', 'success': True}
    
    def _handle_exit(self, *args, agent=None, **kwargs) -> Dict[str, Any]:
        """Handle exit command"""
        return {'response': "Exiting the application.", 'action': 'exit', 'success': True}
