"""
Main Word AI Agent implementation with enhanced capabilities
"""

import time
import random
from typing import Optional, Dict, Any
import asyncio
from dataclasses import dataclass
from abc import ABC, abstractmethod

from src.core.logging import get_logger
from src.config.config import Config
from src.utils.typing_simulator import TypingSimulator
from src.utils.error_handler import ErrorHandler
from src.templates.template_manager import TemplateManager, PaperTemplate

# Import voice capabilities
try:
    from src.voice.elevenlabs_integration import ElevenLabsIntegration
    from src.voice.command_handler import CommandHandler
    from src.voice.command_tutorial import VoiceCommandTutorial
    HAS_VOICE_CAPABILITIES = True
except ImportError:
    HAS_VOICE_CAPABILITIES = False

try:
    from src.utils.word_automation import WordAutomation
    HAS_WORD_AUTOMATION = True
except ImportError:
    HAS_WORD_AUTOMATION = False

logger = get_logger(__name__)

@dataclass
class AgentState:
    """Current state of the agent"""
    is_typing: bool = False
    current_position: int = 0
    last_action: str = ''
    error_count: int = 0
    retry_count: int = 0
    using_word: bool = False
    word_doc_path: Optional[str] = None
    voice_enabled: bool = False

class BaseAgent(ABC):
    """Base class for all agents"""
    
    def __init__(self, config: Config):
        self.config = config
        self.state = AgentState()
        self.error_handler = ErrorHandler()
        self.typing_simulator = TypingSimulator(config.typing_mode)
        
    @abstractmethod
    async def initialize(self) -> None:
        """Initialize the agent"""
        pass
    
    @abstractmethod
    async def process_text(self, text: str, template_id: Optional[str] = None) -> None:
        """Process and type text"""
        pass
    
    @abstractmethod
    async def finalize(self) -> None:
        """Finalize the agent's work"""
        pass

class WordAIAgent(BaseAgent):
    """Main Word AI Agent implementation"""
    
    def __init__(self, config: Config):
        super().__init__(config)
        self._validate_config()
        self._setup_typing_patterns()
        
        # Initialize Word automation if available
        self.word_automation = None
        if config.use_word_automation and HAS_WORD_AUTOMATION:
            try:
                self.word_automation = WordAutomation(
                    use_robot_cursor=config.robot_cursor,
                    cursor_size=config.cursor_size
                )
                logger.info("Word automation initialized successfully")
            except Exception as e:
                logger.error(f"Failed to initialize Word automation: {str(e)}")
        
        # Initialize voice capabilities if enabled
        self.voice_integration = None
        self.command_handler = None
        self.voice_tutorial = None
        if config.voice_command_enabled and HAS_VOICE_CAPABILITIES:
            try:
                self.voice_integration = ElevenLabsIntegration()
                self.command_handler = CommandHandler()
                self.voice_tutorial = VoiceCommandTutorial(self.voice_integration, self.command_handler)
                self.state.voice_enabled = True
                logger.info("Voice command capabilities initialized successfully")
            except Exception as e:
                logger.error(f"Failed to initialize voice capabilities: {str(e)}")
                self.state.voice_enabled = False
        
    def _validate_config(self) -> None:
        """Validate agent configuration"""
        if not self.config.api_key:
            raise ValueError("API key is required")
        if not 0 <= self.config.delay <= 10:
            raise ValueError("Delay must be between 0 and 10 seconds")
        
    def _setup_typing_patterns(self) -> None:
        """Setup typing patterns based on configuration"""
        self.typing_patterns = {
            'fast': (0.05, 0.1),
            'realistic': (0.1, 0.3),
            'slow': (0.3, 0.5)
        }
        
    async def initialize(self) -> None:
        """Initialize the agent"""
        try:
            logger.info("Initializing Word AI Agent...")
            # Add initialization logic here
            await asyncio.sleep(self.config.delay)
            
            # Initialize template manager
            self.template_manager = TemplateManager()
            
            # Initialize Word if needed
            if self.config.use_word_automation and self.word_automation:
                self.state.using_word = True
                # Launch Word in a non-blocking way
                asyncio.create_task(self._launch_word())
            
            # Start voice command listener if enabled
            if self.state.voice_enabled and self.voice_integration and self.command_handler:
                await self._start_voice_command_listener()
            
            logger.info("Agent initialized successfully")
        except Exception as e:
            self.error_handler.handle_error(e)
            raise
    
    async def _launch_word(self) -> None:
        """Launch Microsoft Word in a separate task"""
        try:
            success = self.word_automation.open_word()
            if success:
                self.word_automation.open_document(self.config.word_doc_path)
            else:
                logger.error("Failed to launch Microsoft Word")
                self.state.using_word = False
        except Exception as e:
            logger.error(f"Error launching Word: {str(e)}")
            self.state.using_word = False
    
    async def _start_voice_command_listener(self) -> None:
        """Start the voice command listener"""
        try:
            if not self.voice_integration:
                logger.error("Voice integration not initialized")
                return
            
            # Define the command callback function
            def handle_voice_command(command_text):
                logger.info(f"Voice command received: {command_text}")
                # Process the command
                if self.command_handler:
                    result = self.command_handler.process_command(command_text, agent=self)
                    
                    # Log the result
                    success_str = "successful" if result.get('success', False) else "failed"
                    logger.info(f"Command processing {success_str}: {result.get('response', '')}")
                    
                    # Execute the action if needed
                    self._execute_voice_command_action(result)
                    
                    # Speak the response if available
                    response = result.get('response', '')
                    if response and hasattr(self.config, 'voice_speak_responses') and getattr(self.config, 'voice_speak_responses', False):
                        try:
                            self.voice_integration.speak(response)
                        except Exception as speak_err:
                            logger.error(f"Error speaking response: {speak_err}")
            
            # Start listening for voice commands
            self.voice_integration.start_voice_command_listener(handle_voice_command)
            logger.info("Voice command listener started successfully")
            
            # Announce that voice commands are active
            if hasattr(self.config, 'voice_speak_responses') and getattr(self.config, 'voice_speak_responses', False):
                welcome_msg = "Voice commands activated. Say 'help' for available commands."
                asyncio.create_task(self._speak_async(welcome_msg))
                
                # Start voice tutorial if it's the first time
                if self.voice_tutorial and hasattr(self.config, 'first_run') and getattr(self.config, 'first_run', True):
                    # Wait a moment before starting the tutorial
                    await asyncio.sleep(2.0)
                    self.voice_tutorial.start_tutorial()
                    # Mark as not first run anymore
                    if hasattr(self.config, 'first_run'):
                        self.config.first_run = False
                
        except Exception as e:
            logger.error(f"Error starting voice command listener: {e}")
            self.state.voice_enabled = False
    
    async def _speak_async(self, text):
        """Speak text asynchronously"""
        try:
            if self.voice_integration:
                self.voice_integration.speak(text)
        except Exception as e:
            logger.error(f"Error in async speech: {e}")
    
    def _execute_voice_command_action(self, command_result):
        """Execute actions based on voice command results"""
        if not command_result.get('success', False):
            return
        
        action = command_result.get('action')
        if not action:
            return
        
        try:
            # Map actions to agent methods
            action_map = {
                'start': self._action_start_writing,
                'stop': self._action_stop_writing,
                'pause': self._action_pause,
                'continue': self._action_continue,
                'write_about': self._action_write_about,
                'change_topic': self._action_change_topic,
                'add_section': self._action_add_section,
                'add_standard_section': self._action_add_standard_section,
                'add_references': self._action_add_references,
                'format': self._action_format,
                'text_style': self._action_text_style,
                'navigation': self._action_navigation,
                'goto_section': self._action_goto_section,
                'delete': self._action_delete,
                'undo': self._action_undo,
                'redo': self._action_redo,
                'typing_speed': self._action_typing_speed,
                'enable_feature': self._action_enable_feature,
                'disable_feature': self._action_disable_feature,
                'help': self._action_help,
                'list_commands': self._action_list_commands,
                'save': self._action_save,
                'exit': self._action_exit,
            }
            
            # Execute the action if available
            if action in action_map:
                action_method = action_map[action]
                asyncio.create_task(action_method(command_result))
            else:
                logger.warning(f"No implementation for action: {action}")
                
        except Exception as e:
            logger.error(f"Error executing voice command action: {e}")
    
    # Voice command action handlers
    async def _action_start_writing(self, cmd_result):
        """Start writing action"""
        logger.info("Starting writing process")
        # Implementation depends on specific application needs
    
    async def _action_stop_writing(self, cmd_result):
        """Stop writing action"""
        logger.info("Stopping writing process")
        # Implementation depends on specific application needs
    
    async def _action_pause(self, cmd_result):
        """Pause action"""
        logger.info("Pausing agent")
        # Implementation depends on specific application needs
    
    async def _action_continue(self, cmd_result):
        """Continue action"""
        logger.info("Continuing agent operation")
        # Implementation depends on specific application needs
    
    async def _action_write_about(self, cmd_result):
        """Write about a topic"""
        topic = cmd_result.get('topic', '')
        logger.info(f"Writing about new topic: {topic}")
        if topic and self.state.using_word and self.word_automation:
            # Start a new document section with the topic
            self.word_automation.insert_heading(f"Topic: {topic}")
    
    async def _action_change_topic(self, cmd_result):
        """Change current topic"""
        topic = cmd_result.get('topic', '')
        logger.info(f"Changing topic to: {topic}")
        # Implementation depends on specific application needs
    
    async def _action_add_section(self, cmd_result):
        """Add a new section"""
        section_type = cmd_result.get('section_type', '')
        section_name = cmd_result.get('section_name', '')
        logger.info(f"Adding {section_type}: {section_name}")
        if section_name and self.state.using_word and self.word_automation:
            if section_type == 'section':
                self.word_automation.insert_heading(section_name)
            else:  # paragraph
                self.word_automation.insert_paragraph(f"{section_name}\n")
    
    async def _action_add_standard_section(self, cmd_result):
        """Add a standard section"""
        section_name = cmd_result.get('section_name', '')
        logger.info(f"Adding standard section: {section_name}")
        if section_name and self.state.using_word and self.word_automation:
            # Add appropriate standard section
            if section_name.lower() == 'introduction':
                self.word_automation.insert_heading("Introduction")
                self.word_automation.insert_paragraph("This section introduces the topic.\n")
            elif section_name.lower() == 'conclusion':
                self.word_automation.insert_heading("Conclusion")
                self.word_automation.insert_paragraph("This section concludes the paper.\n")
    
    async def _action_add_references(self, cmd_result):
        """Add references section"""
        section_name = cmd_result.get('section_name', '')
        logger.info(f"Adding {section_name} section")
        if self.state.using_word and self.word_automation:
            self.word_automation.insert_heading(section_name.title())
            self.word_automation.insert_paragraph("1. Reference 1\n2. Reference 2\n")
    
    async def _action_format(self, cmd_result):
        """Format content"""
        format_type = cmd_result.get('format_type', '')
        logger.info(f"Formatting as: {format_type}")
        # Implementation depends on specific application needs
    
    async def _action_text_style(self, cmd_result):
        """Apply text style"""
        text = cmd_result.get('text', '')
        style = cmd_result.get('style', '')
        logger.info(f"Applying {style} style to '{text}'")
        # Implementation depends on specific application needs
    
    async def _action_navigation(self, cmd_result):
        """Navigate in document"""
        position = cmd_result.get('position', '')
        logger.info(f"Navigating to: {position}")
        if self.state.using_word and self.word_automation:
            if position in ['top', 'start']:
                self.word_automation.navigate_to_start()
            elif position in ['bottom', 'end']:
                self.word_automation.navigate_to_end()
    
    async def _action_goto_section(self, cmd_result):
        """Go to specific section"""
        section_name = cmd_result.get('section_name', '')
        logger.info(f"Going to section: {section_name}")
        # Implementation depends on specific application needs
    
    async def _action_delete(self, cmd_result):
        """Delete content"""
        target_position = cmd_result.get('target_position', '')
        target_type = cmd_result.get('target_type', '')
        logger.info(f"Deleting {target_position} {target_type}")
        # Implementation depends on specific application needs
    
    async def _action_undo(self, cmd_result):
        """Undo action"""
        logger.info("Undoing last action")
        if self.state.using_word and self.word_automation:
            self.word_automation.send_keys('^z')  # Ctrl+Z for undo
    
    async def _action_redo(self, cmd_result):
        """Redo action"""
        logger.info("Redoing last undone action")
        if self.state.using_word and self.word_automation:
            self.word_automation.send_keys('^y')  # Ctrl+Y for redo
    
    async def _action_typing_speed(self, cmd_result):
        """Change typing speed"""
        speed = cmd_result.get('speed', '')
        logger.info(f"Changing typing speed to: {speed}")
        if speed in ['fast', 'realistic', 'slow']:
            self.config.typing_mode = speed
            self.typing_simulator = TypingSimulator(speed)
    
    async def _action_enable_feature(self, cmd_result):
        """Enable a feature"""
        feature = cmd_result.get('feature', '')
        logger.info(f"Enabling feature: {feature}")
        # Map feature names to config attributes
        feature_map = {
            'self improve': 'self_improve',
            'self improvement': 'self_improve',
            'self evolve': 'self_evolve',
            'self evolution': 'self_evolve',
            'track position': 'track_position',
            'robot cursor': 'robot_cursor',
            'word automation': 'use_word_automation',
        }
        if feature.lower() in feature_map:
            setattr(self.config, feature_map[feature.lower()], True)
            logger.info(f"Enabled {feature}")
    
    async def _action_disable_feature(self, cmd_result):
        """Disable a feature"""
        feature = cmd_result.get('feature', '')
        logger.info(f"Disabling feature: {feature}")
        # Map feature names to config attributes
        feature_map = {
            'self improve': 'self_improve',
            'self improvement': 'self_improve',
            'self evolve': 'self_evolve',
            'self evolution': 'self_evolve',
            'track position': 'track_position',
            'robot cursor': 'robot_cursor',
            'word automation': 'use_word_automation',
        }
        if feature.lower() in feature_map:
            setattr(self.config, feature_map[feature.lower()], False)
            logger.info(f"Disabled {feature}")
    
    async def _action_save(self, cmd_result):
        """Save document"""
        logger.info("Saving document")
        if self.state.using_word and self.word_automation:
            self.word_automation.save_document()
    
    async def _action_exit(self, cmd_result):
        """Exit application"""
        logger.info("Exiting application")
        # Implementation depends on specific application needs
        await self.finalize()
    
    async def _action_help(self, cmd_result):
        """Help action"""
        logger.info("Providing help information")
        if self.voice_tutorial:
            # Check if there's a specific help category
            help_category = None
            if 'params' in cmd_result and 'match_groups' in cmd_result['params'] and cmd_result['params']['match_groups']:
                # Extract help category if available
                match_text = cmd_result['params']['match_groups'][0] if cmd_result['params']['match_groups'] else ''
                if match_text:
                    # Extract category from "help with [category] commands"
                    import re
                    category_match = re.search(r'with\s+([\w\s]+)\s+commands', match_text)
                    if category_match:
                        help_category = category_match.group(1).strip().lower()
                        
            # Provide help based on category
            self.voice_tutorial.provide_command_help(help_category)
            
            # Ask what they want to do next
            await asyncio.sleep(2.0)  # Wait for help to finish
            self.voice_tutorial.ask_question("next_step")
    
    async def _action_list_commands(self, cmd_result):
        """List commands action"""
        logger.info("Listing available commands")
        if self.voice_tutorial:
            # List all command categories
            self.voice_integration.speak(
                "I can help you with different types of commands including: "
                "basic control, document creation, navigation, editing, "
                "formatting, and system commands. Say 'help with' followed by "
                "a category name for more information."
            )
    
    async def process_text(self, text: str, template_id: Optional[str] = None) -> None:
        """Process and type text with error handling"""
        try:
            self.state.is_typing = True
            logger.info(f"Processing text: {text[:50]}...")
            
            # Get template if provided or use default
            template = None
            if template_id is not None:
                # If a template_id was provided
                template = self.template_manager.get_template_by_id(template_id)
            elif self.config.template_id:
                # If config has a template_id
                template = self.template_manager.get_template_by_id(self.config.template_id)
                
            # Generate paper content based on template
            content = self._generate_paper_content(text, template)
            
            # If using Word, type into Word document
            if self.state.using_word and self.word_automation:
                # Wait for Word to be ready
                for _ in range(10):  # Wait up to 10 seconds
                    if self.word_automation._is_connected:
                        break
                    await asyncio.sleep(1)
                
                # Insert template structure if available
                if template and template_id:
                    self.word_automation.insert_template(template_id, text)
                    self.state.is_typing = False
                    logger.info("Template inserted successfully in Word")
                    return
                
                # Type the generated content into Word
                typing_delay = self.typing_patterns[self.config.typing_mode][0]
                self.word_automation.type_text(content, delay=typing_delay)
                self.state.is_typing = False
                logger.info("Text processing completed successfully in Word")
                return
            
            # If not using Word, simulate typing
            for char in content:
                if self.state.retry_count >= self.config.max_retries:
                    raise Exception("Maximum retries exceeded")
                    
                try:
                    await self.typing_simulator.type_char(char)
                    await asyncio.sleep(random.uniform(*self.typing_patterns[self.config.typing_mode]))
                    self.state.current_position += 1
                except Exception as e:
                    self.state.error_count += 1
                    self.error_handler.handle_error(e)
                    self.state.retry_count += 1
            
            self.state.is_typing = False
            logger.info("Text processing completed successfully")
        except Exception as e:
            self.state.is_typing = False
            self.error_handler.handle_error(e)
            raise
    
    def _generate_paper_content(self, topic: str, template: Optional[PaperTemplate] = None) -> str:
        """Generate paper content based on topic and template"""
        # If no template, generate simple content
        if not template:
            return f"Paper on: {topic}\n\nThis is a simple paper about {topic}.\n\nIntroduction:\nBody:\nConclusion:\n"
        
        # Generate structured content based on template
        content = f"Paper on: {topic}\n\n"
        
        # Add sections from template
        for section in template.sections:
            # Check if section is a string or an object with title/description
            if isinstance(section, str):
                content += f"\n## {section}\n"
                
                if section == "Title":
                    content += f"{topic.upper()}\n"
                elif section == "Abstract":
                    content += f"This paper examines {topic} in detail...\n"
                else:
                    content += f"Content for {section} about {topic}...\n"
            else:
                # Assuming section is an object with title and description
                content += f"{section.title}\n"
                content += f"{section.description}\n\n"
        
        # Add formatting notes if available
        if hasattr(template, 'formatting') and template.formatting:
            content += "\n\n--- Formatting Notes ---\n"
            for key, value in template.formatting.items():
                content += f"{key}: {value}\n"
        
        # Add citation style if available
        if hasattr(template, 'citation_style') and template.citation_style:
            content += f"Citation Style: {template.citation_style}\n"
        
        return content
    
    async def finalize(self) -> None:
        """Finalize the agent's work"""
        try:
            logger.info("Finalizing Word AI Agent...")
            
            # Stop voice command listener if active
            if self.state.voice_enabled and self.voice_integration:
                # Stop tutorial if running
                if self.voice_tutorial:
                    self.voice_tutorial.stop_tutorial()
                    logger.info("Voice command tutorial stopped")
                
                self.voice_integration.stop_voice_command_listener()
                logger.info("Voice command listener stopped")
            
            # Clean up Word automation if active
            if self.state.using_word and self.word_automation:
                # Save document if needed
                if self.config.word_doc_path:
                    self.word_automation.save_document(self.config.word_doc_path)
                    logger.info(f"Document saved to {self.config.word_doc_path}")
                
                # Clean up resources
                self.word_automation.cleanup()
                logger.info("Word automation resources cleaned up")
            
            logger.info("Agent finalized successfully")
        except Exception as e:
            self.error_handler.handle_error(e)
            raise

class AgentFactory:
    """Factory for creating agents"""
    
    @staticmethod
    def create_agent(config: Config) -> BaseAgent:
        """Create a new agent instance"""
        return WordAIAgent(config)
