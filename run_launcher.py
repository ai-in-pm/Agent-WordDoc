"""
Launcher script for the Word AI Agent

Provides options to run the agent with different functionality modes.
"""

import os
import sys
import asyncio
import argparse
import logging
import traceback
from datetime import datetime

from src.agents.word_agent import AgentFactory, WordAIAgent  
from src.config.config import Config, DEFAULT_CONFIG
from src.utils.error_handler import ErrorHandler
from src.templates.template_manager import TemplateManager
from src.interfaces.gui_interface import run_gui
from src.core.logging import setup_logger, get_logger

# Initialize logger
logger = setup_logger(__name__)

def parse_arguments():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description="Word AI Agent Launcher")
    
    # Mode selection - mutually exclusive
    mode_group = parser.add_argument_group('Mode')
    mode_group.add_argument('--gui', action='store_true', help='Run with GUI interface')
    mode_group.add_argument('--cli', action='store_true', help='Run with command-line interface')
    mode_group.add_argument('--service', action='store_true', help='Run as a background service')
    mode_group.add_argument('--interactive', action='store_true', help='Run in interactive mode')
    
    # Basic options
    parser.add_argument('--topic', help='Paper topic')
    parser.add_argument('--config', help='Path to configuration file')
    parser.add_argument('--typing-mode', choices=['fast', 'realistic', 'slow'], default='realistic',
                      help='Typing behavior mode')
    parser.add_argument('--template-id', help='Template ID to use')
    parser.add_argument('--output', help='Output document path (for saving Word documents)')
    
    # Feature flags
    feature_group = parser.add_argument_group('Features')
    feature_group.add_argument('--verbose', action='store_true', help='Enable verbose output')
    feature_group.add_argument('--iterative', action='store_true', help='Enable iterative processing')
    feature_group.add_argument('--self-improve', action='store_true', help='Enable self-improvement')
    feature_group.add_argument('--self-evolve', action='store_true', help='Enable self-evolution')
    feature_group.add_argument('--track-position', action='store_true', help='Enable position tracking')
    
    # Robot options
    robot_group = parser.add_argument_group('Robot Options')
    robot_group.add_argument('--robot-cursor', action='store_true', help='Show robot cursor')
    robot_group.add_argument('--no-robot-cursor', action='store_true', help='Disable robot cursor')
    robot_group.add_argument('--cursor-size', choices=['standard', 'large', 'extra_large'], default='standard',
                           help='Size of the robot cursor')
    robot_group.add_argument('--use-autoit', action='store_true', help='Use AutoIt')
    robot_group.add_argument('--no-autoit', action='store_true', help='Disable AutoIt')
    
    # Word automation options
    word_group = parser.add_argument_group('Word Automation')
    word_group.add_argument('--use-word', action='store_true', help='Enable Microsoft Word automation')
    word_group.add_argument('--word-doc', help='Path to Word document (will be created if not exists)')
    
    # Voice command options
    voice_group = parser.add_argument_group('Voice Commands')
    voice_group.add_argument('--voice-commands', action='store_true', help='Enable voice command recognition')
    voice_group.add_argument('--voice-language', choices=['en-US', 'fr-FR', 'es-ES', 'de-DE', 'it-IT', 'pt-PT', 'zh-CN', 'ja-JP', 'ko-KR'], 
                            default='en-US', help='Voice command language')
    voice_group.add_argument('--voice-threshold', type=float, default=0.5, 
                            help='Voice recognition threshold (0.0-1.0)')
    voice_group.add_argument('--voice-device', type=int, default=0, 
                            help='Microphone device index')
    voice_group.add_argument('--voice-speak-responses', action='store_true', 
                            help='Enable spoken responses to voice commands')
    voice_group.add_argument('--first-run', action='store_true',
                            help='Treat as first run and enable voice tutorial')
    
    # Advanced options
    advanced_group = parser.add_argument_group('Advanced')
    advanced_group.add_argument('--delay', type=float, default=3.0, help='Delay before starting')
    advanced_group.add_argument('--max-retries', type=int, default=3, help='Maximum retries on error')
    advanced_group.add_argument('--retry-delay', type=float, default=1.0, help='Delay between retries')
    advanced_group.add_argument('--log-level', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                              default='INFO', help='Logging level')
    advanced_group.add_argument('--plugins-dir', help='Plugins directory')
    
    return parser.parse_args()

def create_config_from_args(args):
    """Create configuration from command-line arguments"""
    config = Config(
        api_key="",  # Will be loaded from environment if not specified
        typing_mode=args.typing_mode,
        verbose=args.verbose,
        iterative=args.iterative,
        self_improve=args.self_improve,
        self_evolve=args.self_evolve,
        track_position=args.track_position,
        robot_cursor=not args.no_robot_cursor if hasattr(args, 'no_robot_cursor') and args.no_robot_cursor else True,
        cursor_size=args.cursor_size if hasattr(args, 'cursor_size') else 'standard',
        use_autoit=not args.no_autoit if hasattr(args, 'no_autoit') and args.no_autoit else True,
        delay=args.delay,
        max_retries=args.max_retries,
        retry_delay=args.retry_delay,
        log_level=args.log_level,
        template_id=args.template_id if hasattr(args, 'template_id') and args.template_id else 'research-paper',
        use_word_automation=args.use_word if hasattr(args, 'use_word') and args.use_word else False,
        word_doc_path=args.word_doc if hasattr(args, 'word_doc') and args.word_doc else (args.output if hasattr(args, 'output') and args.output else None),
        voice_command_enabled=args.voice_commands if hasattr(args, 'voice_commands') and args.voice_commands else False,
        voice_command_language=args.voice_language if hasattr(args, 'voice_language') and args.voice_language else 'en-US',
        voice_command_threshold=args.voice_threshold if hasattr(args, 'voice_threshold') and args.voice_threshold else 0.5,
        voice_command_device_index=args.voice_device if hasattr(args, 'voice_device') and args.voice_device else 0,
        voice_speak_responses=args.voice_speak_responses if hasattr(args, 'voice_speak_responses') and args.voice_speak_responses else False,
        first_run=args.first_run if hasattr(args, 'first_run') and args.first_run else False
    )
    return config

async def run_agent(config, topic, template_id):
    """Run the agent with the specified configuration"""
    error_handler = ErrorHandler()
    
    try:
        # Create agent
        logger.info(f"Creating agent for topic: {topic}")
        agent = AgentFactory.create_agent(config)
        
        # Initialize agent
        logger.info("Initializing agent")
        await agent.initialize()
        
        # Process text
        text = f"Writing paper on: {topic}"
        logger.info(f"Processing text: {text}")
        await agent.process_text(text, template_id)
        
        # Finalize agent
        logger.info("Finalizing agent")
        await agent.finalize()
        
        logger.info("Agent completed successfully")
        return True
    except Exception as e:
        error_handler.handle_error(e)
        return False

def run_cli(args, config):
    """Run in command-line interface mode"""
    logger.info("Running in CLI mode")
    
    # Set default API key for testing if not provided
    if not config.api_key:
        logger.info("No API key provided, using test key for development")
        config.api_key = "test_key_for_development"
    
    template_id = args.template_id
    
    asyncio.run(run_agent(config, args.topic, template_id))

def run_interactive(args, config):
    """Run in interactive mode"""
    logger.info("Running in interactive mode")
    print("Word AI Agent Interactive Mode")
    print("============================")
    
    # Get API key if not already set
    if not config.api_key:
        try:
            config.api_key = input("Enter your API key (or press Enter to use a test key): ") or "test_key_for_development"
            print(f"Using API key: {'*' * 5 + config.api_key[-4:] if len(config.api_key) > 8 else 'test_key'}")
        except (EOFError, KeyboardInterrupt):
            print("\nNo API key provided, using test key for development")
            config.api_key = "test_key_for_development"
    
    # Get topic if not specified
    try:
        topic = args.topic or input("Enter paper topic (default: Earned Value Management): ") or "Earned Value Management"
        print(f"Topic: {topic}")
    except (EOFError, KeyboardInterrupt):
        print("\nNo topic provided, using default topic")
        topic = "Earned Value Management"
    
    # Get typing mode if not specified
    if not args.typing_mode:
        try:
            modes = ["fast", "realistic", "slow"]
            print("Select typing mode:")
            for i, mode in enumerate(modes, 1):
                print(f"{i}. {mode}")
            selection = input("Enter number (default: 2): ") or "2"
            try:
                config.typing_mode = modes[int(selection) - 1]
            except (ValueError, IndexError):
                print(f"Invalid selection, using default: realistic")
                config.typing_mode = "realistic"
        except (EOFError, KeyboardInterrupt):
            print("\nNo mode selected, using default: realistic")
            config.typing_mode = "realistic"
    
    print(f"Typing mode: {config.typing_mode}")
    
    # Ask about features - with error handling
    try:
        if input("Enable self-improvement? (y/n, default: y): ").lower() != 'n':
            config.self_improve = True
            print("Self-improvement: Enabled")
        else:
            config.self_improve = False
            print("Self-improvement: Disabled")
    except (EOFError, KeyboardInterrupt):
        print("\nUsing default: Self-improvement enabled")
        config.self_improve = True
    
    try:
        if input("Enable self-evolution? (y/n, default: y): ").lower() != 'n':
            config.self_evolve = True
            print("Self-evolution: Enabled")
        else:
            config.self_evolve = False
            print("Self-evolution: Disabled")
    except (EOFError, KeyboardInterrupt):
        print("\nUsing default: Self-evolution enabled")
        config.self_evolve = True
    
    # Get template if not specified
    try:
        template_id = args.template_id or input("Enter template ID (default: None): ") or None
        if template_id:
            print(f"Template ID: {template_id}")
    except (EOFError, KeyboardInterrupt):
        print("\nNo template ID provided, using default")
        template_id = None
    
    # Run the agent
    print(f"\nRunning agent with topic: {topic}")
    print(f"Typing mode: {config.typing_mode}")
    print(f"Self-improvement: {'Enabled' if config.self_improve else 'Disabled'}")
    print(f"Self-evolution: {'Enabled' if config.self_evolve else 'Disabled'}")
    print(f"Robot cursor: {'Enabled' if config.robot_cursor else 'Disabled'}")
    print(f"AutoIt: {'Enabled' if config.use_autoit else 'Disabled'}")
    print(f"Template ID: {template_id}")
    print("\nStarting in 3 seconds...")
    
    # Run agent
    asyncio.run(run_agent(config, topic, template_id))

def run_service(args, config):
    """Run as a background service"""
    logger.info("Running as a service")
    print("Service mode not yet implemented")

def main():
    """Main entry point"""
    # Parse arguments
    args = parse_arguments()
    
    # Set up logging
    logging.basicConfig(level=getattr(logging, args.log_level))
    
    # Create config
    if args.config:
        try:
            config = Config.from_yaml(args.config)
        except Exception as e:
            logger.error(f"Error loading config: {str(e)}")
            config = create_config_from_args(args)
    else:
        config = create_config_from_args(args)
    
    # Determine run mode
    if args.gui:
        # Run with GUI
        run_gui()
    elif args.interactive:
        # Run in interactive mode
        run_interactive(args, config)
    elif args.service:
        # Run as a service
        run_service(args, config)
    else:
        # Default to CLI mode
        run_cli(args, config)

if __name__ == "__main__":
    main()
