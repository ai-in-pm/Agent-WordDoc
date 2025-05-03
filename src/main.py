"""
Main entry point for the Word AI Agent
"""

import asyncio
import argparse
import sys
from pathlib import Path
from typing import Optional

from src.config.config import Config, DEFAULT_CONFIG
from src.agents.word_agent import AgentFactory
from src.core.logging import set_log_level
from src.utils.error_handler import ErrorHandler

logger = get_logger(__name__)

def parse_arguments() -> argparse.Namespace:
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description='Word AI Agent')
    parser.add_argument('--topic', default='Earned Value Management',
                        help='Paper topic')
    parser.add_argument('--typing-mode', choices=['fast', 'realistic', 'slow'], default='realistic',
                        help='Typing behavior mode')
    parser.add_argument('--config', type=Path, default='config.yaml',
                        help='Path to configuration file')
    parser.add_argument('--verbose', action='store_true',
                        help='Enable verbose output')
    parser.add_argument('--log-level', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                        default='INFO',
                        help='Log level')
    
    return parser.parse_args()

def load_config(config_path: Path) -> Config:
    """Load configuration from file or use defaults"""
    try:
        if config_path.exists():
            return Config.from_yaml(config_path)
        return Config(**DEFAULT_CONFIG)
    except Exception as e:
        logger.error(f"Error loading configuration: {str(e)}")
        return Config(**DEFAULT_CONFIG)

async def main() -> None:
    """Main function"""
    try:
        args = parse_arguments()
        set_log_level(args.log_level)
        
        # Load configuration
        config = load_config(args.config)
        
        # Create error handler
        error_handler = ErrorHandler()
        
        # Create and initialize agent
        agent = AgentFactory.create_agent(config)
        await agent.initialize()
        
        # Process text
        text = f"Writing paper on: {args.topic}"
        await agent.process_text(text)
        
        # Finalize agent
        await agent.finalize()
        
        logger.info("Agent completed successfully")
        
    except Exception as e:
        logger.critical(f"Critical error: {str(e)}")
        error_handler.handle_error(e)
        sys.exit(1)

if __name__ == "__main__":
    asyncio.run(main())
