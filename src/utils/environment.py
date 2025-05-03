"""
Environment variables handling utility for Agent WordDoc
"""

import os
import logging
from typing import Dict, Optional
from dotenv import load_dotenv
from pathlib import Path

from src.core.logging import get_logger

# Initialize logger
logger = get_logger(__name__)

def load_env_variables(env_file: Optional[str] = None) -> Dict[str, str]:
    """
    Load environment variables from .env file
    
    Args:
        env_file: Path to .env file, if None, looks in default locations
        
    Returns:
        Dict of environment variables
    """
    # If env_file is not provided, try to find it in standard locations
    if env_file is None:
        root_dir = Path(__file__).parent.parent.parent
        env_path = root_dir / '.env'
        if not env_path.exists():
            # Try alternative locations
            alt_locations = [
                root_dir.parent / '.env',
                Path(os.getcwd()) / '.env'
            ]
            for loc in alt_locations:
                if loc.exists():
                    env_path = loc
                    break
    else:
        env_path = Path(env_file)
    
    # Load environment variables
    if env_path.exists():
        logger.info(f"Loading environment variables from {env_path}")
        load_dotenv(dotenv_path=str(env_path))
    else:
        logger.warning(f"No .env file found at {env_path}")
    
    # Get relevant environment variables for the application
    env_vars = {
        'OPENAI_API_KEY': os.environ.get('OPENAI_API_KEY', ''),
        'ELEVENLABS_API_KEY': os.environ.get('ELEVENLABS_API_KEY', ''),
        'ELEVENLABS_VOICE_ID': os.environ.get('ELEVENLABS_VOICE_ID', ''),
    }
    
    # Log loaded variables (masked for security)
    for key, value in env_vars.items():
        masked_value = '****' + value[-4:] if value and len(value) > 4 else '(not set)'
        logger.debug(f"Loaded environment variable {key}: {masked_value}")
    
    return env_vars
