"""
Configuration management for the Word AI Agent
"""

import os
from dataclasses import dataclass
from typing import Dict, Any, Optional
import yaml
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

@dataclass
class Config:
    """Main configuration class"""
    api_key: str = None
    typing_mode: str = 'realistic'
    verbose: bool = False
    iterative: bool = False
    self_improve: bool = False
    self_evolve: bool = False
    track_position: bool = False
    robot_cursor: bool = True
    cursor_size: str = 'standard'  # Options: standard, large, extra_large
    use_autoit: bool = True
    delay: float = 3.0
    max_retries: int = 3
    retry_delay: float = 1.0
    log_level: str = 'INFO'
    template_id: str = 'research-paper'
    use_word_automation: bool = False
    word_doc_path: Optional[str] = None
    voice_command_enabled: bool = False
    voice_command_language: str = 'en-US'
    voice_command_threshold: float = 0.5
    voice_command_device_index: int = 0
    voice_speak_responses: bool = False
    first_run: bool = True  # Flag to track if it's the first run with voice commands

    def __post_init__(self):
        # Load API key from environment if not provided
        if not self.api_key:
            self.api_key = os.environ.get('OPENAI_API_KEY', None)

    @classmethod
    def from_yaml(cls, config_path: str) -> 'Config':
        """Load configuration from YAML file"""
        with open(config_path, 'r') as f:
            config_data = yaml.safe_load(f)
        return cls(**config_data)

    def to_yaml(self, config_path: str) -> None:
        """Save configuration to YAML file"""
        with open(config_path, 'w') as f:
            yaml.dump(self.__dict__, f, default_flow_style=False)

    def validate(self) -> bool:
        """Validate configuration values"""
        # API key validation is handled by the agent
        if self.typing_mode not in ['fast', 'realistic', 'slow']:
            raise ValueError(f"Invalid typing mode: {self.typing_mode}")
        if self.voice_command_language not in ['en-US', 'fr-FR', 'es-ES', 'de-DE', 'it-IT', 'pt-PT', 'zh-CN', 'ja-JP', 'ko-KR']:
            raise ValueError(f"Invalid voice command language: {self.voice_command_language}")
        if self.voice_command_threshold < 0 or self.voice_command_threshold > 1:
            raise ValueError(f"Invalid voice command threshold: {self.voice_command_threshold}")
        return True

# Default configuration
DEFAULT_CONFIG = {
    'api_key': os.environ.get('OPENAI_API_KEY', ''),
    'typing_mode': 'realistic',
    'verbose': False,
    'iterative': False,
    'self_improve': False,
    'self_evolve': False,
    'track_position': False,
    'robot_cursor': True,
    'cursor_size': 'standard',
    'use_autoit': True,
    'delay': 3.0,
    'max_retries': 3,
    'retry_delay': 1.0,
    'log_level': 'INFO',
    'template_id': 'research-paper',
    'use_word_automation': False,
    'word_doc_path': None,
    'voice_command_enabled': False,
    'voice_command_language': 'en-US',
    'voice_command_threshold': 0.5,
    'voice_command_device_index': 0,
    'voice_speak_responses': False,
    'first_run': True
}
