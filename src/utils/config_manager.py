"""
Configuration Management Utility for the Word AI Agent

Provides centralized configuration management with support for multiple environments.
"""

import os
import yaml
import json
from typing import Dict, Any, Optional
from pathlib import Path
from dataclasses import dataclass, asdict

from src.core.logging import get_logger

logger = get_logger(__name__)

class ConfigManager:
    """Manages configuration for the Word AI Agent"""
    
    def __init__(self, config_dir: str = "config"):
        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.configs = {}
        self.active_config = None
        self.active_environment = "default"
    
    def load_config(self, name: str, environment: str = "default") -> Dict[str, Any]:
        """Load configuration from file"""
        config_path = self._get_config_path(name, environment)
        
        if not config_path.exists():
            logger.warning(f"Configuration file {config_path} not found")
            return {}
        
        logger.info(f"Loading configuration from {config_path}")
        
        try:
            if config_path.suffix == ".yaml" or config_path.suffix == ".yml":
                with open(config_path, "r") as f:
                    config = yaml.safe_load(f)
            elif config_path.suffix == ".json":
                with open(config_path, "r") as f:
                    config = json.load(f)
            else:
                logger.error(f"Unsupported configuration format: {config_path.suffix}")
                return {}
            
            self.configs[(name, environment)] = config
            return config
        except Exception as e:
            logger.error(f"Error loading configuration: {str(e)}")
            return {}
    
    def save_config(self, config: Dict[str, Any], name: str, environment: str = "default") -> bool:
        """Save configuration to file"""
        config_path = self._get_config_path(name, environment)
        logger.info(f"Saving configuration to {config_path}")
        
        try:
            if config_path.suffix == ".yaml" or config_path.suffix == ".yml":
                with open(config_path, "w") as f:
                    yaml.dump(config, f, default_flow_style=False)
            elif config_path.suffix == ".json":
                with open(config_path, "w") as f:
                    json.dump(config, f, indent=2)
            else:
                logger.error(f"Unsupported configuration format: {config_path.suffix}")
                return False
            
            self.configs[(name, environment)] = config
            return True
        except Exception as e:
            logger.error(f"Error saving configuration: {str(e)}")
            return False
    
    def get_config(self, name: str, environment: str = "default") -> Dict[str, Any]:
        """Get configuration by name and environment"""
        # Check if already loaded
        if (name, environment) in self.configs:
            return self.configs[(name, environment)]
        
        # Try to load if not already loaded
        return self.load_config(name, environment)
    
    def set_active_config(self, name: str, environment: str = "default") -> bool:
        """Set the active configuration"""
        config = self.get_config(name, environment)
        if not config:
            logger.error(f"Configuration {name} for environment {environment} not found")
            return False
        
        self.active_config = name
        self.active_environment = environment
        logger.info(f"Set active configuration to {name} ({environment})")
        return True
    
    def get_active_config(self) -> Optional[Dict[str, Any]]:
        """Get the active configuration"""
        if not self.active_config:
            logger.warning("No active configuration set")
            return None
        
        return self.get_config(self.active_config, self.active_environment)
    
    def update_config(self, updates: Dict[str, Any], name: str = None, environment: str = None) -> bool:
        """Update configuration with new values"""
        # Use active config if not specified
        if name is None:
            name = self.active_config
        if environment is None:
            environment = self.active_environment
        
        if not name:
            logger.error("No configuration specified and no active configuration set")
            return False
        
        config = self.get_config(name, environment)
        if not config:
            logger.error(f"Configuration {name} for environment {environment} not found")
            return False
        
        # Update configuration
        config.update(updates)
        self.configs[(name, environment)] = config
        
        # Save updated configuration
        return self.save_config(config, name, environment)
    
    def list_configs(self) -> Dict[str, Dict[str, Any]]:
        """List all available configurations"""
        configs = {}
        
        for path in self.config_dir.glob("**/*.yaml"):
            relative_path = path.relative_to(self.config_dir)
            parts = list(relative_path.parts)
            
            if len(parts) >= 2:
                environment = parts[0]
                name = parts[1].split(".")[0]
            else:
                environment = "default"
                name = path.stem
            
            if name not in configs:
                configs[name] = {}
            
            configs[name][environment] = str(path)
        
        return configs
    
    def get_environment_variables(self) -> Dict[str, str]:
        """Get relevant environment variables"""
        env_vars = {}
        for key, value in os.environ.items():
            if key.startswith("WORD_AI_"):
                env_vars[key] = value
        return env_vars
    
    def _get_config_path(self, name: str, environment: str) -> Path:
        """Get path to configuration file"""
        if environment == "default":
            return self.config_dir / f"{name}.yaml"
        else:
            environment_dir = self.config_dir / environment
            environment_dir.mkdir(exist_ok=True)
            return environment_dir / f"{name}.yaml"

# Singleton instance
_config_manager = None

def get_config_manager() -> ConfigManager:
    """Get the singleton ConfigManager instance"""
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager
