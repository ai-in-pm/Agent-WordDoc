"""
Plugin Manager for the Word AI Agent

This module manages the plugin system, allowing for extensibility and customization.
"""

import os
import importlib.util
from typing import Dict, Any, List, Callable, Optional
from pathlib import Path

from src.core.logging import get_logger

logger = get_logger(__name__)

class PluginManager:
    """Manages plugins for the Word AI Agent"""
    
    def __init__(self, plugin_directory: str = "plugins"):
        self.plugin_directory = Path(plugin_directory)
        self.plugins: Dict[str, Any] = {}
        self.hooks: Dict[str, List[Callable]] = {}
        self._initialize_hooks()
    
    def _initialize_hooks(self):
        """Initialize standard hooks"""
        standard_hooks = [
            "before_initialization",
            "after_initialization",
            "before_typing",
            "after_typing",
            "before_processing",
            "after_processing",
            "before_finalization",
            "after_finalization",
            "on_error",
            "on_success"
        ]
        for hook in standard_hooks:
            self.hooks[hook] = []
    
    def discover_plugins(self) -> None:
        """Discover and load plugins from the plugin directory"""
        if not self.plugin_directory.exists():
            logger.warning(f"Plugin directory {self.plugin_directory} does not exist")
            return
        
        for plugin_file in self.plugin_directory.glob("*.py"):
            if plugin_file.name.startswith("__"):
                continue
            
            try:
                self.load_plugin(plugin_file)
            except Exception as e:
                logger.error(f"Error loading plugin {plugin_file.name}: {str(e)}")
    
    def load_plugin(self, plugin_path: Path) -> None:
        """Load a plugin from a file"""
        plugin_name = plugin_path.stem
        
        try:
            spec = importlib.util.spec_from_file_location(plugin_name, plugin_path)
            if spec is None or spec.loader is None:
                logger.error(f"Could not load plugin {plugin_name}")
                return
            
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            
            if not hasattr(module, "register_plugin"):
                logger.error(f"Plugin {plugin_name} does not have a register_plugin function")
                return
            
            module.register_plugin(self)
            self.plugins[plugin_name] = module
            logger.info(f"Loaded plugin {plugin_name}")
        except Exception as e:
            logger.error(f"Error loading plugin {plugin_name}: {str(e)}")
            raise
    
    def register_hook(self, hook_name: str, callback: Callable) -> None:
        """Register a callback for a hook"""
        if hook_name not in self.hooks:
            self.hooks[hook_name] = []
        self.hooks[hook_name].append(callback)
        logger.debug(f"Registered hook {hook_name}")
    
    def execute_hook(self, hook_name: str, **kwargs) -> List[Any]:
        """Execute all callbacks for a hook"""
        if hook_name not in self.hooks:
            logger.warning(f"Hook {hook_name} does not exist")
            return []
        
        results = []
        for callback in self.hooks[hook_name]:
            try:
                results.append(callback(**kwargs))
            except Exception as e:
                logger.error(f"Error executing hook {hook_name}: {str(e)}")
        
        return results
    
    def get_plugin_info(self) -> Dict[str, Dict[str, Any]]:
        """Get information about loaded plugins"""
        plugin_info = {}
        for plugin_name, plugin in self.plugins.items():
            info = {
                "name": plugin_name,
                "version": getattr(plugin, "__version__", "unknown"),
                "description": getattr(plugin, "__description__", ""),
                "author": getattr(plugin, "__author__", "unknown"),
                "hooks": []
            }
            
            for hook_name, callbacks in self.hooks.items():
                for callback in callbacks:
                    if callback.__module__ == plugin.__name__:
                        info["hooks"].append(hook_name)
            
            plugin_info[plugin_name] = info
        
        return plugin_info
