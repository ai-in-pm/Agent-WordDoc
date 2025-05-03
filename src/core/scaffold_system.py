"""
Scaffold System for the Word AI Agent

Provides capability evolution and self-improvement capabilities.
"""

import os
import json
import time
from datetime import datetime
from typing import Dict, Any, List, Optional, Union, Callable
from pathlib import Path
from enum import Enum
import importlib
import inspect

from src.core.logging import get_logger

logger = get_logger(__name__)

class CapabilityType(Enum):
    """Types of capabilities the system can have"""
    CORE = "core"  # Essential functions and operations
    INTERACTION = "interaction"  # Interaction with applications
    ANALYSIS = "analysis"  # Analysis of content and structure
    GENERATION = "generation"  # Content generation capabilities
    ADAPTATION = "adaptation"  # Adaptation to contexts
    META = "meta"  # Capabilities working with other capabilities

class EvolutionStage(Enum):
    """Evolution stages for capabilities"""
    CONCEPTION = "conception"  # Initial idea and basic implementation
    PROTOTYPE = "prototype"  # Working implementation with issues
    STABLE = "stable"  # Reliable implementation for regular use
    OPTIMIZED = "optimized"  # Performance-optimized implementation
    ADVANCED = "advanced"  # Sophisticated implementation with extra features

class Capability:
    """Represents a capability in the scaffold system"""
    
    def __init__(self, name: str, description: str, capability_type: CapabilityType,
                function: Optional[Callable] = None, parameters: List[str] = None,
                evolution_stage: EvolutionStage = EvolutionStage.CONCEPTION):
        self.id = f"{int(time.time() * 1000)}_{id(self)}"
        self.name = name
        self.description = description
        self.capability_type = capability_type
        self.function = function
        self.parameters = parameters or []
        self.evolution_stage = evolution_stage
        self.created_at = datetime.now()
        self.last_evolved = datetime.now()
        self.usage_count = 0
        self.success_count = 0
        self.dependencies = []
        self.evolution_history = []
    
    def use(self, success: bool = True) -> None:
        """Record usage of this capability"""
        self.usage_count += 1
        if success:
            self.success_count += 1
    
    def evolve(self, new_stage: EvolutionStage, description: str) -> None:
        """Evolve this capability to a new stage"""
        old_stage = self.evolution_stage
        self.evolution_stage = new_stage
        self.last_evolved = datetime.now()
        
        # Record evolution history
        self.evolution_history.append({
            "timestamp": self.last_evolved.isoformat(),
            "from_stage": old_stage.value,
            "to_stage": new_stage.value,
            "description": description
        })
    
    def get_success_rate(self) -> float:
        """Get the success rate of this capability"""
        if self.usage_count == 0:
            return 0.0
        return self.success_count / self.usage_count
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert capability to dictionary"""
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "capability_type": self.capability_type.value,
            "parameters": self.parameters,
            "evolution_stage": self.evolution_stage.value,
            "created_at": self.created_at.isoformat(),
            "last_evolved": self.last_evolved.isoformat(),
            "usage_count": self.usage_count,
            "success_count": self.success_count,
            "dependencies": self.dependencies,
            "evolution_history": self.evolution_history
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Capability':
        """Create capability from dictionary"""
        capability = cls(
            name=data["name"],
            description=data["description"],
            capability_type=CapabilityType(data["capability_type"]),
            parameters=data["parameters"],
            evolution_stage=EvolutionStage(data["evolution_stage"])
        )
        capability.id = data["id"]
        capability.created_at = datetime.fromisoformat(data["created_at"])
        capability.last_evolved = datetime.fromisoformat(data["last_evolved"])
        capability.usage_count = data["usage_count"]
        capability.success_count = data["success_count"]
        capability.dependencies = data["dependencies"]
        capability.evolution_history = data["evolution_history"]
        return capability

class ScaffoldSystem:
    """Scaffold system for capability management and evolution"""
    
    def __init__(self, scaffold_file: str = "data/scaffold/agent_scaffold.json",
                capabilities_dir: str = "data/capabilities"):
        self.scaffold_file = Path(scaffold_file)
        self.scaffold_file.parent.mkdir(parents=True, exist_ok=True)
        self.capabilities_dir = Path(capabilities_dir)
        self.capabilities_dir.mkdir(parents=True, exist_ok=True)
        self.capabilities: Dict[str, Capability] = {}
        self.load_capabilities()
    
    def load_capabilities(self) -> None:
        """Load capabilities from file"""
        if not self.scaffold_file.exists():
            logger.info(f"Scaffold file {self.scaffold_file} does not exist")
            return
        
        try:
            with open(self.scaffold_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            capabilities = {}
            for capability_data in data:
                capability = Capability.from_dict(capability_data)
                capabilities[capability.name] = capability
            
            self.capabilities = capabilities
            logger.info(f"Loaded {len(self.capabilities)} capabilities")
            
            # Load capability implementations
            self._load_capability_implementations()
        except Exception as e:
            logger.error(f"Error loading capabilities: {str(e)}")
    
    def _load_capability_implementations(self) -> None:
        """Load capability implementations from files"""
        if not self.capabilities_dir.exists():
            return
        
        # Look for Python modules in capabilities directory
        for module_path in self.capabilities_dir.glob("*.py"):
            if module_path.name.startswith("__"):
                continue
            
            try:
                module_name = module_path.stem
                spec = importlib.util.spec_from_file_location(module_name, module_path)
                if spec is None or spec.loader is None:
                    continue
                
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                
                # Find capability functions in module
                for name, obj in inspect.getmembers(module):
                    if inspect.isfunction(obj) and hasattr(obj, "_capability"):
                        capability_name = obj._capability
                        if capability_name in self.capabilities:
                            self.capabilities[capability_name].function = obj
                            logger.info(f"Loaded function for capability: {capability_name}")
            except Exception as e:
                logger.error(f"Error loading capability implementation from {module_path}: {str(e)}")
    
    def save_capabilities(self) -> None:
        """Save capabilities to file"""
        try:
            with open(self.scaffold_file, "w", encoding="utf-8") as f:
                json.dump([capability.to_dict() for capability in self.capabilities.values()], f, indent=2)
            logger.info(f"Saved {len(self.capabilities)} capabilities")
        except Exception as e:
            logger.error(f"Error saving capabilities: {str(e)}")
    
    def create_capability(self, name: str, description: str, capability_type: CapabilityType,
                          parameters: List[str] = None) -> Capability:
        """Create a new capability"""
        if name in self.capabilities:
            logger.warning(f"Capability {name} already exists")
            return self.capabilities[name]
        
        capability = Capability(
            name=name,
            description=description,
            capability_type=capability_type,
            parameters=parameters or []
        )
        
        self.capabilities[name] = capability
        logger.info(f"Created capability: {name}")
        self.save_capabilities()
        
        return capability
    
    def get_capability(self, name: str) -> Optional[Capability]:
        """Get a capability by name"""
        return self.capabilities.get(name)
    
    def call_capability(self, name: str, *args, **kwargs) -> Any:
        """Call a capability function"""
        capability = self.get_capability(name)
        if not capability:
            logger.error(f"Capability {name} not found")
            return None
        
        if not capability.function:
            logger.error(f"Capability {name} has no implementation")
            return None
        
        try:
            logger.info(f"Calling capability: {name}")
            result = capability.function(*args, **kwargs)
            capability.use(success=True)
            self.save_capabilities()
            return result
        except Exception as e:
            logger.error(f"Error calling capability {name}: {str(e)}")
            capability.use(success=False)
            self.save_capabilities()
            return None
    
    def evolve_capability(self, name: str, new_stage: Optional[EvolutionStage] = None,
                         improvement_description: str = "") -> bool:
        """Evolve a capability to a new stage"""
        capability = self.get_capability(name)
        if not capability:
            logger.error(f"Capability {name} not found")
            return False
        
        # Determine new stage if not specified
        if new_stage is None:
            current_index = list(EvolutionStage).index(capability.evolution_stage)
            if current_index < len(EvolutionStage) - 1:
                new_stage = list(EvolutionStage)[current_index + 1]
            else:
                logger.warning(f"Capability {name} is already at the highest evolution stage")
                return False
        
        # Evolve the capability
        capability.evolve(new_stage, improvement_description)
        logger.info(f"Evolved capability {name} to stage {new_stage.value}")
        self.save_capabilities()
        
        return True
    
    def remove_capability(self, name: str) -> bool:
        """Remove a capability"""
        if name in self.capabilities:
            del self.capabilities[name]
            logger.info(f"Removed capability: {name}")
            self.save_capabilities()
            return True
        
        logger.warning(f"Capability {name} not found")
        return False
    
    def analyze_capabilities(self) -> Dict[str, Any]:
        """Analyze capabilities for improvement opportunities"""
        analysis = {
            "total_capabilities": len(self.capabilities),
            "by_type": {},
            "by_stage": {},
            "improvement_opportunities": [],
            "suggested_new_capabilities": [],
            "analysis": {}
        }
        
        # Count by type and stage
        for capability_type in CapabilityType:
            analysis["by_type"][capability_type.value] = 0
        
        for evolution_stage in EvolutionStage:
            analysis["by_stage"][evolution_stage.value] = 0
        
        for capability in self.capabilities.values():
            analysis["by_type"][capability.capability_type.value] += 1
            analysis["by_stage"][capability.evolution_stage.value] += 1
            
            # Check for improvement opportunities
            if capability.usage_count > 0 and capability.get_success_rate() < 0.8:
                analysis["improvement_opportunities"].append({
                    "name": capability.name,
                    "type": capability.capability_type.value,
                    "stage": capability.evolution_stage.value,
                    "success_rate": capability.get_success_rate(),
                    "usage_count": capability.usage_count,
                    "recommendation": "Improve implementation to increase success rate"
                })
            
            # Check for evolution opportunities
            if (capability.usage_count > 10 and 
                capability.evolution_stage != EvolutionStage.ADVANCED and
                capability.get_success_rate() > 0.9):
                analysis["improvement_opportunities"].append({
                    "name": capability.name,
                    "type": capability.capability_type.value,
                    "stage": capability.evolution_stage.value,
                    "success_rate": capability.get_success_rate(),
                    "usage_count": capability.usage_count,
                    "recommendation": f"Evolve to next stage: {self._get_next_stage(capability.evolution_stage).value}"
                })
        
        # Check for missing capability types
        for capability_type in CapabilityType:
            if analysis["by_type"][capability_type.value] == 0:
                analysis["suggested_new_capabilities"].append({
                    "type": capability_type.value,
                    "reason": f"No capabilities of type {capability_type.value}",
                    "suggestion": f"Create a new {capability_type.value} capability"
                })
        
        # General analysis
        analysis["analysis"] = {
            "overall_success_rate": self._calculate_overall_success_rate(),
            "most_common_type": max(analysis["by_type"], key=lambda k: analysis["by_type"][k]),
            "most_common_stage": max(analysis["by_stage"], key=lambda k: analysis["by_stage"][k]),
            "improvement_opportunities": len(analysis["improvement_opportunities"]),
            "suggested_new": len(analysis["suggested_new_capabilities"])
        }
        
        return analysis
    
    def _calculate_overall_success_rate(self) -> float:
        """Calculate overall success rate across all capabilities"""
        total_usage = sum(capability.usage_count for capability in self.capabilities.values())
        total_success = sum(capability.success_count for capability in self.capabilities.values())
        
        if total_usage == 0:
            return 0.0
        
        return total_success / total_usage
    
    def _get_next_stage(self, current_stage: EvolutionStage) -> EvolutionStage:
        """Get the next evolution stage after the current one"""
        stages = list(EvolutionStage)
        current_index = stages.index(current_stage)
        
        if current_index < len(stages) - 1:
            return stages[current_index + 1]
        return current_stage
    
    def self_evolve(self) -> List[Dict[str, Any]]:
        """Automatically evolve capabilities based on analysis"""
        analysis = self.analyze_capabilities()
        evolved = []
        
        for opportunity in analysis["improvement_opportunities"]:
            capability = self.get_capability(opportunity["name"])
            if not capability:
                continue
            
            if "Evolve to next stage" in opportunity["recommendation"]:
                next_stage = self._get_next_stage(capability.evolution_stage)
                if self.evolve_capability(capability.name, next_stage, 
                                        f"Automatic evolution based on usage ({capability.usage_count}) and success rate ({capability.get_success_rate():.2f})"):
                    evolved.append({
                        "name": capability.name,
                        "from_stage": opportunity["stage"],
                        "to_stage": next_stage.value,
                        "reason": opportunity["recommendation"]
                    })
        
        logger.info(f"Self-evolved {len(evolved)} capabilities")
        return evolved

def capability(name: str):
    """Decorator to mark a function as a capability implementation"""
    def decorator(func):
        func._capability = name
        return func
    return decorator

def create_scaffold_system(scaffold_file: str = "data/scaffold/agent_scaffold.json",
                          capabilities_dir: str = "data/capabilities") -> ScaffoldSystem:
    """Create and return a scaffold system instance"""
    return ScaffoldSystem(scaffold_file, capabilities_dir)
