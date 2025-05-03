"""
Scaffold System for AI Agent in Microsoft Word

This module provides a scaffold system for the AI Agent to evolve and improve itself
over time, independent of the underlying LLM. It enables the agent to develop new
capabilities, strategies, and behaviors through experience.
"""

import datetime
import json
import os
import importlib
import inspect
import re
import sys
import traceback
from enum import Enum
from typing import Dict, List, Any, Optional, Union, Callable, Type
import random
from collections import Counter, defaultdict

from src.memory_system import MemorySystem, MemoryType


class CapabilityType(Enum):
    """Types of capabilities the agent can develop"""
    CORE = "core"                # Core capabilities (essential functions)
    INTERACTION = "interaction"  # Interaction with Word and other applications
    ANALYSIS = "analysis"        # Analysis of document content and structure
    GENERATION = "generation"    # Content generation capabilities
    ADAPTATION = "adaptation"    # Adaptation to different contexts
    META = "meta"                # Meta-capabilities (capabilities about capabilities)


class EvolutionStage(Enum):
    """Stages of capability evolution"""
    CONCEPTION = "conception"    # Initial idea for a capability
    PROTOTYPE = "prototype"      # Basic implementation, may have issues
    STABLE = "stable"            # Reliable implementation
    OPTIMIZED = "optimized"      # Performance-optimized implementation
    ADVANCED = "advanced"        # Advanced implementation with extra features


class Capability:
    """Represents a specific capability the agent has developed"""
    
    def __init__(self, 
                 name: str,
                 description: str,
                 capability_type: CapabilityType,
                 function_code: str,
                 stage: EvolutionStage = EvolutionStage.CONCEPTION,
                 success_count: int = 0,
                 failure_count: int = 0,
                 version: int = 1,
                 dependencies: Optional[List[str]] = None,
                 metadata: Optional[Dict[str, Any]] = None):
        """
        Initialize a new capability
        
        Args:
            name: Name of the capability (used as function name)
            description: Description of what the capability does
            capability_type: Type of capability
            function_code: Python code implementing the capability
            stage: Current evolution stage
            success_count: Number of successful uses
            failure_count: Number of failed uses
            version: Version number
            dependencies: Names of other capabilities this depends on
            metadata: Additional information
        """
        self.name = name
        self.description = description
        self.capability_type = capability_type
        self.function_code = function_code
        self.stage = stage
        self.success_count = success_count
        self.failure_count = failure_count
        self.version = version
        self.dependencies = dependencies or []
        self.metadata = metadata or {}
        self.created_at = datetime.datetime.now()
        self.last_modified = self.created_at
        self.last_used = None
        self.use_count = 0
        self._function = None  # Compiled function object
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert capability to dictionary for serialization"""
        return {
            "name": self.name,
            "description": self.description,
            "capability_type": self.capability_type.value,
            "function_code": self.function_code,
            "stage": self.stage.value,
            "success_count": self.success_count,
            "failure_count": self.failure_count,
            "version": self.version,
            "dependencies": self.dependencies,
            "metadata": self.metadata,
            "created_at": self.created_at.isoformat(),
            "last_modified": self.last_modified.isoformat(),
            "last_used": self.last_used.isoformat() if self.last_used else None,
            "use_count": self.use_count
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Capability':
        """Create capability from dictionary"""
        capability = cls(
            name=data["name"],
            description=data["description"],
            capability_type=CapabilityType(data["capability_type"]),
            function_code=data["function_code"],
            stage=EvolutionStage(data["stage"]),
            success_count=data["success_count"],
            failure_count=data["failure_count"],
            version=data["version"],
            dependencies=data["dependencies"],
            metadata=data["metadata"]
        )
        
        capability.created_at = datetime.datetime.fromisoformat(data["created_at"])
        capability.last_modified = datetime.datetime.fromisoformat(data["last_modified"])
        if data["last_used"]:
            capability.last_used = datetime.datetime.fromisoformat(data["last_used"])
        capability.use_count = data["use_count"]
        
        return capability
    
    def compile(self) -> bool:
        """Compile the function code into a callable function"""
        try:
            # Create a namespace for the function
            namespace = {}
            
            # Add necessary imports
            exec("import time, random, datetime, os, sys", namespace)
            
            # Compile and execute the function code
            exec(self.function_code, namespace)
            
            # Get the function from the namespace
            self._function = namespace.get(self.name)
            
            if not callable(self._function):
                raise ValueError(f"Function '{self.name}' not found in compiled code")
            
            return True
        except Exception as e:
            print(f"[ERROR] Failed to compile capability '{self.name}': {str(e)}")
            traceback.print_exc()
            self._function = None
            return False
    
    def call(self, *args, **kwargs) -> Any:
        """Call the capability function with arguments"""
        if self._function is None and not self.compile():
            raise ValueError(f"Cannot call capability '{self.name}': compilation failed")
        
        self.use_count += 1
        self.last_used = datetime.datetime.now()
        
        try:
            result = self._function(*args, **kwargs)
            self.success_count += 1
            return result
        except Exception as e:
            self.failure_count += 1
            raise e
    
    def evolve(self, new_code: str, description: Optional[str] = None) -> 'Capability':
        """
        Evolve the capability with new code
        
        Args:
            new_code: New implementation code
            description: Updated description (optional)
            
        Returns:
            New evolved capability
        """
        # Determine new stage based on current stage and success rate
        new_stage = self.stage
        success_rate = self.success_count / max(1, self.success_count + self.failure_count)
        
        if success_rate > 0.9 and self.use_count > 10:
            # Promote to next stage if very successful
            if self.stage == EvolutionStage.CONCEPTION:
                new_stage = EvolutionStage.PROTOTYPE
            elif self.stage == EvolutionStage.PROTOTYPE:
                new_stage = EvolutionStage.STABLE
            elif self.stage == EvolutionStage.STABLE:
                new_stage = EvolutionStage.OPTIMIZED
            elif self.stage == EvolutionStage.OPTIMIZED:
                new_stage = EvolutionStage.ADVANCED
        
        # Create evolved capability
        evolved = Capability(
            name=self.name,
            description=description or self.description,
            capability_type=self.capability_type,
            function_code=new_code,
            stage=new_stage,
            success_count=0,  # Reset counts for new version
            failure_count=0,
            version=self.version + 1,
            dependencies=self.dependencies.copy(),
            metadata=self.metadata.copy()
        )
        
        # Copy creation date from original
        evolved.created_at = self.created_at
        
        return evolved


class ScaffoldSystem:
    """Scaffold system for the AI Agent to evolve and improve itself"""
    
    def __init__(self, 
                 memory_system: MemorySystem,
                 scaffold_file: Optional[str] = None,
                 capabilities_dir: Optional[str] = None,
                 max_capabilities: int = 100,
                 verbose: bool = False):
        """
        Initialize the scaffold system
        
        Args:
            memory_system: The memory system to use
            scaffold_file: Path to file for persisting scaffold data
            capabilities_dir: Directory to store capability modules
            max_capabilities: Maximum number of capabilities to store
            verbose: Whether to print debug information
        """
        self.memory_system = memory_system
        self.scaffold_file = scaffold_file or os.path.join('data', 'agent_scaffold.json')
        self.capabilities_dir = capabilities_dir or os.path.join('data', 'capabilities')
        self.max_capabilities = max_capabilities
        self.verbose = verbose
        self.capabilities: Dict[str, Capability] = {}
        
        # Evolution tracking
        self.evolution_history: List[Dict[str, Any]] = []
        
        # Ensure capabilities directory exists
        os.makedirs(self.capabilities_dir, exist_ok=True)
        
        # Load existing scaffold data if file exists
        if scaffold_file and os.path.exists(scaffold_file):
            self.load_scaffold()
    
    def add_capability(self, 
                      name: str,
                      description: str,
                      capability_type: CapabilityType,
                      function_code: str,
                      stage: EvolutionStage = EvolutionStage.CONCEPTION,
                      dependencies: Optional[List[str]] = None,
                      metadata: Optional[Dict[str, Any]] = None) -> Capability:
        """
        Add a new capability to the system
        
        Args:
            name: Name of the capability (used as function name)
            description: Description of what the capability does
            capability_type: Type of capability
            function_code: Python code implementing the capability
            stage: Current evolution stage
            dependencies: Names of other capabilities this depends on
            metadata: Additional information
            
        Returns:
            The created Capability object
        """
        # Validate function code
        if not self._validate_function_code(name, function_code):
            raise ValueError(f"Invalid function code for capability '{name}'")
        
        # Check if capability with this name already exists
        if name in self.capabilities:
            # Evolve existing capability instead of creating new one
            existing = self.capabilities[name]
            evolved = existing.evolve(function_code, description)
            self.capabilities[name] = evolved
            
            # Record evolution in history
            self.evolution_history.append({
                "capability": name,
                "from_version": existing.version,
                "to_version": evolved.version,
                "timestamp": datetime.datetime.now(),
                "from_stage": existing.stage.value,
                "to_stage": evolved.stage.value
            })
            
            if self.verbose:
                print(f"[SCAFFOLD] Evolved capability: {name} (v{evolved.version}, {evolved.stage.value})")
            
            # Save to memory
            self.memory_system.add_memory(
                content=f"Evolved capability: {name} (v{evolved.version}, {evolved.stage.value})",
                memory_type=MemoryType.LEARNING,
                metadata={
                    "capability": evolved.to_dict(),
                    "evolution": self.evolution_history[-1]
                },
                importance=0.8
            )
            
            # Save scaffold
            self.save_scaffold()
            
            return evolved
        
        # Create new capability
        capability = Capability(
            name=name,
            description=description,
            capability_type=capability_type,
            function_code=function_code,
            stage=stage,
            dependencies=dependencies,
            metadata=metadata
        )
        
        self.capabilities[name] = capability
        
        # If we exceed max capabilities, remove least used ones
        if len(self.capabilities) > self.max_capabilities:
            self._prune_capabilities()
        
        # Save scaffold
        self.save_scaffold()
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Added new capability: {name} ({capability_type.value})",
            memory_type=MemoryType.LEARNING,
            metadata={
                "capability": capability.to_dict(),
                "capability_type": capability_type.value
            },
            importance=0.9  # New capabilities are important
        )
        
        if self.verbose:
            print(f"[SCAFFOLD] Added new capability: {name} ({capability_type.value})")
        
        return capability
    
    def get_capability(self, name: str) -> Optional[Capability]:
        """
        Get a capability by name
        
        Args:
            name: Name of the capability
            
        Returns:
            The Capability object, or None if not found
        """
        return self.capabilities.get(name)
    
    def call_capability(self, name: str, *args, **kwargs) -> Any:
        """
        Call a capability by name
        
        Args:
            name: Name of the capability
            *args: Positional arguments to pass to the capability
            **kwargs: Keyword arguments to pass to the capability
            
        Returns:
            The result of the capability function
            
        Raises:
            ValueError: If the capability does not exist
        """
        capability = self.get_capability(name)
        if capability is None:
            raise ValueError(f"Capability '{name}' not found")
        
        try:
            result = capability.call(*args, **kwargs)
            
            # Record successful use in memory
            self.memory_system.add_memory(
                content=f"Successfully used capability: {name}",
                memory_type=MemoryType.PROCEDURAL,
                metadata={
                    "capability": name,
                    "args": str(args),
                    "kwargs": str(kwargs),
                    "success": True
                },
                importance=0.4  # Lower importance for routine usage
            )
            
            return result
        except Exception as e:
            # Record failure in memory
            self.memory_system.add_memory(
                content=f"Failed to use capability: {name} - {str(e)}",
                memory_type=MemoryType.PROCEDURAL,
                metadata={
                    "capability": name,
                    "args": str(args),
                    "kwargs": str(kwargs),
                    "error": str(e),
                    "success": False
                },
                importance=0.7  # Higher importance for failures
            )
            
            raise e
    
    def find_capabilities(self, 
                         capability_type: Optional[CapabilityType] = None,
                         stage: Optional[EvolutionStage] = None,
                         min_success_rate: float = 0.0) -> List[Capability]:
        """
        Find capabilities matching criteria
        
        Args:
            capability_type: Filter by capability type
            stage: Filter by evolution stage
            min_success_rate: Minimum success rate (0.0 to 1.0)
            
        Returns:
            List of matching Capability objects
        """
        results = []
        
        for capability in self.capabilities.values():
            # Apply filters
            if capability_type and capability.capability_type != capability_type:
                continue
                
            if stage and capability.stage != stage:
                continue
                
            # Calculate success rate
            total = capability.success_count + capability.failure_count
            success_rate = capability.success_count / total if total > 0 else 0.0
            
            if success_rate < min_success_rate:
                continue
                
            results.append(capability)
        
        # Sort by success rate (highest first)
        results.sort(key=lambda c: c.success_count / max(1, c.success_count + c.failure_count), reverse=True)
        
        return results
    
    def generate_capability(self, 
                           description: str,
                           capability_type: CapabilityType,
                           dependencies: Optional[List[str]] = None,
                           context: Optional[Dict[str, Any]] = None) -> Optional[Capability]:
        """
        Generate a new capability based on description
        
        Args:
            description: Description of what the capability should do
            capability_type: Type of capability
            dependencies: Names of other capabilities this depends on
            context: Additional context for generation
            
        Returns:
            The generated Capability object, or None if generation failed
        """
        # This is a placeholder for the actual implementation
        # In a real implementation, this would use the LLM to generate code
        # based on the description and context
        
        # For now, we'll use a simple template-based approach
        name = self._generate_capability_name(description)
        
        # Generate function signature
        signature = f"def {name}(self"
        
        # Add parameters based on context
        if context and "parameters" in context:
            for param in context["parameters"]:
                signature += f", {param}"
        
        signature += "):"
        
        # Generate function body
        body = f"""
    \"\"\"
    {description}
    \"\"\"
    # TODO: Implement {description}
    print(f"[CAPABILITY] Executing {name}")
    return True
"""
        
        # Combine signature and body
        function_code = signature + body
        
        try:
            # Add the capability
            return self.add_capability(
                name=name,
                description=description,
                capability_type=capability_type,
                function_code=function_code,
                dependencies=dependencies,
                metadata=context
            )
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to generate capability: {str(e)}")
            return None
    
    def evolve_capability(self, 
                         name: str,
                         improvement_description: str,
                         context: Optional[Dict[str, Any]] = None) -> Optional[Capability]:
        """
        Evolve an existing capability with improvements
        
        Args:
            name: Name of the capability to evolve
            improvement_description: Description of the improvements to make
            context: Additional context for evolution
            
        Returns:
            The evolved Capability object, or None if evolution failed
        """
        capability = self.get_capability(name)
        if capability is None:
            if self.verbose:
                print(f"[ERROR] Cannot evolve: capability '{name}' not found")
            return None
        
        # This is a placeholder for the actual implementation
        # In a real implementation, this would use the LLM to evolve the code
        # based on the improvement description and context
        
        # For now, we'll use a simple approach to modify the existing code
        new_code = capability.function_code
        
        # Add a comment with the improvement description
        new_code = new_code.replace("\"\"\"", f"\"\"\"\nImproved: {improvement_description}\n", 1)
        
        # Add a print statement to show the improvement
        new_code = new_code.replace("print(f\"[CAPABILITY]", f"print(f\"[IMPROVED CAPABILITY]", 1)
        
        try:
            # Update the capability
            return self.add_capability(
                name=name,
                description=capability.description,
                capability_type=capability.capability_type,
                function_code=new_code,
                dependencies=capability.dependencies,
                metadata={**capability.metadata, "improvement": improvement_description}
            )
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to evolve capability: {str(e)}")
            return None
    
    def analyze_capabilities(self) -> Dict[str, Any]:
        """
        Analyze capabilities to identify patterns and improvement opportunities
        
        Returns:
            Dictionary with analysis results
        """
        if not self.capabilities:
            return {
                "count": 0,
                "types": {},
                "stages": {},
                "success_rate": 0.0,
                "improvement_opportunities": []
            }
        
        # Count by type and stage
        type_counts = {}
        stage_counts = {}
        
        for capability in self.capabilities.values():
            # Count by type
            type_name = capability.capability_type.value
            type_counts[type_name] = type_counts.get(type_name, 0) + 1
            
            # Count by stage
            stage_name = capability.stage.value
            stage_counts[stage_name] = stage_counts.get(stage_name, 0) + 1
        
        # Calculate overall success rate
        total_successes = sum(c.success_count for c in self.capabilities.values())
        total_failures = sum(c.failure_count for c in self.capabilities.values())
        total_attempts = total_successes + total_failures
        success_rate = total_successes / total_attempts if total_attempts > 0 else 0.0
        
        # Identify improvement opportunities
        improvement_opportunities = []
        
        for name, capability in self.capabilities.items():
            # Check for capabilities with high failure rates
            total = capability.success_count + capability.failure_count
            if total >= 5:  # Only consider capabilities with sufficient usage
                success_rate = capability.success_count / total
                if success_rate < 0.7:  # Less than 70% success rate
                    improvement_opportunities.append({
                        "capability": name,
                        "type": "high_failure_rate",
                        "success_rate": success_rate,
                        "priority": "high" if success_rate < 0.5 else "medium"
                    })
            
            # Check for capabilities stuck in early stages
            if capability.stage in [EvolutionStage.CONCEPTION, EvolutionStage.PROTOTYPE] and capability.version > 2:
                improvement_opportunities.append({
                    "capability": name,
                    "type": "stuck_in_early_stage",
                    "stage": capability.stage.value,
                    "version": capability.version,
                    "priority": "medium"
                })
        
        return {
            "count": len(self.capabilities),
            "types": type_counts,
            "stages": stage_counts,
            "success_rate": success_rate,
            "improvement_opportunities": improvement_opportunities
        }
    
    def save_scaffold(self) -> bool:
        """
        Save scaffold data to file
        
        Returns:
            True if successful, False otherwise
        """
        if not self.scaffold_file:
            return False
        
        try:
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(self.scaffold_file), exist_ok=True)
            
            with open(self.scaffold_file, 'w') as f:
                json_data = {
                    "capabilities": {name: cap.to_dict() for name, cap in self.capabilities.items()},
                    "evolution_history": self.evolution_history
                }
                json.dump(json_data, f, indent=2)
            
            if self.verbose:
                print(f"[SCAFFOLD] Saved {len(self.capabilities)} capabilities to {self.scaffold_file}")
            
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to save scaffold data: {str(e)}")
            return False
    
    def load_scaffold(self) -> bool:
        """
        Load scaffold data from file
        
        Returns:
            True if successful, False otherwise
        """
        if not self.scaffold_file or not os.path.exists(self.scaffold_file):
            return False
        
        try:
            with open(self.scaffold_file, 'r') as f:
                json_data = json.load(f)
                
                self.capabilities = {
                    name: Capability.from_dict(cap_data) 
                    for name, cap_data in json_data.get("capabilities", {}).items()
                }
                
                self.evolution_history = json_data.get("evolution_history", [])
            
            if self.verbose:
                print(f"[SCAFFOLD] Loaded {len(self.capabilities)} capabilities from {self.scaffold_file}")
            
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to load scaffold data: {str(e)}")
            return False
    
    def _prune_capabilities(self) -> int:
        """
        Remove least used capabilities to stay under max_capabilities
        
        Returns:
            Number of capabilities pruned
        """
        if len(self.capabilities) <= self.max_capabilities:
            return 0
        
        # Sort by use count (lowest first)
        sorted_capabilities = sorted(
            self.capabilities.items(),
            key=lambda item: item[1].use_count
        )
        
        # Calculate how many to remove
        to_remove = len(self.capabilities) - self.max_capabilities
        
        # Remove least used capabilities
        for name, _ in sorted_capabilities[:to_remove]:
            del self.capabilities[name]
        
        if self.verbose:
            print(f"[SCAFFOLD] Pruned {to_remove} least used capabilities")
        
        return to_remove
    
    def _validate_function_code(self, name: str, code: str) -> bool:
        """
        Validate that the function code is safe and properly formatted
        
        Args:
            name: Expected function name
            code: Function code to validate
            
        Returns:
            True if valid, False otherwise
        """
        # Check if code contains the function definition
        if not re.search(rf"def\s+{re.escape(name)}\s*\(", code):
            if self.verbose:
                print(f"[ERROR] Function '{name}' not found in code")
            return False
        
        # Check for potentially unsafe operations
        unsafe_patterns = [
            r"__import__\s*\(",
            r"eval\s*\(",
            r"exec\s*\(",
            r"globals\s*\(",
            r"locals\s*\(",
            r"getattr\s*\(",
            r"setattr\s*\(",
            r"delattr\s*\(",
            r"open\s*\(.+,\s*['\"](w|a)",  # Writing to files
            r"os\.(system|popen|spawn|exec)",
            r"subprocess\.(call|run|Popen)",
            r"importlib",
        ]
        
        for pattern in unsafe_patterns:
            if re.search(pattern, code):
                if self.verbose:
                    print(f"[ERROR] Unsafe pattern found in code: {pattern}")
                return False
        
        # Try to compile the code to check for syntax errors
        try:
            compile(code, "<string>", "exec")
            return True
        except SyntaxError as e:
            if self.verbose:
                print(f"[ERROR] Syntax error in code: {str(e)}")
            return False
    
    def _generate_capability_name(self, description: str) -> str:
        """
        Generate a function name from a description
        
        Args:
            description: Capability description
            
        Returns:
            A valid function name
        """
        # Convert description to lowercase and remove non-alphanumeric characters
        name = re.sub(r'[^a-z0-9\s]', '', description.lower())
        
        # Split into words and join with underscores
        words = name.split()
        if not words:
            return "unnamed_capability"
        
        # Use first 4 words maximum
        name = '_'.join(words[:4])
        
        # Ensure name starts with a letter
        if not name[0].isalpha():
            name = 'capability_' + name
        
        # Ensure name is unique
        base_name = name
        counter = 1
        while name in self.capabilities:
            name = f"{base_name}_{counter}"
            counter += 1
        
        return name


class CapabilityManager:
    """Manager for dynamically loading and using capabilities"""
    
    def __init__(self, scaffold_system: ScaffoldSystem):
        """
        Initialize the capability manager
        
        Args:
            scaffold_system: The scaffold system to use
        """
        self.scaffold_system = scaffold_system
        self.loaded_capabilities = {}
    
    def load_capability(self, name: str) -> bool:
        """
        Load a capability into the manager
        
        Args:
            name: Name of the capability to load
            
        Returns:
            True if loaded successfully, False otherwise
        """
        capability = self.scaffold_system.get_capability(name)
        if capability is None:
            return False
        
        # Compile the capability
        if not capability.compile():
            return False
        
        self.loaded_capabilities[name] = capability
        return True
    
    def call(self, name: str, *args, **kwargs) -> Any:
        """
        Call a capability
        
        Args:
            name: Name of the capability to call
            *args: Positional arguments
            **kwargs: Keyword arguments
            
        Returns:
            Result of the capability function
        """
        # Load the capability if not already loaded
        if name not in self.loaded_capabilities and not self.load_capability(name):
            raise ValueError(f"Failed to load capability '{name}'")
        
        # Call the capability
        return self.scaffold_system.call_capability(name, *args, **kwargs)
    
    def has_capability(self, name: str) -> bool:
        """
        Check if a capability exists
        
        Args:
            name: Name of the capability
            
        Returns:
            True if the capability exists, False otherwise
        """
        return name in self.scaffold_system.capabilities
    
    def get_capabilities_by_type(self, capability_type: CapabilityType) -> List[str]:
        """
        Get names of capabilities of a specific type
        
        Args:
            capability_type: Type of capabilities to get
            
        Returns:
            List of capability names
        """
        capabilities = self.scaffold_system.find_capabilities(capability_type=capability_type)
        return [cap.name for cap in capabilities]
