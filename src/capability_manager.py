"""
Capability Manager for AI Agent in Microsoft Word

This module provides a capability manager for dynamically loading, executing,
and evolving agent capabilities. It serves as the interface between the agent
and its dynamically generated capabilities.
"""

import os
import sys
import importlib.util
import inspect
import traceback
from typing import Dict, List, Any, Optional, Union, Callable, Type

from src.scaffold_system import ScaffoldSystem, Capability, CapabilityType, EvolutionStage


class DynamicCapability:
    """Wrapper for a dynamically loaded capability"""
    
    def __init__(self, name: str, function: Callable, metadata: Dict[str, Any]):
        """
        Initialize a dynamic capability
        
        Args:
            name: Name of the capability
            function: The callable function
            metadata: Metadata about the capability
        """
        self.name = name
        self.function = function
        self.metadata = metadata
    
    def __call__(self, *args, **kwargs):
        """Call the capability function"""
        return self.function(*args, **kwargs)


class CapabilityRegistry:
    """Registry for managing and accessing capabilities"""
    
    def __init__(self, scaffold_system: ScaffoldSystem):
        """
        Initialize the capability registry
        
        Args:
            scaffold_system: The scaffold system to use
        """
        self.scaffold_system = scaffold_system
        self.dynamic_capabilities: Dict[str, DynamicCapability] = {}
        self.static_capabilities: Dict[str, Callable] = {}
    
    def register_static_capability(self, name: str, function: Callable) -> None:
        """
        Register a static capability (defined in code)
        
        Args:
            name: Name of the capability
            function: The capability function
        """
        self.static_capabilities[name] = function
    
    def load_dynamic_capability(self, name: str) -> bool:
        """
        Load a dynamic capability from the scaffold system
        
        Args:
            name: Name of the capability
            
        Returns:
            True if loaded successfully, False otherwise
        """
        capability = self.scaffold_system.get_capability(name)
        if capability is None:
            return False
        
        # Compile the capability
        if not capability.compile():
            return False
        
        # Create a wrapper function that calls the capability
        def wrapper(*args, **kwargs):
            return self.scaffold_system.call_capability(name, *args, **kwargs)
        
        # Create a dynamic capability
        self.dynamic_capabilities[name] = DynamicCapability(
            name=name,
            function=wrapper,
            metadata=capability.metadata
        )
        
        return True
    
    def get_capability(self, name: str) -> Optional[Callable]:
        """
        Get a capability by name
        
        Args:
            name: Name of the capability
            
        Returns:
            The capability function, or None if not found
        """
        # Check static capabilities first
        if name in self.static_capabilities:
            return self.static_capabilities[name]
        
        # Check if dynamic capability is already loaded
        if name in self.dynamic_capabilities:
            return self.dynamic_capabilities[name]
        
        # Try to load dynamic capability
        if self.load_dynamic_capability(name):
            return self.dynamic_capabilities[name]
        
        return None
    
    def list_capabilities(self) -> Dict[str, List[str]]:
        """
        List all available capabilities
        
        Returns:
            Dictionary with static and dynamic capability names
        """
        return {
            "static": list(self.static_capabilities.keys()),
            "dynamic": list(self.dynamic_capabilities.keys()) + [
                cap.name for cap in self.scaffold_system.capabilities.values()
                if cap.name not in self.dynamic_capabilities
            ]
        }
    
    def has_capability(self, name: str) -> bool:
        """
        Check if a capability exists
        
        Args:
            name: Name of the capability
            
        Returns:
            True if the capability exists, False otherwise
        """
        return (
            name in self.static_capabilities or
            name in self.dynamic_capabilities or
            name in self.scaffold_system.capabilities
        )


class CapabilityEvolver:
    """System for evolving capabilities based on usage patterns and feedback"""
    
    def __init__(self, scaffold_system: ScaffoldSystem):
        """
        Initialize the capability evolver
        
        Args:
            scaffold_system: The scaffold system to use
        """
        self.scaffold_system = scaffold_system
        self.evolution_strategies = {
            "error_correction": self._evolve_error_correction,
            "performance_optimization": self._evolve_performance_optimization,
            "feature_addition": self._evolve_feature_addition,
            "code_cleanup": self._evolve_code_cleanup
        }
    
    def evolve_capability(self, 
                         name: str, 
                         strategy: str, 
                         context: Optional[Dict[str, Any]] = None) -> Optional[Capability]:
        """
        Evolve a capability using a specific strategy
        
        Args:
            name: Name of the capability to evolve
            strategy: Evolution strategy to use
            context: Additional context for evolution
            
        Returns:
            The evolved capability, or None if evolution failed
        """
        if strategy not in self.evolution_strategies:
            raise ValueError(f"Unknown evolution strategy: {strategy}")
        
        capability = self.scaffold_system.get_capability(name)
        if capability is None:
            return None
        
        # Apply the evolution strategy
        evolution_func = self.evolution_strategies[strategy]
        return evolution_func(capability, context or {})
    
    def suggest_evolutions(self) -> List[Dict[str, Any]]:
        """
        Suggest capability evolutions based on analysis
        
        Returns:
            List of suggested evolutions
        """
        suggestions = []
        
        # Get capability analysis
        analysis = self.scaffold_system.analyze_capabilities()
        
        # Process improvement opportunities
        for opportunity in analysis.get("improvement_opportunities", []):
            capability_name = opportunity.get("capability")
            if not capability_name:
                continue
                
            opportunity_type = opportunity.get("type")
            
            if opportunity_type == "high_failure_rate":
                suggestions.append({
                    "capability": capability_name,
                    "strategy": "error_correction",
                    "reason": f"High failure rate ({opportunity.get('success_rate', 0):.2f})",
                    "priority": opportunity.get("priority", "medium")
                })
            
            elif opportunity_type == "stuck_in_early_stage":
                suggestions.append({
                    "capability": capability_name,
                    "strategy": "feature_addition",
                    "reason": f"Stuck in {opportunity.get('stage')} stage after {opportunity.get('version')} versions",
                    "priority": opportunity.get("priority", "medium")
                })
        
        # Look for capabilities that could benefit from optimization
        for name, capability in self.scaffold_system.capabilities.items():
            if capability.stage == EvolutionStage.STABLE and capability.use_count > 20:
                suggestions.append({
                    "capability": name,
                    "strategy": "performance_optimization",
                    "reason": f"Frequently used stable capability ({capability.use_count} uses)",
                    "priority": "low"
                })
            
            # Suggest code cleanup for complex capabilities
            if len(capability.function_code.split('\n')) > 30:
                suggestions.append({
                    "capability": name,
                    "strategy": "code_cleanup",
                    "reason": "Complex code that could benefit from cleanup",
                    "priority": "low"
                })
        
        return suggestions
    
    def _evolve_error_correction(self, 
                               capability: Capability, 
                               context: Dict[str, Any]) -> Optional[Capability]:
        """
        Evolve a capability to fix errors
        
        Args:
            capability: The capability to evolve
            context: Additional context
            
        Returns:
            The evolved capability, or None if evolution failed
        """
        # This is a placeholder for the actual implementation
        # In a real implementation, this would analyze error patterns
        # and modify the code to fix them
        
        # Add error handling to the function
        code_lines = capability.function_code.split('\n')
        
        # Find the function body indentation
        indentation = ""
        for line in code_lines:
            if line.strip().startswith("def"):
                # Get indentation from the next line
                for i, next_line in enumerate(code_lines[code_lines.index(line) + 1:]):
                    if next_line.strip():
                        indentation = next_line[:len(next_line) - len(next_line.lstrip())]
                        break
                break
        
        # Add try-except block
        new_code_lines = []
        body_started = False
        
        for line in code_lines:
            new_code_lines.append(line)
            
            if line.strip().startswith("def") and line.strip().endswith(":"):
                body_started = True
            elif body_started and line.strip() and not line.strip().startswith("#"):
                # Insert try-except after the first non-comment line in the body
                new_code_lines.append(f"{indentation}try:")
                
                # Indent the rest of the function body
                i = len(new_code_lines)
                while i < len(code_lines):
                    if code_lines[i].strip():
                        new_code_lines.append(f"{indentation}    {code_lines[i].lstrip()}")
                    else:
                        new_code_lines.append(code_lines[i])
                    i += 1
                
                # Add except block
                new_code_lines.append(f"{indentation}except Exception as e:")
                new_code_lines.append(f"{indentation}    print(f\"[ERROR] {capability.name} failed: {{str(e)}}\")")
                new_code_lines.append(f"{indentation}    return False")
                
                # We've processed all lines
                break
        
        new_code = "\n".join(new_code_lines)
        
        # Evolve the capability with the new code
        return self.scaffold_system.evolve_capability(
            name=capability.name,
            improvement_description="Added error handling to prevent failures",
            context={"strategy": "error_correction"}
        )
    
    def _evolve_performance_optimization(self, 
                                       capability: Capability, 
                                       context: Dict[str, Any]) -> Optional[Capability]:
        """
        Evolve a capability to improve performance
        
        Args:
            capability: The capability to evolve
            context: Additional context
            
        Returns:
            The evolved capability, or None if evolution failed
        """
        # This is a placeholder for the actual implementation
        # In a real implementation, this would analyze the code
        # and optimize it for performance
        
        # For now, just add a comment about optimization
        new_code = capability.function_code.replace(
            "\"\"\"",
            "\"\"\"\nOptimized for better performance\n",
            1
        )
        
        # Evolve the capability with the new code
        return self.scaffold_system.evolve_capability(
            name=capability.name,
            improvement_description="Optimized for better performance",
            context={"strategy": "performance_optimization"}
        )
    
    def _evolve_feature_addition(self, 
                               capability: Capability, 
                               context: Dict[str, Any]) -> Optional[Capability]:
        """
        Evolve a capability to add new features
        
        Args:
            capability: The capability to evolve
            context: Additional context
            
        Returns:
            The evolved capability, or None if evolution failed
        """
        # This is a placeholder for the actual implementation
        # In a real implementation, this would analyze the code
        # and add new features based on usage patterns
        
        # For now, just add a comment about new features
        new_code = capability.function_code.replace(
            "\"\"\"",
            "\"\"\"\nAdded new features for better functionality\n",
            1
        )
        
        # Evolve the capability with the new code
        return self.scaffold_system.evolve_capability(
            name=capability.name,
            improvement_description="Added new features for better functionality",
            context={"strategy": "feature_addition"}
        )
    
    def _evolve_code_cleanup(self, 
                           capability: Capability, 
                           context: Dict[str, Any]) -> Optional[Capability]:
        """
        Evolve a capability to clean up the code
        
        Args:
            capability: The capability to evolve
            context: Additional context
            
        Returns:
            The evolved capability, or None if evolution failed
        """
        # This is a placeholder for the actual implementation
        # In a real implementation, this would analyze the code
        # and clean it up for better readability and maintainability
        
        # For now, just add a comment about code cleanup
        new_code = capability.function_code.replace(
            "\"\"\"",
            "\"\"\"\nCleaned up code for better readability and maintainability\n",
            1
        )
        
        # Evolve the capability with the new code
        return self.scaffold_system.evolve_capability(
            name=capability.name,
            improvement_description="Cleaned up code for better readability and maintainability",
            context={"strategy": "code_cleanup"}
        )


class CapabilityFactory:
    """Factory for creating new capabilities"""
    
    def __init__(self, scaffold_system: ScaffoldSystem):
        """
        Initialize the capability factory
        
        Args:
            scaffold_system: The scaffold system to use
        """
        self.scaffold_system = scaffold_system
        self.templates = {
            CapabilityType.CORE: self._create_core_template,
            CapabilityType.INTERACTION: self._create_interaction_template,
            CapabilityType.ANALYSIS: self._create_analysis_template,
            CapabilityType.GENERATION: self._create_generation_template,
            CapabilityType.ADAPTATION: self._create_adaptation_template,
            CapabilityType.META: self._create_meta_template
        }
    
    def create_capability(self, 
                         name: str,
                         description: str,
                         capability_type: CapabilityType,
                         parameters: Optional[List[str]] = None,
                         dependencies: Optional[List[str]] = None,
                         code_hints: Optional[Dict[str, Any]] = None) -> Optional[Capability]:
        """
        Create a new capability
        
        Args:
            name: Name of the capability
            description: Description of what the capability does
            capability_type: Type of capability
            parameters: List of parameter names
            dependencies: Names of other capabilities this depends on
            code_hints: Hints for code generation
            
        Returns:
            The created capability, or None if creation failed
        """
        # Get the template function for this capability type
        template_func = self.templates.get(capability_type, self._create_default_template)
        
        # Generate the function code
        function_code = template_func(name, description, parameters or [], code_hints or {})
        
        # Create the capability
        try:
            return self.scaffold_system.add_capability(
                name=name,
                description=description,
                capability_type=capability_type,
                function_code=function_code,
                dependencies=dependencies,
                metadata={
                    "parameters": parameters or [],
                    "code_hints": code_hints or {}
                }
            )
        except Exception as e:
            print(f"[ERROR] Failed to create capability: {str(e)}")
            traceback.print_exc()
            return None
    
    def _create_default_template(self, 
                               name: str, 
                               description: str, 
                               parameters: List[str],
                               hints: Dict[str, Any]) -> str:
        """
        Create a default capability template
        
        Args:
            name: Capability name
            description: Capability description
            parameters: List of parameter names
            hints: Code generation hints
            
        Returns:
            Function code
        """
        # Create function signature
        params_str = ", ".join(parameters)
        signature = f"def {name}(self{', ' + params_str if params_str else ''}):"
        
        # Create function body
        body = f"""
    \"\"\"
    {description}
    \"\"\"
    print(f"[CAPABILITY] Executing {name}")
    
    # TODO: Implement {description}
    
    return True
"""
        
        return signature + body
    
    def _create_core_template(self, 
                            name: str, 
                            description: str, 
                            parameters: List[str],
                            hints: Dict[str, Any]) -> str:
        """Create a template for core capabilities"""
        # Create function signature
        params_str = ", ".join(parameters)
        signature = f"def {name}(self{', ' + params_str if params_str else ''}):"
        
        # Create function body
        body = f"""
    \"\"\"
    {description}
    
    This is a core capability that provides essential functionality.
    \"\"\"
    print(f"[CORE CAPABILITY] Executing {name}")
    
    # TODO: Implement {description}
    
    return True
"""
        
        return signature + body
    
    def _create_interaction_template(self, 
                                   name: str, 
                                   description: str, 
                                   parameters: List[str],
                                   hints: Dict[str, Any]) -> str:
        """Create a template for interaction capabilities"""
        # Create function signature
        params_str = ", ".join(parameters)
        signature = f"def {name}(self{', ' + params_str if params_str else ''}):"
        
        # Create function body
        body = f"""
    \"\"\"
    {description}
    
    This capability handles interaction with Word or other applications.
    \"\"\"
    print(f"[INTERACTION CAPABILITY] Executing {name}")
    
    # Check if Word is active
    if hasattr(self, 'check_word_active') and callable(self.check_word_active):
        if not self.check_word_active():
            print("[WARNING] Word is not active")
            return False
    
    # TODO: Implement {description}
    
    return True
"""
        
        return signature + body
    
    def _create_analysis_template(self, 
                                name: str, 
                                description: str, 
                                parameters: List[str],
                                hints: Dict[str, Any]) -> str:
        """Create a template for analysis capabilities"""
        # Create function signature
        params_str = ", ".join(parameters)
        signature = f"def {name}(self{', ' + params_str if params_str else ''}):"
        
        # Create function body
        body = f"""
    \"\"\"
    {description}
    
    This capability analyzes document content or structure.
    \"\"\"
    print(f"[ANALYSIS CAPABILITY] Executing {name}")
    
    # TODO: Implement {description}
    
    # Return analysis results
    return {{
        "success": True,
        "results": {{}},
        "timestamp": datetime.datetime.now().isoformat()
    }}
"""
        
        return signature + body
    
    def _create_generation_template(self, 
                                  name: str, 
                                  description: str, 
                                  parameters: List[str],
                                  hints: Dict[str, Any]) -> str:
        """Create a template for generation capabilities"""
        # Create function signature
        params_str = ", ".join(parameters)
        signature = f"def {name}(self{', ' + params_str if params_str else ''}):"
        
        # Create function body
        body = f"""
    \"\"\"
    {description}
    
    This capability generates content or performs creative tasks.
    \"\"\"
    print(f"[GENERATION CAPABILITY] Executing {name}")
    
    # TODO: Implement {description}
    
    # Return generated content
    return {{
        "success": True,
        "content": "Generated content will go here",
        "timestamp": datetime.datetime.now().isoformat()
    }}
"""
        
        return signature + body
    
    def _create_adaptation_template(self, 
                                  name: str, 
                                  description: str, 
                                  parameters: List[str],
                                  hints: Dict[str, Any]) -> str:
        """Create a template for adaptation capabilities"""
        # Create function signature
        params_str = ", ".join(parameters)
        signature = f"def {name}(self{', ' + params_str if params_str else ''}):"
        
        # Create function body
        body = f"""
    \"\"\"
    {description}
    
    This capability adapts to different contexts or situations.
    \"\"\"
    print(f"[ADAPTATION CAPABILITY] Executing {name}")
    
    # Get current context if available
    context = None
    if hasattr(self, 'understand_context') and callable(self.understand_context):
        context = self.understand_context()
    
    # TODO: Implement {description}
    
    return True
"""
        
        return signature + body
    
    def _create_meta_template(self, 
                            name: str, 
                            description: str, 
                            parameters: List[str],
                            hints: Dict[str, Any]) -> str:
        """Create a template for meta capabilities"""
        # Create function signature
        params_str = ", ".join(parameters)
        signature = f"def {name}(self{', ' + params_str if params_str else ''}):"
        
        # Create function body
        body = f"""
    \"\"\"
    {description}
    
    This is a meta-capability that works with other capabilities.
    \"\"\"
    print(f"[META CAPABILITY] Executing {name}")
    
    # Get available capabilities
    capabilities = None
    if hasattr(self, 'capability_registry') and hasattr(self.capability_registry, 'list_capabilities'):
        capabilities = self.capability_registry.list_capabilities()
    
    # TODO: Implement {description}
    
    return True
"""
        
        return signature + body
