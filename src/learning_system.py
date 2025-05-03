"""
Learning System for AI Agent in Microsoft Word

This module provides a learning system for the AI Agent to improve itself
by learning from mistakes and adapting its behavior over time.
"""

import datetime
import json
import os
from enum import Enum
from typing import Dict, List, Any, Optional, Union
import random
from collections import Counter, defaultdict

from src.memory_system import MemorySystem, MemoryType


class LearningType(Enum):
    """Types of learning the agent can perform"""
    ERROR_CORRECTION = "error_correction"  # Learning from errors
    BEHAVIOR_ADAPTATION = "behavior_adaptation"  # Adapting behavior based on patterns
    SKILL_IMPROVEMENT = "skill_improvement"  # Improving specific skills
    EFFICIENCY_OPTIMIZATION = "efficiency_optimization"  # Optimizing for efficiency
    CONTEXT_AWARENESS = "context_awareness"  # Improving context understanding


class Improvement:
    """Represents a specific improvement the agent has learned"""
    
    def __init__(self, 
                 description: str,
                 learning_type: LearningType,
                 trigger_pattern: str,
                 new_behavior: str,
                 confidence: float = 0.5,
                 success_count: int = 0,
                 failure_count: int = 0,
                 metadata: Optional[Dict[str, Any]] = None):
        """
        Initialize a new improvement
        
        Args:
            description: Description of the improvement
            learning_type: Type of learning
            trigger_pattern: Pattern that triggers this improvement
            new_behavior: New behavior to apply
            confidence: Confidence in this improvement (0.0 to 1.0)
            success_count: Number of successful applications
            failure_count: Number of failed applications
            metadata: Additional information
        """
        self.description = description
        self.learning_type = learning_type
        self.trigger_pattern = trigger_pattern
        self.new_behavior = new_behavior
        self.confidence = confidence
        self.success_count = success_count
        self.failure_count = failure_count
        self.metadata = metadata or {}
        self.created_at = datetime.datetime.now()
        self.last_applied = None
        self.application_count = 0
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert improvement to dictionary for serialization"""
        return {
            "description": self.description,
            "learning_type": self.learning_type.value,
            "trigger_pattern": self.trigger_pattern,
            "new_behavior": self.new_behavior,
            "confidence": self.confidence,
            "success_count": self.success_count,
            "failure_count": self.failure_count,
            "metadata": self.metadata,
            "created_at": self.created_at.isoformat(),
            "last_applied": self.last_applied.isoformat() if self.last_applied else None,
            "application_count": self.application_count
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Improvement':
        """Create improvement from dictionary"""
        improvement = cls(
            description=data["description"],
            learning_type=LearningType(data["learning_type"]),
            trigger_pattern=data["trigger_pattern"],
            new_behavior=data["new_behavior"],
            confidence=data["confidence"],
            success_count=data["success_count"],
            failure_count=data["failure_count"],
            metadata=data["metadata"]
        )
        
        improvement.created_at = datetime.datetime.fromisoformat(data["created_at"])
        if data["last_applied"]:
            improvement.last_applied = datetime.datetime.fromisoformat(data["last_applied"])
        improvement.application_count = data["application_count"]
        
        return improvement
    
    def apply(self) -> None:
        """Record that this improvement was applied"""
        self.last_applied = datetime.datetime.now()
        self.application_count += 1
    
    def record_success(self) -> None:
        """Record a successful application of this improvement"""
        self.success_count += 1
        self._update_confidence()
    
    def record_failure(self) -> None:
        """Record a failed application of this improvement"""
        self.failure_count += 1
        self._update_confidence()
    
    def _update_confidence(self) -> None:
        """Update confidence based on success and failure counts"""
        total = self.success_count + self.failure_count
        if total > 0:
            # Base confidence on success rate with a minimum of 0.1
            self.confidence = max(0.1, self.success_count / total)
            
            # Apply a decay factor for older improvements
            age_days = (datetime.datetime.now() - self.created_at).days
            if age_days > 30:  # For improvements older than 30 days
                decay_factor = 0.95 ** (age_days / 30)  # 5% decay per month
                self.confidence *= decay_factor


class LearningSystem:
    """Learning system for the AI Agent to improve from mistakes"""
    
    def __init__(self, 
                 memory_system: MemorySystem,
                 learning_file: Optional[str] = None,
                 max_improvements: int = 100,
                 verbose: bool = False):
        """
        Initialize the learning system
        
        Args:
            memory_system: The memory system to use
            learning_file: Path to file for persisting learned improvements
            max_improvements: Maximum number of improvements to store
            verbose: Whether to print debug information
        """
        self.memory_system = memory_system
        self.learning_file = learning_file or os.path.join('data', 'agent_learning.json')
        self.max_improvements = max_improvements
        self.verbose = verbose
        self.improvements: List[Improvement] = []
        
        # Error tracking
        self.error_history: List[Dict[str, Any]] = []
        self.error_patterns: Dict[str, int] = defaultdict(int)
        
        # Load existing improvements if file exists
        if learning_file and os.path.exists(learning_file):
            self.load_improvements()
    
    def add_improvement(self, 
                       description: str,
                       learning_type: LearningType,
                       trigger_pattern: str,
                       new_behavior: str,
                       confidence: float = 0.5,
                       metadata: Optional[Dict[str, Any]] = None) -> Improvement:
        """
        Add a new improvement to the system
        
        Args:
            description: Description of the improvement
            learning_type: Type of learning
            trigger_pattern: Pattern that triggers this improvement
            new_behavior: New behavior to apply
            confidence: Initial confidence (0.0 to 1.0)
            metadata: Additional information
            
        Returns:
            The created Improvement object
        """
        # Check if similar improvement already exists
        for existing in self.improvements:
            if existing.trigger_pattern == trigger_pattern:
                # Update existing improvement instead of creating a new one
                existing.new_behavior = new_behavior
                existing.description = description
                existing.confidence = (existing.confidence + confidence) / 2  # Average confidence
                if metadata:
                    existing.metadata.update(metadata)
                
                if self.verbose:
                    print(f"[LEARNING] Updated existing improvement: {description}")
                
                # Save improvements
                if self.learning_file:
                    self.save_improvements()
                
                return existing
        
        # Create new improvement
        improvement = Improvement(
            description=description,
            learning_type=learning_type,
            trigger_pattern=trigger_pattern,
            new_behavior=new_behavior,
            confidence=confidence,
            metadata=metadata
        )
        
        self.improvements.append(improvement)
        
        # If we exceed max improvements, remove least confident ones
        if len(self.improvements) > self.max_improvements:
            self._prune_improvements()
        
        # Save improvements
        if self.learning_file:
            self.save_improvements()
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Learned improvement: {description}",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "improvement": improvement.to_dict(),
                "learning_type": learning_type.value
            },
            importance=0.8  # Learning is important
        )
        
        if self.verbose:
            print(f"[LEARNING] Added new improvement: {description}")
        
        return improvement
    
    def track_error(self, 
                   operation: str, 
                   error_message: str, 
                   context: Optional[Dict[str, Any]] = None) -> None:
        """
        Track an error for pattern analysis
        
        Args:
            operation: Operation that failed
            error_message: Error message
            context: Context in which the error occurred
        """
        error_data = {
            "operation": operation,
            "error_message": error_message,
            "timestamp": datetime.datetime.now(),
            "context": context or {}
        }
        
        self.error_history.append(error_data)
        
        # Update error patterns
        error_key = f"{operation}:{error_message}"
        self.error_patterns[error_key] += 1
        
        # If this is a recurring error, try to learn from it
        if self.error_patterns[error_key] >= 3:  # Error occurred at least 3 times
            self._learn_from_error(operation, error_message, context)
    
    def _learn_from_error(self, 
                         operation: str, 
                         error_message: str, 
                         context: Optional[Dict[str, Any]] = None) -> Optional[Improvement]:
        """
        Learn from a recurring error
        
        Args:
            operation: Operation that failed
            error_message: Error message
            context: Context in which the error occurred
            
        Returns:
            The created Improvement object, if any
        """
        # Get similar errors from history
        similar_errors = [
            e for e in self.error_history 
            if e["operation"] == operation and e["error_message"] == error_message
        ]
        
        if len(similar_errors) < 3:
            return None  # Not enough data to learn
        
        # Analyze error context to find patterns
        context_keys = set()
        for error in similar_errors:
            if error["context"]:
                context_keys.update(error["context"].keys())
        
        # Find common context values
        common_context = {}
        for key in context_keys:
            values = [e["context"].get(key) for e in similar_errors if key in e["context"]]
            if len(values) >= 3:  # At least 3 occurrences
                value_counts = Counter(values)
                most_common = value_counts.most_common(1)
                if most_common and most_common[0][1] >= 2:  # Value appears at least twice
                    common_context[key] = most_common[0][0]
        
        # Generate improvement based on error type
        if operation == "Word Startup":
            return self._generate_word_startup_improvement(error_message, common_context)
        elif operation == "Typing":
            return self._generate_typing_improvement(error_message, common_context)
        elif operation == "Document Observation":
            return self._generate_observation_improvement(error_message, common_context)
        else:
            # Generic improvement
            return self._generate_generic_improvement(operation, error_message, common_context)
    
    def _generate_word_startup_improvement(self, 
                                          error_message: str, 
                                          context: Dict[str, Any]) -> Optional[Improvement]:
        """Generate improvement for Word startup errors"""
        if "not found" in error_message.lower() or "not running" in error_message.lower():
            return self.add_improvement(
                description="Improve Word startup reliability",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"Word Startup:.*{error_message}",
                new_behavior="Wait longer before interacting with Word after startup",
                confidence=0.7,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Increase startup wait time from 5 to 8 seconds"
                }
            )
        elif "blank document" in error_message.lower():
            return self.add_improvement(
                description="Improve blank document selection",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"Word Startup:.*{error_message}",
                new_behavior="Use keyboard shortcut Ctrl+N instead of clicking",
                confidence=0.8,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Use keyboard shortcut instead of mouse click"
                }
            )
        return None
    
    def _generate_typing_improvement(self, 
                                    error_message: str, 
                                    context: Dict[str, Any]) -> Optional[Improvement]:
        """Generate improvement for typing errors"""
        if "key" in error_message.lower() and "not found" in error_message.lower():
            return self.add_improvement(
                description="Improve keyboard interaction reliability",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"Typing:.*{error_message}",
                new_behavior="Add delay between key presses and use press_and_release",
                confidence=0.7,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Add 50ms delay between key presses"
                }
            )
        elif "active window" in error_message.lower():
            return self.add_improvement(
                description="Ensure Word is active before typing",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"Typing:.*{error_message}",
                new_behavior="Check if Word is active and activate it if needed",
                confidence=0.8,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Add window activation check before typing"
                }
            )
        return None
    
    def _generate_observation_improvement(self, 
                                         error_message: str, 
                                         context: Dict[str, Any]) -> Optional[Improvement]:
        """Generate improvement for document observation errors"""
        if "active" in error_message.lower():
            return self.add_improvement(
                description="Improve document observation reliability",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"Document Observation:.*{error_message}",
                new_behavior="Retry observation with increased delay",
                confidence=0.6,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Add retry with 1-second delay"
                }
            )
        return None
    
    def _generate_generic_improvement(self, 
                                     operation: str, 
                                     error_message: str, 
                                     context: Dict[str, Any]) -> Optional[Improvement]:
        """Generate generic improvement for other errors"""
        # Extract key terms from error message
        error_terms = set(error_message.lower().split())
        
        if "timeout" in error_terms or "time" in error_terms and "out" in error_terms:
            return self.add_improvement(
                description=f"Increase timeout for {operation}",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"{operation}:.*timeout",
                new_behavior=f"Double the timeout duration for {operation}",
                confidence=0.6,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Increase timeout duration"
                }
            )
        elif "permission" in error_terms or "access" in error_terms and "denied" in error_terms:
            return self.add_improvement(
                description=f"Handle permission issues in {operation}",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"{operation}:.*permission|access denied",
                new_behavior="Request elevated permissions or try alternative approach",
                confidence=0.5,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Use alternative approach with proper permissions"
                }
            )
        
        # Generic fallback improvement
        return self.add_improvement(
            description=f"Improve reliability of {operation}",
            learning_type=LearningType.ERROR_CORRECTION,
            trigger_pattern=f"{operation}:.*error",
            new_behavior=f"Add additional error checking and recovery for {operation}",
            confidence=0.4,  # Lower confidence for generic improvements
            metadata={
                "original_error": error_message,
                "context": context,
                "solution": "Add more robust error handling"
            }
        )
    
    def find_applicable_improvements(self, 
                                    operation: str, 
                                    context: Optional[Dict[str, Any]] = None) -> List[Improvement]:
        """
        Find improvements that apply to the current operation and context
        
        Args:
            operation: Current operation
            context: Current context
            
        Returns:
            List of applicable improvements
        """
        applicable = []
        
        for improvement in self.improvements:
            # Check if trigger pattern matches operation
            if operation.lower() in improvement.trigger_pattern.lower():
                # If we have context, check if it matches the improvement's metadata
                if context and improvement.metadata.get("context"):
                    # Check for context match (at least one key-value pair must match)
                    imp_context = improvement.metadata["context"]
                    if any(context.get(k) == v for k, v in imp_context.items() if k in context):
                        applicable.append(improvement)
                else:
                    # No context to match, so include based on operation match only
                    applicable.append(improvement)
        
        # Sort by confidence
        applicable.sort(key=lambda imp: imp.confidence, reverse=True)
        
        return applicable
    
    def apply_improvement(self, 
                         improvement: Improvement, 
                         success: bool = True) -> None:
        """
        Record application of an improvement
        
        Args:
            improvement: The improvement that was applied
            success: Whether the application was successful
        """
        improvement.apply()
        
        if success:
            improvement.record_success()
        else:
            improvement.record_failure()
        
        # Save improvements
        if self.learning_file:
            self.save_improvements()
        
        if self.verbose:
            result = "successfully" if success else "unsuccessfully"
            print(f"[LEARNING] Applied improvement '{improvement.description}' {result}")
    
    def analyze_behavior_patterns(self) -> List[Dict[str, Any]]:
        """
        Analyze behavior patterns to identify potential improvements
        
        Returns:
            List of behavior pattern analyses
        """
        # Get recent procedural memories
        procedural_memories = self.memory_system.get_memories(
            memory_type=MemoryType.PROCEDURAL,
            limit=100
        )
        
        # Extract actions from memories
        actions = []
        for memory in procedural_memories:
            if "action" in memory.metadata:
                actions.append(memory.metadata["action"])
        
        # Count action frequencies
        action_counts = Counter(actions)
        
        # Identify patterns
        patterns = []
        
        # Look for frequent actions that might be optimized
        for action, count in action_counts.most_common(5):
            if count >= 10:  # Action performed at least 10 times
                patterns.append({
                    "type": "frequent_action",
                    "action": action,
                    "count": count,
                    "potential_improvement": f"Optimize frequently performed action: {action}"
                })
        
        # Look for action sequences
        if len(actions) >= 3:
            for i in range(len(actions) - 2):
                sequence = tuple(actions[i:i+3])
                patterns.append({
                    "type": "action_sequence",
                    "sequence": sequence,
                    "potential_improvement": f"Combine action sequence: {' -> '.join(sequence)}"
                })
        
        return patterns
    
    def generate_behavior_improvements(self) -> List[Improvement]:
        """
        Generate improvements based on behavior pattern analysis
        
        Returns:
            List of generated improvements
        """
        patterns = self.analyze_behavior_patterns()
        improvements = []
        
        for pattern in patterns:
            if pattern["type"] == "frequent_action":
                action = pattern["action"]
                improvement = self.add_improvement(
                    description=f"Optimize frequent action: {action}",
                    learning_type=LearningType.EFFICIENCY_OPTIMIZATION,
                    trigger_pattern=f"action:{action}",
                    new_behavior=f"Use optimized implementation for {action}",
                    confidence=0.6,
                    metadata={
                        "pattern": pattern,
                        "solution": f"Create specialized method for {action}"
                    }
                )
                improvements.append(improvement)
            
            elif pattern["type"] == "action_sequence":
                sequence = pattern["sequence"]
                improvement = self.add_improvement(
                    description=f"Combine action sequence: {' -> '.join(sequence)}",
                    learning_type=LearningType.EFFICIENCY_OPTIMIZATION,
                    trigger_pattern=f"sequence:{':'.join(sequence)}",
                    new_behavior=f"Use combined implementation for sequence",
                    confidence=0.5,
                    metadata={
                        "pattern": pattern,
                        "solution": f"Create method that combines {' + '.join(sequence)}"
                    }
                )
                improvements.append(improvement)
        
        return improvements
    
    def save_improvements(self) -> bool:
        """
        Save improvements to file
        
        Returns:
            True if successful, False otherwise
        """
        if not self.learning_file:
            return False
        
        try:
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(self.learning_file), exist_ok=True)
            
            with open(self.learning_file, 'w') as f:
                json_data = {
                    "improvements": [imp.to_dict() for imp in self.improvements],
                    "error_patterns": dict(self.error_patterns)
                }
                json.dump(json_data, f, indent=2)
            
            if self.verbose:
                print(f"[LEARNING] Saved {len(self.improvements)} improvements to {self.learning_file}")
            
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to save improvements: {str(e)}")
            return False
    
    def load_improvements(self) -> bool:
        """
        Load improvements from file
        
        Returns:
            True if successful, False otherwise
        """
        if not self.learning_file or not os.path.exists(self.learning_file):
            return False
        
        try:
            with open(self.learning_file, 'r') as f:
                json_data = json.load(f)
                
                self.improvements = [
                    Improvement.from_dict(imp_data) 
                    for imp_data in json_data.get("improvements", [])
                ]
                
                self.error_patterns = defaultdict(int, json_data.get("error_patterns", {}))
            
            if self.verbose:
                print(f"[LEARNING] Loaded {len(self.improvements)} improvements from {self.learning_file}")
            
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to load improvements: {str(e)}")
            return False
    
    def _prune_improvements(self) -> int:
        """
        Remove least confident improvements to stay under max_improvements
        
        Returns:
            Number of improvements pruned
        """
        if len(self.improvements) <= self.max_improvements:
            return 0
        
        # Sort by confidence (lowest first)
        self.improvements.sort(key=lambda imp: imp.confidence)
        
        # Calculate how many to remove
        to_remove = len(self.improvements) - self.max_improvements
        
        # Remove least confident improvements
        self.improvements = self.improvements[to_remove:]
        
        if self.verbose:
            print(f"[LEARNING] Pruned {to_remove} least confident improvements")
        
        return to_remove
    
    def summarize_learning(self) -> Dict[str, Any]:
        """
        Get a summary of learning system
        
        Returns:
            Dictionary with learning statistics
        """
        if not self.improvements:
            return {
                "count": 0,
                "types": {},
                "avg_confidence": 0.0,
                "success_rate": 0.0,
                "error_patterns": len(self.error_patterns)
            }
        
        # Count by type
        type_counts = {}
        for imp in self.improvements:
            type_name = imp.learning_type.value
            type_counts[type_name] = type_counts.get(type_name, 0) + 1
        
        # Calculate average confidence
        avg_confidence = sum(imp.confidence for imp in self.improvements) / len(self.improvements)
        
        # Calculate overall success rate
        total_applications = sum(imp.application_count for imp in self.improvements)
        total_successes = sum(imp.success_count for imp in self.improvements)
        success_rate = total_successes / total_applications if total_applications > 0 else 0.0
        
        return {
            "count": len(self.improvements),
            "types": type_counts,
            "avg_confidence": avg_confidence,
            "success_rate": success_rate,
            "error_patterns": len(self.error_patterns),
            "total_applications": total_applications,
            "total_successes": total_successes
        }
