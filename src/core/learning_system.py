"""
Learning System for the Word AI Agent

Provides self-improvement capabilities based on experiences.
"""

import os
import json
import time
import re
from datetime import datetime
from typing import Dict, Any, List, Optional, Union, Set
from pathlib import Path
from enum import Enum

from src.core.logging import get_logger

logger = get_logger(__name__)

class LearningType(Enum):
    """Types of learning improvements"""
    ERROR_CORRECTION = "error_correction"  # Learning from errors
    BEHAVIOR_ADAPTATION = "behavior_adaptation"  # Adapting behavior
    SKILL_IMPROVEMENT = "skill_improvement"  # Improving skills
    EFFICIENCY_OPTIMIZATION = "efficiency_optimization"  # Optimizing performance
    CONTEXT_AWARENESS = "context_awareness"  # Improving context understanding

class Improvement:
    """Represents a learned improvement"""
    
    def __init__(self, description: str, learning_type: LearningType,
                trigger_pattern: str, new_behavior: str,
                confidence: float = 0.5, metadata: Optional[Dict[str, Any]] = None):
        self.id = f"{int(time.time() * 1000)}_{id(self)}"
        self.description = description
        self.learning_type = learning_type
        self.trigger_pattern = trigger_pattern
        self.new_behavior = new_behavior
        self.confidence = confidence  # 0.0 to 1.0
        self.created_at = datetime.now()
        self.last_applied = None
        self.application_count = 0
        self.success_count = 0
        self.metadata = metadata or {}
    
    def apply(self, success: bool = True) -> None:
        """Record an application of this improvement"""
        self.last_applied = datetime.now()
        self.application_count += 1
        if success:
            self.success_count += 1
    
    def get_success_rate(self) -> float:
        """Get the success rate of this improvement"""
        if self.application_count == 0:
            return 0.0
        return self.success_count / self.application_count
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert improvement to dictionary"""
        return {
            "id": self.id,
            "description": self.description,
            "learning_type": self.learning_type.value,
            "trigger_pattern": self.trigger_pattern,
            "new_behavior": self.new_behavior,
            "confidence": self.confidence,
            "created_at": self.created_at.isoformat(),
            "last_applied": self.last_applied.isoformat() if self.last_applied else None,
            "application_count": self.application_count,
            "success_count": self.success_count,
            "metadata": self.metadata
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
            metadata=data.get("metadata", {})
        )
        improvement.id = data["id"]
        improvement.created_at = datetime.fromisoformat(data["created_at"])
        if data["last_applied"]:
            improvement.last_applied = datetime.fromisoformat(data["last_applied"])
        improvement.application_count = data["application_count"]
        improvement.success_count = data["success_count"]
        return improvement

class LearningSystem:
    """Learning system for self-improvement"""
    
    def __init__(self, learning_file: str = "data/learning/agent_learning.json"):
        self.learning_file = Path(learning_file)
        self.learning_file.parent.mkdir(parents=True, exist_ok=True)
        self.improvements: List[Improvement] = []
        self.load_improvements()
    
    def load_improvements(self) -> None:
        """Load improvements from file"""
        if not self.learning_file.exists():
            logger.info(f"Learning file {self.learning_file} does not exist")
            return
        
        try:
            with open(self.learning_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            self.improvements = [Improvement.from_dict(improvement_data) for improvement_data in data]
            logger.info(f"Loaded {len(self.improvements)} improvements")
        except Exception as e:
            logger.error(f"Error loading improvements: {str(e)}")
    
    def save_improvements(self) -> None:
        """Save improvements to file"""
        try:
            with open(self.learning_file, "w", encoding="utf-8") as f:
                json.dump([improvement.to_dict() for improvement in self.improvements], f, indent=2)
            logger.info(f"Saved {len(self.improvements)} improvements")
        except Exception as e:
            logger.error(f"Error saving improvements: {str(e)}")
    
    def add_improvement(self, description: str, learning_type: LearningType,
                       trigger_pattern: str, new_behavior: str,
                       confidence: float = 0.5, metadata: Optional[Dict[str, Any]] = None) -> Improvement:
        """Add a new improvement"""
        improvement = Improvement(
            description=description,
            learning_type=learning_type,
            trigger_pattern=trigger_pattern,
            new_behavior=new_behavior,
            confidence=confidence,
            metadata=metadata
        )
        
        self.improvements.append(improvement)
        logger.info(f"Added improvement: {description}")
        self.save_improvements()
        
        return improvement
    
    def find_applicable_improvements(self, operation: str, context: Dict[str, Any] = None) -> List[Improvement]:
        """Find improvements applicable to an operation and context"""
        if context is None:
            context = {}
        
        applicable = []
        
        for improvement in self.improvements:
            # Check if trigger pattern matches operation
            if re.search(improvement.trigger_pattern, operation, re.IGNORECASE):
                # Check if context matches metadata requirements
                if self._context_matches_metadata(context, improvement.metadata):
                    applicable.append(improvement)
        
        # Sort by confidence and success rate
        applicable.sort(key=lambda imp: (imp.confidence, imp.get_success_rate()), reverse=True)
        
        return applicable
    
    def _context_matches_metadata(self, context: Dict[str, Any], metadata: Dict[str, Any]) -> bool:
        """Check if context matches metadata requirements"""
        if "context" not in metadata:
            # No context requirements
            return True
        
        metadata_context = metadata["context"]
        
        # Check if all required context keys match
        for key, value in metadata_context.items():
            if key not in context or context[key] != value:
                return False
        
        return True
    
    def apply_improvement(self, improvement_id: str, success: bool = True) -> bool:
        """Record application of an improvement"""
        for improvement in self.improvements:
            if improvement.id == improvement_id:
                improvement.apply(success)
                logger.info(f"Applied improvement: {improvement.description}")
                self.save_improvements()
                return True
        
        logger.warning(f"Improvement not found: {improvement_id}")
        return False
    
    def remove_improvement(self, improvement_id: str) -> bool:
        """Remove an improvement"""
        for i, improvement in enumerate(self.improvements):
            if improvement.id == improvement_id:
                del self.improvements[i]
                logger.info(f"Removed improvement: {improvement.description}")
                self.save_improvements()
                return True
        
        logger.warning(f"Improvement not found: {improvement_id}")
        return False
    
    def update_confidence(self, improvement_id: str, new_confidence: float) -> bool:
        """Update confidence score for an improvement"""
        for improvement in self.improvements:
            if improvement.id == improvement_id:
                improvement.confidence = max(0.0, min(1.0, new_confidence))
                logger.info(f"Updated confidence for improvement: {improvement.description}")
                self.save_improvements()
                return True
        
        logger.warning(f"Improvement not found: {improvement_id}")
        return False
    
    def analyze_errors(self, error_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Analyze errors and suggest improvements"""
        suggestions = []
        
        # Extract error information
        error_type = error_data.get("type", "unknown")
        error_message = error_data.get("message", "")
        operation = error_data.get("operation", "unknown")
        context = error_data.get("context", {})
        
        # Check if we already have improvements for this error
        existing = self.find_applicable_improvements(operation, context)
        
        if existing:
            # We already have improvements, but they might not be working
            for improvement in existing:
                if improvement.get_success_rate() < 0.5:
                    # This improvement isn't very successful
                    suggestions.append({
                        "type": "update",
                        "improvement_id": improvement.id,
                        "description": improvement.description,
                        "current_behavior": improvement.new_behavior,
                        "suggested_behavior": self._generate_alternative_behavior(improvement),
                        "reason": f"Low success rate: {improvement.get_success_rate():.2f}"
                    })
        else:
            # No existing improvements, suggest new ones
            if "timeout" in error_message.lower():
                # Timeout error
                suggestions.append({
                    "type": "new",
                    "learning_type": LearningType.EFFICIENCY_OPTIMIZATION,
                    "description": "Increase timeout for operations",
                    "trigger_pattern": f"{operation}.*",
                    "new_behavior": "Increase timeout by 50%",
                    "confidence": 0.7,
                    "metadata": {
                        "context": context,
                        "error_type": error_type
                    }
                })
            elif "permission" in error_message.lower():
                # Permission error
                suggestions.append({
                    "type": "new",
                    "learning_type": LearningType.ERROR_CORRECTION,
                    "description": "Request appropriate permissions",
                    "trigger_pattern": f"{operation}.*",
                    "new_behavior": "Check and request permissions before operation",
                    "confidence": 0.8,
                    "metadata": {
                        "context": context,
                        "error_type": error_type
                    }
                })
            else:
                # Generic error
                suggestions.append({
                    "type": "new",
                    "learning_type": LearningType.ERROR_CORRECTION,
                    "description": f"Handle {error_type} errors in {operation}",
                    "trigger_pattern": f"{operation}.*",
                    "new_behavior": "Add error handling and retry logic",
                    "confidence": 0.6,
                    "metadata": {
                        "context": context,
                        "error_type": error_type
                    }
                })
        
        return suggestions
    
    def _generate_alternative_behavior(self, improvement: Improvement) -> str:
        """Generate an alternative behavior for an improvement that isn't working well"""
        # This would ideally use more sophisticated logic
        if "retry" in improvement.new_behavior.lower():
            return improvement.new_behavior.replace("retry", "retry with backoff")
        elif "increase" in improvement.new_behavior.lower():
            # Extract numbers and increase them
            numbers = re.findall(r'\d+', improvement.new_behavior)
            if numbers:
                for number in numbers:
                    improved_number = int(number) * 2
                    return improvement.new_behavior.replace(number, str(improved_number))
            return improvement.new_behavior + " by a larger amount"
        else:
            return "Try an alternative approach: " + improvement.new_behavior
    
    def summarize_learning(self) -> Dict[str, Any]:
        """Summarize learning statistics"""
        if not self.improvements:
            return {
                "count": 0,
                "success_rate": 0.0,
                "by_type": {},
                "top_improvements": []
            }
        
        # Calculate overall statistics
        total_applications = sum(imp.application_count for imp in self.improvements)
        total_successes = sum(imp.success_count for imp in self.improvements)
        success_rate = total_successes / total_applications if total_applications > 0 else 0.0
        
        # Group by type
        by_type = {}
        for learning_type in LearningType:
            type_improvements = [imp for imp in self.improvements if imp.learning_type == learning_type]
            type_applications = sum(imp.application_count for imp in type_improvements)
            type_successes = sum(imp.success_count for imp in type_improvements)
            type_success_rate = type_successes / type_applications if type_applications > 0 else 0.0
            
            by_type[learning_type.value] = {
                "count": len(type_improvements),
                "applications": type_applications,
                "successes": type_successes,
                "success_rate": type_success_rate
            }
        
        # Find top improvements by success rate
        active_improvements = [imp for imp in self.improvements if imp.application_count > 0]
        active_improvements.sort(key=lambda imp: imp.get_success_rate() * imp.confidence, reverse=True)
        top_improvements = [{
            "id": imp.id,
            "description": imp.description,
            "type": imp.learning_type.value,
            "success_rate": imp.get_success_rate(),
            "applications": imp.application_count,
            "confidence": imp.confidence
        } for imp in active_improvements[:5]]  # Top 5
        
        return {
            "count": len(self.improvements),
            "applications": total_applications,
            "successes": total_successes,
            "success_rate": success_rate,
            "by_type": by_type,
            "top_improvements": top_improvements
        }

def create_learning_system(learning_file: str = "data/learning/agent_learning.json") -> LearningSystem:
    """Create and return a learning system instance"""
    return LearningSystem(learning_file)
