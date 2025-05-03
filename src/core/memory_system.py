"""
Memory System for the Word AI Agent

Provides persistent memory capabilities for maintaining context awareness.
"""

import os
import json
import time
from datetime import datetime
from typing import Dict, Any, List, Optional, Union
from pathlib import Path
from enum import Enum

from src.core.logging import get_logger

logger = get_logger(__name__)

class MemoryType(Enum):
    """Types of memories the system can store"""
    SPATIAL = "spatial"  # Location in document
    TEMPORAL = "temporal"  # Time-based memories
    CONTEXTUAL = "contextual"  # Content understanding
    PROCEDURAL = "procedural"  # How to perform tasks
    DOCUMENT = "document"  # Document structure
    LEARNING = "learning"  # Improvements

class Memory:
    """Represents a single memory"""
    
    def __init__(self, content: str, memory_type: MemoryType, importance: float = 0.5,
                metadata: Optional[Dict[str, Any]] = None):
        self.id = f"{int(time.time() * 1000)}_{id(self)}"
        self.content = content
        self.memory_type = memory_type
        self.importance = importance  # 0.0 to 1.0
        self.created_at = datetime.now()
        self.accessed_at = datetime.now()
        self.access_count = 0
        self.metadata = metadata or {}
    
    def access(self) -> None:
        """Record memory access"""
        self.accessed_at = datetime.now()
        self.access_count += 1
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert memory to dictionary"""
        return {
            "id": self.id,
            "content": self.content,
            "memory_type": self.memory_type.value,
            "importance": self.importance,
            "created_at": self.created_at.isoformat(),
            "accessed_at": self.accessed_at.isoformat(),
            "access_count": self.access_count,
            "metadata": self.metadata
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Memory':
        """Create memory from dictionary"""
        memory = cls(
            content=data["content"],
            memory_type=MemoryType(data["memory_type"]),
            importance=data["importance"],
            metadata=data.get("metadata", {})
        )
        memory.id = data["id"]
        memory.created_at = datetime.fromisoformat(data["created_at"])
        memory.accessed_at = datetime.fromisoformat(data["accessed_at"])
        memory.access_count = data["access_count"]
        return memory

class MemorySystem:
    """Memory system for maintaining context awareness"""
    
    def __init__(self, memory_file: str = "data/memory/agent_memory.json"):
        self.memory_file = Path(memory_file)
        self.memory_file.parent.mkdir(parents=True, exist_ok=True)
        self.memories: List[Memory] = []
        self.load_memories()
    
    def load_memories(self) -> None:
        """Load memories from file"""
        if not self.memory_file.exists():
            logger.info(f"Memory file {self.memory_file} does not exist")
            return
        
        try:
            with open(self.memory_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            self.memories = [Memory.from_dict(memory_data) for memory_data in data]
            logger.info(f"Loaded {len(self.memories)} memories")
        except Exception as e:
            logger.error(f"Error loading memories: {str(e)}")
    
    def save_memories(self) -> None:
        """Save memories to file"""
        try:
            with open(self.memory_file, "w", encoding="utf-8") as f:
                json.dump([memory.to_dict() for memory in self.memories], f, indent=2)
            logger.info(f"Saved {len(self.memories)} memories")
        except Exception as e:
            logger.error(f"Error saving memories: {str(e)}")
    
    def add_memory(self, content: str, memory_type: MemoryType, importance: float = 0.5,
                  metadata: Optional[Dict[str, Any]] = None) -> Memory:
        """Add a new memory"""
        memory = Memory(content, memory_type, importance, metadata)
        self.memories.append(memory)
        logger.info(f"Added memory: {content[:50]}...")
        self.save_memories()
        return memory
    
    def get_memories(self, memory_type: Optional[MemoryType] = None,
                    query: Optional[str] = None,
                    min_importance: float = 0.0) -> List[Memory]:
        """Get memories by type and/or query"""
        results = self.memories
        
        # Filter by type
        if memory_type:
            results = [m for m in results if m.memory_type == memory_type]
        
        # Filter by importance
        results = [m for m in results if m.importance >= min_importance]
        
        # Filter by query
        if query:
            query = query.lower()
            results = [m for m in results if query in m.content.lower()]
        
        # Update access stats
        for memory in results:
            memory.access()
        
        # Save updates
        self.save_memories()
        
        # Sort by importance and recency
        results.sort(key=lambda m: (m.importance, m.accessed_at), reverse=True)
        
        return results
    
    def forget_memory(self, memory_id: str) -> bool:
        """Delete a memory by ID"""
        for i, memory in enumerate(self.memories):
            if memory.id == memory_id:
                del self.memories[i]
                logger.info(f"Deleted memory: {memory_id}")
                self.save_memories()
                return True
        
        logger.warning(f"Memory not found: {memory_id}")
        return False
    
    def get_spatial_context(self) -> Dict[str, Any]:
        """Get spatial context (document position information)"""
        memories = self.get_memories(memory_type=MemoryType.SPATIAL)
        
        if not memories:
            return {"position": 0, "section": "unknown", "history": []}
        
        # Extract most recent position information
        latest = memories[0]
        history = []
        
        for memory in memories[:10]:  # Get up to 10 recent positions
            if "position" in memory.metadata:
                history.append({
                    "position": memory.metadata["position"],
                    "section": memory.metadata.get("section", "unknown"),
                    "timestamp": memory.created_at.isoformat()
                })
        
        # Return current position and history
        return {
            "position": latest.metadata.get("position", 0),
            "section": latest.metadata.get("section", "unknown"),
            "history": history
        }
    
    def update_spatial_context(self, position: int, section: str = "unknown") -> None:
        """Update spatial context with current position"""
        self.add_memory(
            content=f"Currently at position {position} in section {section}",
            memory_type=MemoryType.SPATIAL,
            importance=0.7,
            metadata={
                "position": position,
                "section": section
            }
        )
    
    def summarize_document_context(self) -> Dict[str, Any]:
        """Summarize document context"""
        # Get document memories
        doc_memories = self.get_memories(memory_type=MemoryType.DOCUMENT)
        content_memories = self.get_memories(memory_type=MemoryType.CONTEXTUAL)
        
        sections = {}
        topics = set()
        
        # Extract section information
        for memory in doc_memories:
            if "section" in memory.metadata:
                section = memory.metadata["section"]
                if section not in sections:
                    sections[section] = {
                        "content": memory.content,
                        "importance": memory.importance,
                        "updated": memory.created_at.isoformat()
                    }
        
        # Extract topics from content
        for memory in content_memories:
            if "topics" in memory.metadata:
                for topic in memory.metadata["topics"]:
                    topics.add(topic)
        
        return {
            "sections": sections,
            "topics": list(topics),
            "section_count": len(sections),
            "spatial_context": self.get_spatial_context()
        }

def create_memory_system(memory_file: str = "data/memory/agent_memory.json") -> MemorySystem:
    """Create and return a memory system instance"""
    return MemorySystem(memory_file)
