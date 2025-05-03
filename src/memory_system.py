"""
Memory System for AI Agent in Microsoft Word

This module provides a memory system for the AI Agent to understand and remember
its surroundings and context within Microsoft Word documents.
"""

import datetime
import json
import os
from enum import Enum
from typing import Dict, List, Any, Optional, Union


class MemoryType(Enum):
    """Types of memories the agent can have"""
    SPATIAL = "spatial"  # Location within document (section, paragraph, etc.)
    TEMPORAL = "temporal"  # Time-based memories (when something happened)
    CONTEXTUAL = "contextual"  # Content understanding and relationships
    PROCEDURAL = "procedural"  # How to perform specific tasks
    DOCUMENT = "document"  # Document properties and structure
    LEARNING = "learning"  # Memories related to learning and self-improvement


class Memory:
    """Individual memory object"""

    def __init__(self,
                 content: str,
                 memory_type: MemoryType,
                 metadata: Optional[Dict[str, Any]] = None,
                 importance: float = 0.5,
                 timestamp: Optional[datetime.datetime] = None):
        """
        Initialize a new memory

        Args:
            content: The actual memory content
            memory_type: Type of memory (spatial, temporal, etc.)
            metadata: Additional information about the memory
            importance: How important this memory is (0.0 to 1.0)
            timestamp: When this memory was created
        """
        self.content = content
        self.memory_type = memory_type
        self.metadata = metadata or {}
        self.importance = importance
        self.timestamp = timestamp or datetime.datetime.now()
        self.last_accessed = self.timestamp
        self.access_count = 0

    def to_dict(self) -> Dict[str, Any]:
        """Convert memory to dictionary for serialization"""
        return {
            "content": self.content,
            "memory_type": self.memory_type.value,
            "metadata": self.metadata,
            "importance": self.importance,
            "timestamp": self.timestamp.isoformat(),
            "last_accessed": self.last_accessed.isoformat(),
            "access_count": self.access_count
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Memory':
        """Create memory from dictionary"""
        return cls(
            content=data["content"],
            memory_type=MemoryType(data["memory_type"]),
            metadata=data["metadata"],
            importance=data["importance"],
            timestamp=datetime.datetime.fromisoformat(data["timestamp"])
        )

    def access(self) -> None:
        """Update memory access statistics"""
        self.last_accessed = datetime.datetime.now()
        self.access_count += 1


class MemorySystem:
    """Memory system for the AI Agent"""

    def __init__(self,
                 memory_file: Optional[str] = None,
                 max_memories: int = 1000,
                 verbose: bool = False):
        """
        Initialize the memory system

        Args:
            memory_file: Path to file for persisting memories
            max_memories: Maximum number of memories to store
            verbose: Whether to print debug information
        """
        self.memories: List[Memory] = []
        self.memory_file = memory_file
        self.max_memories = max_memories
        self.verbose = verbose

        # Load existing memories if file exists
        if memory_file and os.path.exists(memory_file):
            self.load_memories()

    def add_memory(self,
                  content: str,
                  memory_type: MemoryType,
                  metadata: Optional[Dict[str, Any]] = None,
                  importance: float = 0.5) -> Memory:
        """
        Add a new memory to the system

        Args:
            content: The memory content
            memory_type: Type of memory
            metadata: Additional information
            importance: Memory importance (0.0 to 1.0)

        Returns:
            The created Memory object
        """
        memory = Memory(
            content=content,
            memory_type=memory_type,
            metadata=metadata,
            importance=importance
        )

        self.memories.append(memory)

        # If we exceed max memories, remove least important ones
        if len(self.memories) > self.max_memories:
            self._prune_memories()

        # Save memories if we have a file
        if self.memory_file:
            self.save_memories()

        if self.verbose:
            print(f"[MEMORY] Added: {content[:50]}...")

        return memory

    def get_memories(self,
                    memory_type: Optional[MemoryType] = None,
                    query: Optional[str] = None,
                    limit: int = 10,
                    recency_weight: float = 0.5) -> List[Memory]:
        """
        Retrieve memories based on type and/or query

        Args:
            memory_type: Filter by memory type
            query: Search term for memory content
            limit: Maximum number of memories to return
            recency_weight: How much to weight recent memories (0.0 to 1.0)

        Returns:
            List of matching Memory objects
        """
        # Filter memories
        filtered_memories = self.memories

        if memory_type:
            filtered_memories = [m for m in filtered_memories
                               if m.memory_type == memory_type]

        if query:
            query = query.lower()
            filtered_memories = [m for m in filtered_memories
                               if query in m.content.lower()]

        # Sort by importance and recency
        def score_memory(memory: Memory) -> float:
            # Calculate recency score (newer = higher score)
            age = (datetime.datetime.now() - memory.timestamp).total_seconds()
            max_age = 60 * 60 * 24 * 30  # 30 days in seconds
            recency_score = 1.0 - min(age / max_age, 1.0)

            # Combine importance and recency
            return (memory.importance * (1 - recency_weight) +
                    recency_score * recency_weight)

        sorted_memories = sorted(filtered_memories,
                                key=score_memory,
                                reverse=True)

        # Update access statistics
        for memory in sorted_memories[:limit]:
            memory.access()

        # Save if we updated access counts
        if sorted_memories and self.memory_file:
            self.save_memories()

        return sorted_memories[:limit]

    def update_memory(self,
                     memory: Memory,
                     content: Optional[str] = None,
                     metadata: Optional[Dict[str, Any]] = None,
                     importance: Optional[float] = None) -> Memory:
        """
        Update an existing memory

        Args:
            memory: The memory to update
            content: New content (if None, keep existing)
            metadata: New metadata (if None, keep existing)
            importance: New importance (if None, keep existing)

        Returns:
            The updated Memory object
        """
        if content is not None:
            memory.content = content

        if metadata is not None:
            memory.metadata = metadata

        if importance is not None:
            memory.importance = importance

        # Update timestamp to reflect the change
        memory.timestamp = datetime.datetime.now()

        # Save memories if we have a file
        if self.memory_file:
            self.save_memories()

        if self.verbose:
            print(f"[MEMORY] Updated: {memory.content[:50]}...")

        return memory

    def clear_memories(self, memory_type: Optional[MemoryType] = None) -> int:
        """
        Clear memories of a specific type or all memories

        Args:
            memory_type: Type of memories to clear (None for all)

        Returns:
            Number of memories cleared
        """
        if memory_type:
            original_count = len(self.memories)
            self.memories = [m for m in self.memories
                           if m.memory_type != memory_type]
            cleared_count = original_count - len(self.memories)
        else:
            cleared_count = len(self.memories)
            self.memories = []

        # Save memories if we have a file
        if self.memory_file:
            self.save_memories()

        if self.verbose:
            print(f"[MEMORY] Cleared {cleared_count} memories")

        return cleared_count

    def save_memories(self) -> bool:
        """
        Save memories to file

        Returns:
            True if successful, False otherwise
        """
        if not self.memory_file:
            return False

        try:
            with open(self.memory_file, 'w') as f:
                json_data = {
                    "memories": [m.to_dict() for m in self.memories]
                }
                json.dump(json_data, f, indent=2)

            if self.verbose:
                print(f"[MEMORY] Saved {len(self.memories)} memories to {self.memory_file}")

            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to save memories: {str(e)}")
            return False

    def load_memories(self) -> bool:
        """
        Load memories from file

        Returns:
            True if successful, False otherwise
        """
        if not self.memory_file or not os.path.exists(self.memory_file):
            return False

        try:
            with open(self.memory_file, 'r') as f:
                json_data = json.load(f)

                self.memories = [
                    Memory.from_dict(m_data)
                    for m_data in json_data.get("memories", [])
                ]

            if self.verbose:
                print(f"[MEMORY] Loaded {len(self.memories)} memories from {self.memory_file}")

            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to load memories: {str(e)}")
            return False

    def _prune_memories(self) -> int:
        """
        Remove least important memories to stay under max_memories

        Returns:
            Number of memories pruned
        """
        if len(self.memories) <= self.max_memories:
            return 0

        # Sort by importance (lowest first)
        self.memories.sort(key=lambda m: m.importance)

        # Calculate how many to remove
        to_remove = len(self.memories) - self.max_memories

        # Remove least important memories
        self.memories = self.memories[to_remove:]

        if self.verbose:
            print(f"[MEMORY] Pruned {to_remove} least important memories")

        return to_remove

    def summarize_memories(self, memory_type: Optional[MemoryType] = None) -> Dict[str, Any]:
        """
        Get a summary of memories

        Args:
            memory_type: Type of memories to summarize (None for all)

        Returns:
            Dictionary with memory statistics
        """
        filtered_memories = self.memories
        if memory_type:
            filtered_memories = [m for m in filtered_memories
                               if m.memory_type == memory_type]

        if not filtered_memories:
            return {
                "count": 0,
                "types": {},
                "oldest": None,
                "newest": None,
                "avg_importance": 0.0
            }

        # Count by type
        type_counts = {}
        for m in filtered_memories:
            type_name = m.memory_type.value
            type_counts[type_name] = type_counts.get(type_name, 0) + 1

        # Get oldest and newest
        oldest = min(filtered_memories, key=lambda m: m.timestamp)
        newest = max(filtered_memories, key=lambda m: m.timestamp)

        # Calculate average importance
        avg_importance = sum(m.importance for m in filtered_memories) / len(filtered_memories)

        return {
            "count": len(filtered_memories),
            "types": type_counts,
            "oldest": oldest.timestamp.isoformat(),
            "newest": newest.timestamp.isoformat(),
            "avg_importance": avg_importance
        }


class DocumentMemory:
    """Specialized memory system for document structure and content"""

    def __init__(self, memory_system: MemorySystem):
        """
        Initialize document memory

        Args:
            memory_system: The main memory system to use
        """
        self.memory_system = memory_system
        self.current_position = {
            "section": None,
            "section_index": 0,
            "paragraph": None,
            "line": None,
            "character": None,
            "last_updated": datetime.datetime.now()
        }
        self.document_structure = {
            "title": None,
            "sections": [],
            "current_section_index": 0
        }
        self.position_history = []  # Track position changes over time

    def update_position(self,
                       section: Optional[str] = None,
                       section_index: Optional[int] = None,
                       paragraph: Optional[int] = None,
                       line: Optional[int] = None,
                       character: Optional[int] = None) -> None:
        """
        Update the agent's current position in the document

        Args:
            section: Current section name
            section_index: Current section index
            paragraph: Current paragraph number
            line: Current line number
            character: Current character position
        """
        # Save previous position to history before updating
        if any(x is not None for x in [section, section_index, paragraph, line, character]):
            # Only add to history if something is actually changing
            self.position_history.append(self.current_position.copy())

            # Limit history size to prevent memory bloat
            if len(self.position_history) > 20:  # Keep last 20 positions
                self.position_history = self.position_history[-20:]

        # Update only provided values
        if section is not None:
            self.current_position["section"] = section

        if section_index is not None:
            self.current_position["section_index"] = section_index

        if paragraph is not None:
            self.current_position["paragraph"] = paragraph

        if line is not None:
            self.current_position["line"] = line

        if character is not None:
            self.current_position["character"] = character

        # Update timestamp
        self.current_position["last_updated"] = datetime.datetime.now()

        # Create a memory of the position change
        position_str = f"Position: {self.current_position['section']} - "
        if self.current_position['section_index'] is not None:
            position_str += f"Section #{self.current_position['section_index']} - "
        position_str += f"P{self.current_position['paragraph']} "
        position_str += f"L{self.current_position['line']} "
        position_str += f"C{self.current_position['character']}"

        # Calculate importance based on how significant the position change is
        importance = 0.3  # Default importance

        # If we moved to a new section, that's more important
        if len(self.position_history) > 0 and self.position_history[-1]["section"] != self.current_position["section"]:
            importance = 0.7

        self.memory_system.add_memory(
            content=position_str,
            memory_type=MemoryType.SPATIAL,
            metadata=self.current_position.copy(),
            importance=importance
        )

    def remember_document_structure(self,
                                   title: str,
                                   sections: List[str]) -> None:
        """
        Remember the overall document structure

        Args:
            title: Document title
            sections: List of section names/headings
        """
        self.document_structure["title"] = title
        self.document_structure["sections"] = sections

        # Create a memory of the document structure
        structure_str = f"Document '{title}' has {len(sections)} sections: "
        structure_str += ", ".join(sections)

        self.memory_system.add_memory(
            content=structure_str,
            memory_type=MemoryType.DOCUMENT,
            metadata=self.document_structure.copy(),
            importance=0.9  # High importance as this is key context
        )

    def remember_section_content(self,
                                section: str,
                                content_summary: str,
                                full_content: Optional[str] = None) -> None:
        """
        Remember content of a specific section

        Args:
            section: Section name
            content_summary: Brief summary of section content
            full_content: Full section content (optional)
        """
        metadata = {
            "section": section,
            "full_content": full_content
        }

        # Create a memory of the section content
        content_str = f"Section '{section}' contains: {content_summary}"

        self.memory_system.add_memory(
            content=content_str,
            memory_type=MemoryType.CONTEXTUAL,
            metadata=metadata,
            importance=0.8  # High importance as this is key context
        )

    def remember_document_change(self,
                                change_type: str,
                                description: str,
                                location: Optional[Dict[str, Any]] = None) -> None:
        """
        Remember a change made to the document

        Args:
            change_type: Type of change (add, delete, modify)
            description: Description of the change
            location: Where the change occurred
        """
        metadata = {
            "change_type": change_type,
            "location": location or self.current_position.copy()
        }

        # Create a memory of the document change
        change_str = f"{change_type.capitalize()} at "

        if location and "section" in location:
            change_str += f"{location['section']}: "
        elif self.current_position["section"]:
            change_str += f"{self.current_position['section']}: "

        change_str += description

        self.memory_system.add_memory(
            content=change_str,
            memory_type=MemoryType.TEMPORAL,
            metadata=metadata,
            importance=0.7  # Moderately high importance
        )

    def get_position_history(self, limit: int = 5) -> List[Dict[str, Any]]:
        """
        Get the history of document positions

        Args:
            limit: Maximum number of historical positions to return

        Returns:
            List of historical positions, most recent first
        """
        # Return the most recent positions first
        return self.position_history[-limit:][::-1] if self.position_history else []

    def get_current_context(self) -> Dict[str, Any]:
        """
        Get the current document context

        Returns:
            Dictionary with current document context
        """
        # Get current section
        current_section = self.current_position["section"]

        # Get memories related to current section
        section_memories = self.memory_system.get_memories(
            memory_type=MemoryType.CONTEXTUAL,
            query=current_section,
            limit=5
        )

        # Get recent changes
        recent_changes = self.memory_system.get_memories(
            memory_type=MemoryType.TEMPORAL,
            limit=5
        )

        # Get recent position memories
        position_memories = self.memory_system.get_memories(
            memory_type=MemoryType.SPATIAL,
            limit=3
        )

        # Get position history
        position_history = self.get_position_history(limit=3)

        # Calculate time since last position update
        time_since_update = (datetime.datetime.now() - self.current_position["last_updated"]).total_seconds()
        position_freshness = "current" if time_since_update < 60 else "stale"

        # Build context dictionary
        context = {
            "current_position": self.current_position.copy(),
            "position_freshness": position_freshness,
            "position_history": position_history,
            "document_structure": self.document_structure.copy(),
            "section_content": [m.content for m in section_memories],
            "recent_changes": [m.content for m in recent_changes],
            "recent_positions": [m.content for m in position_memories]
        }

        return context
