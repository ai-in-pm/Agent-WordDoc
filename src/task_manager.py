"""
Task Manager for Iterative Agent in Microsoft Word

This module provides a task management system for breaking down complex tasks
into smaller, manageable steps and tracking progress through incremental execution.
"""

import datetime
import json
import os
from enum import Enum
from typing import Dict, List, Any, Optional, Union, Callable

from src.memory_system import MemorySystem, MemoryType
from src.iterative_agent import Task, TaskStatus


class TaskType(Enum):
    """Types of tasks that can be performed"""
    DOCUMENT_CREATION = "document_creation"  # Creating a new document
    CONTENT_GENERATION = "content_generation"  # Generating content for a document
    FORMATTING = "formatting"  # Formatting document content
    RESEARCH = "research"  # Researching information for a document
    EDITING = "editing"  # Editing existing content
    REVIEW = "review"  # Reviewing document content
    ANALYSIS = "analysis"  # Analyzing document content


class TaskPriority(Enum):
    """Priority levels for tasks"""
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"


class TaskTemplate:
    """Template for creating common types of tasks"""
    
    @staticmethod
    def create_document_task(title: str, sections: List[str]) -> Dict[str, Any]:
        """
        Create a task for document creation
        
        Args:
            title: Title of the document
            sections: List of section names
            
        Returns:
            Task definition dictionary
        """
        return {
            "name": f"Create document: {title}",
            "description": f"Create a new document titled '{title}' with the following sections: {', '.join(sections)}",
            "estimated_steps": len(sections) + 2,  # Title + sections + conclusion
            "metadata": {
                "type": TaskType.DOCUMENT_CREATION.value,
                "title": title,
                "sections": sections
            }
        }
    
    @staticmethod
    def create_section_task(section_name: str, content_description: str) -> Dict[str, Any]:
        """
        Create a task for section content generation
        
        Args:
            section_name: Name of the section
            content_description: Description of the content to generate
            
        Returns:
            Task definition dictionary
        """
        return {
            "name": f"Write section: {section_name}",
            "description": f"Generate content for the '{section_name}' section. {content_description}",
            "estimated_steps": 3,  # Planning, writing, reviewing
            "metadata": {
                "type": TaskType.CONTENT_GENERATION.value,
                "section_name": section_name,
                "content_description": content_description
            }
        }
    
    @staticmethod
    def create_research_task(topic: str, focus_areas: List[str]) -> Dict[str, Any]:
        """
        Create a task for research
        
        Args:
            topic: Research topic
            focus_areas: Specific areas to focus on
            
        Returns:
            Task definition dictionary
        """
        return {
            "name": f"Research: {topic}",
            "description": f"Research information about '{topic}' focusing on: {', '.join(focus_areas)}",
            "estimated_steps": len(focus_areas) + 1,  # Each focus area + synthesis
            "metadata": {
                "type": TaskType.RESEARCH.value,
                "topic": topic,
                "focus_areas": focus_areas
            }
        }
    
    @staticmethod
    def create_editing_task(section_name: str, editing_goals: List[str]) -> Dict[str, Any]:
        """
        Create a task for editing content
        
        Args:
            section_name: Name of the section to edit
            editing_goals: Specific editing goals
            
        Returns:
            Task definition dictionary
        """
        return {
            "name": f"Edit section: {section_name}",
            "description": f"Edit the content of the '{section_name}' section with these goals: {', '.join(editing_goals)}",
            "estimated_steps": len(editing_goals),
            "metadata": {
                "type": TaskType.EDITING.value,
                "section_name": section_name,
                "editing_goals": editing_goals
            }
        }


class TaskBreakdownStrategy:
    """Strategies for breaking down tasks into smaller steps"""
    
    @staticmethod
    def break_down_document_creation(task: Task) -> List[Dict[str, Any]]:
        """
        Break down a document creation task
        
        Args:
            task: The task to break down
            
        Returns:
            List of subtask definitions
        """
        metadata = task.metadata
        sections = metadata.get("sections", [])
        
        subtasks = []
        
        # Add task for document setup
        subtasks.append({
            "name": "Document setup",
            "description": f"Set up the document with title '{metadata.get('title', 'Document')}' and prepare the structure",
            "estimated_steps": 2,
            "metadata": {
                "type": "setup",
                "parent_type": TaskType.DOCUMENT_CREATION.value
            }
        })
        
        # Add tasks for each section
        for section in sections:
            subtasks.append(TaskTemplate.create_section_task(
                section_name=section,
                content_description=f"Write content for the {section} section of the document."
            ))
        
        # Add task for review and finalization
        subtasks.append({
            "name": "Review and finalize",
            "description": "Review the entire document and make final adjustments",
            "estimated_steps": 3,
            "metadata": {
                "type": "review",
                "parent_type": TaskType.DOCUMENT_CREATION.value
            }
        })
        
        return subtasks
    
    @staticmethod
    def break_down_content_generation(task: Task) -> List[Dict[str, Any]]:
        """
        Break down a content generation task
        
        Args:
            task: The task to break down
            
        Returns:
            List of subtask definitions
        """
        metadata = task.metadata
        section_name = metadata.get("section_name", "Section")
        
        subtasks = []
        
        # Add task for planning
        subtasks.append({
            "name": f"Plan {section_name} content",
            "description": f"Plan the structure and key points for the {section_name} section",
            "estimated_steps": 1,
            "metadata": {
                "type": "planning",
                "parent_type": TaskType.CONTENT_GENERATION.value,
                "section_name": section_name
            }
        })
        
        # Add task for writing
        subtasks.append({
            "name": f"Write {section_name} content",
            "description": f"Write the content for the {section_name} section based on the plan",
            "estimated_steps": 2,
            "metadata": {
                "type": "writing",
                "parent_type": TaskType.CONTENT_GENERATION.value,
                "section_name": section_name
            }
        })
        
        # Add task for review
        subtasks.append({
            "name": f"Review {section_name} content",
            "description": f"Review and refine the content for the {section_name} section",
            "estimated_steps": 1,
            "metadata": {
                "type": "review",
                "parent_type": TaskType.CONTENT_GENERATION.value,
                "section_name": section_name
            }
        })
        
        return subtasks
    
    @staticmethod
    def break_down_research(task: Task) -> List[Dict[str, Any]]:
        """
        Break down a research task
        
        Args:
            task: The task to break down
            
        Returns:
            List of subtask definitions
        """
        metadata = task.metadata
        topic = metadata.get("topic", "Topic")
        focus_areas = metadata.get("focus_areas", [])
        
        subtasks = []
        
        # Add tasks for each focus area
        for area in focus_areas:
            subtasks.append({
                "name": f"Research {area}",
                "description": f"Research information about {area} related to {topic}",
                "estimated_steps": 2,
                "metadata": {
                    "type": "research_area",
                    "parent_type": TaskType.RESEARCH.value,
                    "topic": topic,
                    "focus_area": area
                }
            })
        
        # Add task for synthesis
        subtasks.append({
            "name": f"Synthesize {topic} research",
            "description": f"Synthesize the research findings about {topic} into a cohesive summary",
            "estimated_steps": 1,
            "metadata": {
                "type": "synthesis",
                "parent_type": TaskType.RESEARCH.value,
                "topic": topic
            }
        })
        
        return subtasks


class AdvancedTaskManager:
    """Advanced manager for tasks with additional features"""
    
    def __init__(self, 
                 memory_system: MemorySystem,
                 tasks_file: Optional[str] = None,
                 verbose: bool = False):
        """
        Initialize the advanced task manager
        
        Args:
            memory_system: The memory system to use
            tasks_file: Path to file for persisting tasks
            verbose: Whether to print debug information
        """
        self.memory_system = memory_system
        self.tasks_file = tasks_file or os.path.join('data', 'agent_tasks.json')
        self.verbose = verbose
        self.tasks: Dict[str, Task] = {}
        
        # Task breakdown strategies
        self.breakdown_strategies = {
            TaskType.DOCUMENT_CREATION.value: TaskBreakdownStrategy.break_down_document_creation,
            TaskType.CONTENT_GENERATION.value: TaskBreakdownStrategy.break_down_content_generation,
            TaskType.RESEARCH.value: TaskBreakdownStrategy.break_down_research
        }
        
        # Load existing tasks if file exists
        if tasks_file and os.path.exists(tasks_file):
            self.load_tasks()
    
    def create_task_from_template(self, 
                                 template_type: TaskType, 
                                 **kwargs) -> Task:
        """
        Create a task using a template
        
        Args:
            template_type: Type of task template to use
            **kwargs: Arguments for the template
            
        Returns:
            The created Task object
        """
        task_def = None
        
        if template_type == TaskType.DOCUMENT_CREATION:
            task_def = TaskTemplate.create_document_task(
                title=kwargs.get("title", "Document"),
                sections=kwargs.get("sections", [])
            )
        elif template_type == TaskType.CONTENT_GENERATION:
            task_def = TaskTemplate.create_section_task(
                section_name=kwargs.get("section_name", "Section"),
                content_description=kwargs.get("content_description", "")
            )
        elif template_type == TaskType.RESEARCH:
            task_def = TaskTemplate.create_research_task(
                topic=kwargs.get("topic", "Topic"),
                focus_areas=kwargs.get("focus_areas", [])
            )
        elif template_type == TaskType.EDITING:
            task_def = TaskTemplate.create_editing_task(
                section_name=kwargs.get("section_name", "Section"),
                editing_goals=kwargs.get("editing_goals", [])
            )
        else:
            raise ValueError(f"Unsupported template type: {template_type}")
        
        # Create the task
        task = Task(
            name=task_def["name"],
            description=task_def["description"],
            estimated_steps=task_def["estimated_steps"],
            metadata=task_def["metadata"]
        )
        
        self.tasks[task.id] = task
        
        # Save tasks
        self.save_tasks()
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Created task from template: {task.name}",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "task_id": task.id,
                "task_name": task.name,
                "task_type": template_type.value
            },
            importance=0.7
        )
        
        if self.verbose:
            print(f"[TASK] Created from template: {task.name} (ID: {task.id})")
        
        return task
    
    def auto_break_down_task(self, task_id: str) -> List[Task]:
        """
        Automatically break down a task based on its type
        
        Args:
            task_id: ID of the task to break down
            
        Returns:
            List of created subtask objects
        """
        task = self.get_task(task_id)
        if not task:
            if self.verbose:
                print(f"[ERROR] Task not found: {task_id}")
            return []
        
        # Get the task type
        task_type = task.metadata.get("type")
        if not task_type or task_type not in self.breakdown_strategies:
            if self.verbose:
                print(f"[ERROR] No breakdown strategy for task type: {task_type}")
            return []
        
        # Get the breakdown strategy
        strategy = self.breakdown_strategies[task_type]
        
        # Break down the task
        subtask_defs = strategy(task)
        
        # Create the subtasks
        created_subtasks = []
        for subtask_def in subtask_defs:
            subtask = Task(
                name=subtask_def["name"],
                description=subtask_def["description"],
                parent_id=task_id,
                estimated_steps=subtask_def.get("estimated_steps", 1),
                metadata=subtask_def.get("metadata", {})
            )
            
            self.tasks[subtask.id] = subtask
            task.add_subtask(subtask.id)
            created_subtasks.append(subtask)
        
        # Save tasks
        self.save_tasks()
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Auto-broke down task: {task.name} into {len(created_subtasks)} subtasks",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "parent_task_id": task_id,
                "parent_task_name": task.name,
                "subtask_count": len(created_subtasks),
                "subtask_ids": [task.id for task in created_subtasks]
            },
            importance=0.7
        )
        
        if self.verbose:
            print(f"[TASK] Auto-broke down: {task.name} into {len(created_subtasks)} subtasks")
        
        return created_subtasks
    
    def get_task(self, task_id: str) -> Optional[Task]:
        """
        Get a task by ID
        
        Args:
            task_id: ID of the task
            
        Returns:
            The Task object, or None if not found
        """
        return self.tasks.get(task_id)
    
    def get_tasks_by_status(self, status: TaskStatus) -> List[Task]:
        """
        Get tasks with a specific status
        
        Args:
            status: Status to filter by
            
        Returns:
            List of Task objects with the specified status
        """
        return [task for task in self.tasks.values() if task.status == status]
    
    def get_tasks_by_type(self, task_type: TaskType) -> List[Task]:
        """
        Get tasks of a specific type
        
        Args:
            task_type: Type to filter by
            
        Returns:
            List of Task objects of the specified type
        """
        return [
            task for task in self.tasks.values() 
            if task.metadata.get("type") == task_type.value
        ]
    
    def save_tasks(self) -> bool:
        """
        Save tasks to file
        
        Returns:
            True if successful, False otherwise
        """
        if not self.tasks_file:
            return False
        
        try:
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(self.tasks_file), exist_ok=True)
            
            with open(self.tasks_file, 'w') as f:
                json_data = {
                    "tasks": {task_id: task.to_dict() for task_id, task in self.tasks.items()}
                }
                json.dump(json_data, f, indent=2)
            
            if self.verbose:
                print(f"[TASK] Saved {len(self.tasks)} tasks to {self.tasks_file}")
            
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to save tasks: {str(e)}")
            return False
    
    def load_tasks(self) -> bool:
        """
        Load tasks from file
        
        Returns:
            True if successful, False otherwise
        """
        if not self.tasks_file or not os.path.exists(self.tasks_file):
            return False
        
        try:
            with open(self.tasks_file, 'r') as f:
                json_data = json.load(f)
                
                self.tasks = {
                    task_id: Task.from_dict(task_data) 
                    for task_id, task_data in json_data.get("tasks", {}).items()
                }
            
            if self.verbose:
                print(f"[TASK] Loaded {len(self.tasks)} tasks from {self.tasks_file}")
            
            return True
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to load tasks: {str(e)}")
            return False
