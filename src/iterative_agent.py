"""
Iterative Agent for AI Agent in Microsoft Word

This module provides an Iterative Agent that guides the WordAIAssistant
by breaking down tasks into smaller, incremental steps and focusing on
making steady progress rather than trying to solve everything at once.
"""

import datetime
import json
import os
import time
from enum import Enum
from typing import Dict, List, Any, Optional, Union, Callable

from src.memory_system import MemorySystem, MemoryType


class TaskStatus(Enum):
    """Status of a task in the iterative process"""
    PENDING = "pending"        # Task is waiting to be started
    IN_PROGRESS = "in_progress"  # Task is currently being worked on
    COMPLETED = "completed"    # Task has been completed successfully
    BLOCKED = "blocked"        # Task is blocked by dependencies or issues
    FAILED = "failed"          # Task has failed and needs intervention


class Task:
    """Represents a single task in the iterative process"""
    
    def __init__(self, 
                 name: str,
                 description: str,
                 parent_id: Optional[str] = None,
                 dependencies: Optional[List[str]] = None,
                 estimated_steps: int = 1,
                 metadata: Optional[Dict[str, Any]] = None):
        """
        Initialize a new task
        
        Args:
            name: Short name of the task
            description: Detailed description of what needs to be done
            parent_id: ID of the parent task (if this is a subtask)
            dependencies: List of task IDs that must be completed before this one
            estimated_steps: Estimated number of steps to complete this task
            metadata: Additional information about the task
        """
        self.id = f"task_{int(time.time())}_{hash(name) % 10000:04d}"
        self.name = name
        self.description = description
        self.parent_id = parent_id
        self.dependencies = dependencies or []
        self.estimated_steps = estimated_steps
        self.metadata = metadata or {}
        self.status = TaskStatus.PENDING
        self.progress = 0.0  # 0.0 to 1.0
        self.steps_completed = 0
        self.created_at = datetime.datetime.now()
        self.started_at = None
        self.completed_at = None
        self.subtasks: List[str] = []  # IDs of subtasks
        self.notes: List[Dict[str, Any]] = []
        self.current_step_description = ""
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert task to dictionary for serialization"""
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "parent_id": self.parent_id,
            "dependencies": self.dependencies,
            "estimated_steps": self.estimated_steps,
            "metadata": self.metadata,
            "status": self.status.value,
            "progress": self.progress,
            "steps_completed": self.steps_completed,
            "created_at": self.created_at.isoformat(),
            "started_at": self.started_at.isoformat() if self.started_at else None,
            "completed_at": self.completed_at.isoformat() if self.completed_at else None,
            "subtasks": self.subtasks,
            "notes": self.notes,
            "current_step_description": self.current_step_description
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Task':
        """Create task from dictionary"""
        task = cls(
            name=data["name"],
            description=data["description"],
            parent_id=data["parent_id"],
            dependencies=data["dependencies"],
            estimated_steps=data["estimated_steps"],
            metadata=data["metadata"]
        )
        
        task.id = data["id"]
        task.status = TaskStatus(data["status"])
        task.progress = data["progress"]
        task.steps_completed = data["steps_completed"]
        task.created_at = datetime.datetime.fromisoformat(data["created_at"])
        
        if data["started_at"]:
            task.started_at = datetime.datetime.fromisoformat(data["started_at"])
        
        if data["completed_at"]:
            task.completed_at = datetime.datetime.fromisoformat(data["completed_at"])
        
        task.subtasks = data["subtasks"]
        task.notes = data["notes"]
        task.current_step_description = data["current_step_description"]
        
        return task
    
    def start(self) -> None:
        """Mark the task as started"""
        if self.status == TaskStatus.PENDING:
            self.status = TaskStatus.IN_PROGRESS
            self.started_at = datetime.datetime.now()
    
    def complete(self) -> None:
        """Mark the task as completed"""
        self.status = TaskStatus.COMPLETED
        self.progress = 1.0
        self.completed_at = datetime.datetime.now()
    
    def fail(self, reason: str) -> None:
        """Mark the task as failed"""
        self.status = TaskStatus.FAILED
        self.add_note("failure", reason)
    
    def block(self, reason: str) -> None:
        """Mark the task as blocked"""
        self.status = TaskStatus.BLOCKED
        self.add_note("blocked", reason)
    
    def add_step(self, description: str) -> None:
        """
        Add a completed step to the task
        
        Args:
            description: Description of the completed step
        """
        self.steps_completed += 1
        self.progress = min(1.0, self.steps_completed / max(1, self.estimated_steps))
        self.add_note("step_completed", description)
        self.current_step_description = description
    
    def add_subtask(self, subtask_id: str) -> None:
        """
        Add a subtask to this task
        
        Args:
            subtask_id: ID of the subtask
        """
        if subtask_id not in self.subtasks:
            self.subtasks.append(subtask_id)
    
    def add_note(self, note_type: str, content: str) -> None:
        """
        Add a note to the task
        
        Args:
            note_type: Type of note (e.g., "observation", "issue", "progress")
            content: Content of the note
        """
        self.notes.append({
            "type": note_type,
            "content": content,
            "timestamp": datetime.datetime.now().isoformat()
        })
    
    def update_progress(self, progress: float) -> None:
        """
        Update the progress of the task
        
        Args:
            progress: New progress value (0.0 to 1.0)
        """
        self.progress = max(0.0, min(1.0, progress))
        
        # If progress is 1.0, mark as completed
        if self.progress >= 1.0 and self.status != TaskStatus.COMPLETED:
            self.complete()


class TaskManager:
    """Manager for tasks in the iterative process"""
    
    def __init__(self, 
                 memory_system: MemorySystem,
                 tasks_file: Optional[str] = None,
                 verbose: bool = False):
        """
        Initialize the task manager
        
        Args:
            memory_system: The memory system to use
            tasks_file: Path to file for persisting tasks
            verbose: Whether to print debug information
        """
        self.memory_system = memory_system
        self.tasks_file = tasks_file or os.path.join('data', 'agent_tasks.json')
        self.verbose = verbose
        self.tasks: Dict[str, Task] = {}
        
        # Load existing tasks if file exists
        if tasks_file and os.path.exists(tasks_file):
            self.load_tasks()
    
    def create_task(self, 
                   name: str,
                   description: str,
                   parent_id: Optional[str] = None,
                   dependencies: Optional[List[str]] = None,
                   estimated_steps: int = 1,
                   metadata: Optional[Dict[str, Any]] = None) -> Task:
        """
        Create a new task
        
        Args:
            name: Short name of the task
            description: Detailed description of what needs to be done
            parent_id: ID of the parent task (if this is a subtask)
            dependencies: List of task IDs that must be completed before this one
            estimated_steps: Estimated number of steps to complete this task
            metadata: Additional information about the task
            
        Returns:
            The created Task object
        """
        task = Task(
            name=name,
            description=description,
            parent_id=parent_id,
            dependencies=dependencies,
            estimated_steps=estimated_steps,
            metadata=metadata
        )
        
        self.tasks[task.id] = task
        
        # If this is a subtask, add it to the parent
        if parent_id and parent_id in self.tasks:
            self.tasks[parent_id].add_subtask(task.id)
        
        # Save tasks
        self.save_tasks()
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Created task: {name}",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "task_id": task.id,
                "task_name": name,
                "task_description": description
            },
            importance=0.7
        )
        
        if self.verbose:
            print(f"[TASK] Created: {name} (ID: {task.id})")
        
        return task
    
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
    
    def get_next_task(self) -> Optional[Task]:
        """
        Get the next task to work on
        
        Returns:
            The next Task object to work on, or None if no tasks are ready
        """
        # First, check for in-progress tasks
        in_progress = self.get_tasks_by_status(TaskStatus.IN_PROGRESS)
        if in_progress:
            # Return the in-progress task with the highest progress
            return max(in_progress, key=lambda t: t.progress)
        
        # Next, check for pending tasks with no dependencies or all dependencies completed
        pending = self.get_tasks_by_status(TaskStatus.PENDING)
        ready_tasks = []
        
        for task in pending:
            dependencies_met = True
            for dep_id in task.dependencies:
                dep_task = self.get_task(dep_id)
                if not dep_task or dep_task.status != TaskStatus.COMPLETED:
                    dependencies_met = False
                    break
            
            if dependencies_met:
                ready_tasks.append(task)
        
        if ready_tasks:
            # Return the oldest ready task
            return min(ready_tasks, key=lambda t: t.created_at)
        
        # No tasks are ready
        return None
    
    def update_task_progress(self, 
                            task_id: str, 
                            step_description: str, 
                            progress_increment: Optional[float] = None) -> bool:
        """
        Update the progress of a task
        
        Args:
            task_id: ID of the task
            step_description: Description of the step that was completed
            progress_increment: Amount to increment progress (0.0 to 1.0), or None to use steps
            
        Returns:
            True if successful, False otherwise
        """
        task = self.get_task(task_id)
        if not task:
            if self.verbose:
                print(f"[ERROR] Task not found: {task_id}")
            return False
        
        # Start the task if it's pending
        if task.status == TaskStatus.PENDING:
            task.start()
        
        # Add the step
        task.add_step(step_description)
        
        # Update progress if specified
        if progress_increment is not None:
            new_progress = task.progress + progress_increment
            task.update_progress(new_progress)
        
        # Save tasks
        self.save_tasks()
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Updated task progress: {task.name} - {step_description}",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "task_id": task.id,
                "task_name": task.name,
                "step_description": step_description,
                "progress": task.progress
            },
            importance=0.5
        )
        
        if self.verbose:
            print(f"[TASK] Progress: {task.name} - {task.progress:.0%} - {step_description}")
        
        return True
    
    def complete_task(self, task_id: str, final_note: Optional[str] = None) -> bool:
        """
        Mark a task as completed
        
        Args:
            task_id: ID of the task
            final_note: Optional final note to add to the task
            
        Returns:
            True if successful, False otherwise
        """
        task = self.get_task(task_id)
        if not task:
            if self.verbose:
                print(f"[ERROR] Task not found: {task_id}")
            return False
        
        # Add final note if provided
        if final_note:
            task.add_note("completion", final_note)
        
        # Mark as completed
        task.complete()
        
        # Save tasks
        self.save_tasks()
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Completed task: {task.name}",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "task_id": task.id,
                "task_name": task.name,
                "completion_time": task.completed_at.isoformat() if task.completed_at else None
            },
            importance=0.8
        )
        
        if self.verbose:
            print(f"[TASK] Completed: {task.name}")
        
        return True
    
    def fail_task(self, task_id: str, reason: str) -> bool:
        """
        Mark a task as failed
        
        Args:
            task_id: ID of the task
            reason: Reason for failure
            
        Returns:
            True if successful, False otherwise
        """
        task = self.get_task(task_id)
        if not task:
            if self.verbose:
                print(f"[ERROR] Task not found: {task_id}")
            return False
        
        # Mark as failed
        task.fail(reason)
        
        # Save tasks
        self.save_tasks()
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Failed task: {task.name} - {reason}",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "task_id": task.id,
                "task_name": task.name,
                "failure_reason": reason
            },
            importance=0.9  # High importance for failures
        )
        
        if self.verbose:
            print(f"[TASK] Failed: {task.name} - {reason}")
        
        return True
    
    def break_down_task(self, 
                       task_id: str, 
                       subtasks: List[Dict[str, Any]]) -> List[Task]:
        """
        Break down a task into subtasks
        
        Args:
            task_id: ID of the task to break down
            subtasks: List of subtask definitions, each with name, description, and estimated_steps
            
        Returns:
            List of created subtask objects
        """
        parent_task = self.get_task(task_id)
        if not parent_task:
            if self.verbose:
                print(f"[ERROR] Parent task not found: {task_id}")
            return []
        
        created_subtasks = []
        
        for subtask_def in subtasks:
            subtask = self.create_task(
                name=subtask_def["name"],
                description=subtask_def["description"],
                parent_id=task_id,
                estimated_steps=subtask_def.get("estimated_steps", 1),
                metadata=subtask_def.get("metadata", {})
            )
            created_subtasks.append(subtask)
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Broke down task: {parent_task.name} into {len(created_subtasks)} subtasks",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "parent_task_id": task_id,
                "parent_task_name": parent_task.name,
                "subtask_count": len(created_subtasks),
                "subtask_ids": [task.id for task in created_subtasks]
            },
            importance=0.7
        )
        
        if self.verbose:
            print(f"[TASK] Broke down: {parent_task.name} into {len(created_subtasks)} subtasks")
        
        return created_subtasks
    
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
    
    def get_task_summary(self) -> Dict[str, Any]:
        """
        Get a summary of all tasks
        
        Returns:
            Dictionary with task statistics
        """
        # Count tasks by status
        status_counts = {}
        for status in TaskStatus:
            status_counts[status.value] = len(self.get_tasks_by_status(status))
        
        # Calculate overall progress
        total_tasks = len(self.tasks)
        completed_tasks = status_counts[TaskStatus.COMPLETED.value]
        overall_progress = completed_tasks / total_tasks if total_tasks > 0 else 0.0
        
        # Get top-level tasks (no parent)
        top_level_tasks = [task for task in self.tasks.values() if not task.parent_id]
        
        return {
            "total_tasks": total_tasks,
            "status_counts": status_counts,
            "overall_progress": overall_progress,
            "top_level_tasks": len(top_level_tasks),
            "completed_steps": sum(task.steps_completed for task in self.tasks.values())
        }


class IterativeAgent:
    """Agent that guides the WordAIAssistant through incremental steps"""
    
    def __init__(self, 
                 memory_system: MemorySystem,
                 tasks_file: Optional[str] = None,
                 max_step_size: float = 0.1,  # Maximum progress per step (0.0 to 1.0)
                 verbose: bool = False):
        """
        Initialize the iterative agent
        
        Args:
            memory_system: The memory system to use
            tasks_file: Path to file for persisting tasks
            max_step_size: Maximum progress to make in a single step (0.0 to 1.0)
            verbose: Whether to print debug information
        """
        self.memory_system = memory_system
        self.task_manager = TaskManager(
            memory_system=memory_system,
            tasks_file=tasks_file,
            verbose=verbose
        )
        self.max_step_size = max_step_size
        self.verbose = verbose
        self.current_task_id = None
        self.current_step_number = 0
    
    def create_main_task(self, name: str, description: str, estimated_steps: int = 5) -> Task:
        """
        Create a main task to work on
        
        Args:
            name: Name of the task
            description: Detailed description of the task
            estimated_steps: Estimated number of steps to complete
            
        Returns:
            The created Task object
        """
        task = self.task_manager.create_task(
            name=name,
            description=description,
            estimated_steps=estimated_steps
        )
        
        self.current_task_id = task.id
        self.current_step_number = 0
        
        # Add to memory system
        self.memory_system.add_memory(
            content=f"Created main task: {name}",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "task_id": task.id,
                "task_name": name,
                "task_description": description
            },
            importance=0.8
        )
        
        if self.verbose:
            print(f"[ITERATIVE] Created main task: {name}")
            print(f"[ITERATIVE] Description: {description}")
            print(f"[ITERATIVE] Estimated steps: {estimated_steps}")
        
        return task
    
    def break_down_current_task(self, subtasks: List[Dict[str, Any]]) -> List[Task]:
        """
        Break down the current task into subtasks
        
        Args:
            subtasks: List of subtask definitions, each with name, description, and estimated_steps
            
        Returns:
            List of created subtask objects
        """
        if not self.current_task_id:
            if self.verbose:
                print("[ERROR] No current task to break down")
            return []
        
        return self.task_manager.break_down_task(self.current_task_id, subtasks)
    
    def get_next_step(self) -> Dict[str, Any]:
        """
        Get the next step to work on
        
        Returns:
            Dictionary with information about the next step
        """
        # Get the next task to work on
        task = None
        
        if self.current_task_id:
            # Try to continue with the current task
            task = self.task_manager.get_task(self.current_task_id)
            
            # If current task is completed or failed, clear it
            if task and (task.status == TaskStatus.COMPLETED or task.status == TaskStatus.FAILED):
                self.current_task_id = None
                self.current_step_number = 0
                task = None
        
        if not task:
            # Get the next task from the manager
            task = self.task_manager.get_next_task()
            
            if task:
                self.current_task_id = task.id
                self.current_step_number = task.steps_completed
        
        if not task:
            # No tasks available
            return {
                "has_next": False,
                "message": "No tasks available. Create a new task to continue."
            }
        
        # Start the task if it's pending
        if task.status == TaskStatus.PENDING:
            task.start()
            self.task_manager.save_tasks()
        
        # Increment step number
        self.current_step_number += 1
        
        # Calculate progress for this step
        progress_so_far = task.progress
        progress_remaining = 1.0 - progress_so_far
        step_progress = min(self.max_step_size, progress_remaining)
        
        # Prepare step information
        step_info = {
            "has_next": True,
            "task_id": task.id,
            "task_name": task.name,
            "task_description": task.description,
            "step_number": self.current_step_number,
            "total_steps": task.estimated_steps,
            "progress_so_far": progress_so_far,
            "step_progress": step_progress,
            "previous_step": task.current_step_description if task.current_step_description else "Starting the task"
        }
        
        if self.verbose:
            print(f"[ITERATIVE] Next step: {task.name} - Step {self.current_step_number}/{task.estimated_steps}")
            print(f"[ITERATIVE] Progress so far: {progress_so_far:.0%}")
            print(f"[ITERATIVE] Previous step: {step_info['previous_step']}")
        
        return step_info
    
    def complete_step(self, 
                     task_id: str, 
                     step_description: str, 
                     progress_increment: Optional[float] = None) -> Dict[str, Any]:
        """
        Complete a step in the task
        
        Args:
            task_id: ID of the task
            step_description: Description of what was accomplished in this step
            progress_increment: Amount to increment progress (0.0 to 1.0), or None to use steps
            
        Returns:
            Dictionary with updated task information
        """
        # Update the task progress
        success = self.task_manager.update_task_progress(
            task_id=task_id,
            step_description=step_description,
            progress_increment=progress_increment
        )
        
        if not success:
            return {
                "success": False,
                "message": f"Failed to update task progress. Task ID: {task_id}"
            }
        
        # Get the updated task
        task = self.task_manager.get_task(task_id)
        
        # Check if the task is now complete
        if task.progress >= 1.0:
            self.task_manager.complete_task(task_id, "Completed all steps")
            
            if self.verbose:
                print(f"[ITERATIVE] Task completed: {task.name}")
            
            # Clear current task if this was it
            if self.current_task_id == task_id:
                self.current_task_id = None
                self.current_step_number = 0
        
        # Return updated task information
        return {
            "success": True,
            "task_id": task.id,
            "task_name": task.name,
            "progress": task.progress,
            "status": task.status.value,
            "steps_completed": task.steps_completed,
            "is_complete": task.status == TaskStatus.COMPLETED
        }
    
    def get_task_summary(self) -> Dict[str, Any]:
        """
        Get a summary of all tasks
        
        Returns:
            Dictionary with task statistics
        """
        summary = self.task_manager.get_task_summary()
        
        if self.verbose:
            print("\n=== Task Summary ===")
            print(f"Total tasks: {summary['total_tasks']}")
            print(f"Overall progress: {summary['overall_progress']:.0%}")
            print(f"Completed steps: {summary['completed_steps']}")
            print(f"Status counts: {summary['status_counts']}")
            print("=====================\n")
        
        return summary
    
    def get_current_context(self) -> Dict[str, Any]:
        """
        Get the current context for the iterative agent
        
        Returns:
            Dictionary with current context information
        """
        context = {
            "has_current_task": self.current_task_id is not None,
            "current_task_id": self.current_task_id,
            "current_step_number": self.current_step_number,
            "task_summary": self.get_task_summary()
        }
        
        if self.current_task_id:
            task = self.task_manager.get_task(self.current_task_id)
            if task:
                context["current_task"] = {
                    "name": task.name,
                    "description": task.description,
                    "progress": task.progress,
                    "status": task.status.value,
                    "steps_completed": task.steps_completed,
                    "estimated_steps": task.estimated_steps,
                    "current_step_description": task.current_step_description
                }
        
        return context
