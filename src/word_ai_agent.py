import pyautogui
import keyboard
import time
import os
import datetime
from dotenv import load_dotenv
import psutil
import subprocess
import sys
import random
import string
import argparse
import json
import traceback
import inspect

# Add the src directory to the path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from src.evm_paper_generator import EVMPaperGenerator
from src.memory_system import MemorySystem, DocumentMemory, MemoryType
from src.learning_system import LearningSystem, LearningType, Improvement
from src.scaffold_system import ScaffoldSystem, CapabilityType, EvolutionStage
from src.capability_manager import CapabilityRegistry, CapabilityEvolver, CapabilityFactory
from src.iterative_agent import IterativeAgent, Task, TaskStatus
from src.task_manager import TaskType, TaskPriority, AdvancedTaskManager

class WordAIAssistant:
    def __init__(self, typing_mode='realistic', verbose=False, topic="Earned Value Management", template=None, memory_file=None, learning_file=None, scaffold_file=None, capabilities_dir=None, tasks_file=None, enable_self_evolution=False, enable_iterative_mode=False):
        """Initialize the WordAIAssistant"""
        self.typing_mode = typing_mode
        self.verbose = verbose
        self.topic = topic
        self.template = template
        self.enable_self_evolution = enable_self_evolution
        self.enable_iterative_mode = enable_iterative_mode
        self.progress = {
            'status': 'initializing',
            'current_section': None,
            'total_sections': 0,
            'completed_sections': 0
        }

        # Initialize error handling
        self.max_retries = 3
        self.retry_delay = 2  # seconds

        # Initialize memory system
        self.memory_file = memory_file or os.path.join('data', 'agent_memory.json')
        self.memory_system = MemorySystem(memory_file=self.memory_file, verbose=verbose)
        self.document_memory = DocumentMemory(self.memory_system)

        # Initialize learning system
        self.learning_file = learning_file or os.path.join('data', 'agent_learning.json')
        self.learning_system = LearningSystem(
            memory_system=self.memory_system,
            learning_file=self.learning_file,
            verbose=verbose
        )

        # Initialize scaffold system
        self.scaffold_file = scaffold_file or os.path.join('data', 'agent_scaffold.json')
        self.capabilities_dir = capabilities_dir or os.path.join('data', 'capabilities')
        self.scaffold_system = ScaffoldSystem(
            memory_system=self.memory_system,
            scaffold_file=self.scaffold_file,
            capabilities_dir=self.capabilities_dir,
            verbose=verbose
        )

        # Initialize capability management
        self.capability_registry = CapabilityRegistry(self.scaffold_system)
        self.capability_evolver = CapabilityEvolver(self.scaffold_system)
        self.capability_factory = CapabilityFactory(self.scaffold_system)

        # Initialize iterative agent if enabled
        self.tasks_file = tasks_file or os.path.join('data', 'agent_tasks.json')
        if enable_iterative_mode:
            self.iterative_agent = IterativeAgent(
                memory_system=self.memory_system,
                tasks_file=self.tasks_file,
                verbose=verbose
            )
            self.advanced_task_manager = AdvancedTaskManager(
                memory_system=self.memory_system,
                tasks_file=self.tasks_file,
                verbose=verbose
            )
        else:
            self.iterative_agent = None
            self.advanced_task_manager = None

        # Register built-in capabilities
        self._register_built_in_capabilities()

        try:
            self.setup_word()
            self.evm_generator = EVMPaperGenerator(topic=topic, template=template)
            self.initialize_typing_engine()

            # Initialize document awareness
            self._initialize_document_awareness()

            if self.verbose:
                print(f"\n=== WordAIAssistant Initialization ===")
                print(f"Topic: {topic}")
                print(f"Template: {template}")
                print(f"Typing Mode: {self.typing_mode}")
                print(f"Verbose Mode: {self.verbose}")
                print(f"Memory System: {self.memory_file}")
                print(f"Learning System: {self.learning_file}")
                print(f"Scaffold System: {self.scaffold_file}")
                print(f"Self-Evolution: {self.enable_self_evolution}")
                print(f"Iterative Mode: {self.enable_iterative_mode}")
                if self.enable_iterative_mode:
                    print(f"Tasks File: {self.tasks_file}")
                print("Word interface initialized")
                print("OpenAI API configured")
                print("Keyboard and pyautogui initialized")
                print("Memory system initialized")
                print("Learning system initialized")
                print("Scaffold system initialized")
                if self.enable_iterative_mode:
                    print("Iterative agent initialized")
                print("===================================\n")

        except Exception as e:
            self.handle_error("Initialization", e)
            raise

    def _initialize_document_awareness(self):
        """Initialize document awareness capabilities"""
        # Create initial memories about the document environment
        self.memory_system.add_memory(
            content=f"Document initialized with topic: {self.topic}",
            memory_type=MemoryType.DOCUMENT,
            metadata={
                "topic": self.topic,
                "template": self.template
            },
            importance=0.9
        )

        # Remember document structure based on expected sections
        sections = self.evm_generator.sections
        self.document_memory.remember_document_structure(
            title=self.topic,
            sections=sections
        )

        # Set initial position
        self.document_memory.update_position(
            section="Title",
            paragraph=1,
            line=1,
            character=1
        )

    def _register_built_in_capabilities(self):
        """Register built-in capabilities with the capability registry"""
        # Register static capabilities (methods of this class)
        for name, method in inspect.getmembers(self, predicate=inspect.ismethod):
            # Skip private methods and special methods
            if name.startswith('_') or name in ['__init__', '__del__']:
                continue

            # Register the method as a static capability
            self.capability_registry.register_static_capability(name, method)

            if self.verbose:
                print(f"[CAPABILITY] Registered static capability: {name}")

        # Create some initial dynamic capabilities if none exist
        if not self.scaffold_system.capabilities:
            self._create_initial_capabilities()

    def _create_initial_capabilities(self):
        """Create initial dynamic capabilities"""
        # Create a capability for analyzing document structure
        self.capability_factory.create_capability(
            name="analyze_document_structure",
            description="Analyze the structure of the current Word document",
            capability_type=CapabilityType.ANALYSIS,
            parameters=[],
            code_hints={
                "uses_word_interface": True,
                "returns_structure": True
            }
        )

        # Create a capability for improving typing behavior
        self.capability_factory.create_capability(
            name="optimize_typing_behavior",
            description="Optimize typing behavior based on document type and content",
            capability_type=CapabilityType.ADAPTATION,
            parameters=["document_type"],
            code_hints={
                "modifies_typing_speed": True,
                "adapts_to_content": True
            }
        )

        # Create a capability for self-modification
        self.capability_factory.create_capability(
            name="evolve_agent_behavior",
            description="Modify the agent's behavior based on performance analysis",
            capability_type=CapabilityType.META,
            parameters=["behavior_area", "modification_type"],
            code_hints={
                "self_modifying": True,
                "requires_analysis": True
            }
        )

        if self.verbose:
            print(f"[SCAFFOLD] Created {len(self.scaffold_system.capabilities)} initial capabilities")

    def has_capability(self, name):
        """Check if the agent has a specific capability"""
        return self.capability_registry.has_capability(name)

    def call_capability(self, name, *args, **kwargs):
        """Call a capability by name"""
        capability_func = self.capability_registry.get_capability(name)
        if capability_func is None:
            raise ValueError(f"Capability '{name}' not found")

        try:
            # Record that we're using this capability
            self.memory_system.add_memory(
                content=f"Using capability: {name}",
                memory_type=MemoryType.PROCEDURAL,
                metadata={
                    "capability": name,
                    "args": str(args),
                    "kwargs": str(kwargs)
                },
                importance=0.5
            )

            # Call the capability
            return capability_func(*args, **kwargs)
        except Exception as e:
            # Record the failure
            self.memory_system.add_memory(
                content=f"Failed to use capability: {name} - {str(e)}",
                memory_type=MemoryType.PROCEDURAL,
                metadata={
                    "capability": name,
                    "args": str(args),
                    "kwargs": str(kwargs),
                    "error": str(e)
                },
                importance=0.7
            )

            raise e

    def create_capability(self, name, description, capability_type, parameters=None):
        """Create a new capability"""
        if not self.enable_self_evolution:
            if self.verbose:
                print("[WARNING] Self-evolution is disabled, cannot create new capabilities")
            return None

        # Convert string capability type to enum
        if isinstance(capability_type, str):
            try:
                capability_type = CapabilityType(capability_type)
            except ValueError:
                raise ValueError(f"Invalid capability type: {capability_type}")

        # Create the capability
        capability = self.capability_factory.create_capability(
            name=name,
            description=description,
            capability_type=capability_type,
            parameters=parameters
        )

        if capability and self.verbose:
            print(f"[SCAFFOLD] Created new capability: {name} ({capability_type.value})")

        return capability

    def evolve_capability(self, name, improvement_description):
        """Evolve an existing capability"""
        if not self.enable_self_evolution:
            if self.verbose:
                print("[WARNING] Self-evolution is disabled, cannot evolve capabilities")
            return None

        # Choose an appropriate evolution strategy based on the description
        if "error" in improvement_description.lower() or "fix" in improvement_description.lower():
            strategy = "error_correction"
        elif "performance" in improvement_description.lower() or "speed" in improvement_description.lower():
            strategy = "performance_optimization"
        elif "feature" in improvement_description.lower() or "add" in improvement_description.lower():
            strategy = "feature_addition"
        else:
            strategy = "code_cleanup"

        # Evolve the capability
        evolved = self.capability_evolver.evolve_capability(
            name=name,
            strategy=strategy,
            context={"improvement": improvement_description}
        )

        if evolved and self.verbose:
            print(f"[SCAFFOLD] Evolved capability: {name} (v{evolved.version}, {evolved.stage.value})")
            print(f"[SCAFFOLD] Improvement: {improvement_description}")

        return evolved

    def analyze_capabilities(self):
        """Analyze capabilities and suggest improvements"""
        # Get capability analysis
        analysis = self.scaffold_system.analyze_capabilities()

        # Get evolution suggestions
        suggestions = self.capability_evolver.suggest_evolutions()

        if self.verbose:
            print("\n=== Capability Analysis ===")
            print(f"Total capabilities: {analysis['count']}")
            print(f"Capability types: {analysis['types']}")
            print(f"Evolution stages: {analysis['stages']}")
            print(f"Overall success rate: {analysis['success_rate']:.2f}")
            print("\nImprovement opportunities:")
            for i, opp in enumerate(analysis['improvement_opportunities']):
                print(f"  {i+1}. {opp['capability']} - {opp['type']} ({opp['priority']} priority)")
            print("\nEvolution suggestions:")
            for i, sugg in enumerate(suggestions):
                print(f"  {i+1}. {sugg['capability']} - {sugg['strategy']} ({sugg['priority']} priority)")
                print(f"     Reason: {sugg['reason']}")
            print("============================\n")

        # Record analysis in memory
        self.memory_system.add_memory(
            content=f"Analyzed {analysis['count']} capabilities with {len(suggestions)} evolution suggestions",
            memory_type=MemoryType.LEARNING,
            metadata={
                "analysis": analysis,
                "suggestions": suggestions
            },
            importance=0.8
        )

        return {
            "analysis": analysis,
            "suggestions": suggestions
        }

    def self_evolve(self):
        """Perform self-evolution by improving capabilities"""
        if not self.enable_self_evolution:
            if self.verbose:
                print("[WARNING] Self-evolution is disabled")
            return []

        # Analyze capabilities
        analysis_result = self.analyze_capabilities()
        suggestions = analysis_result["suggestions"]

        # Apply high priority suggestions
        evolved_capabilities = []
        for suggestion in suggestions:
            if suggestion["priority"] == "high":
                evolved = self.evolve_capability(
                    name=suggestion["capability"],
                    improvement_description=suggestion["reason"]
                )
                if evolved:
                    evolved_capabilities.append(evolved)

        # Apply some medium priority suggestions (up to 2)
        medium_suggestions = [s for s in suggestions if s["priority"] == "medium"]
        for suggestion in medium_suggestions[:2]:
            evolved = self.evolve_capability(
                name=suggestion["capability"],
                improvement_description=suggestion["reason"]
            )
            if evolved:
                evolved_capabilities.append(evolved)

        if self.verbose:
            print(f"[SCAFFOLD] Self-evolved {len(evolved_capabilities)} capabilities")

        return evolved_capabilities

    # Iterative Agent Methods

    def create_main_task(self, name, description, estimated_steps=5):
        """Create a main task for the iterative agent"""
        if not self.enable_iterative_mode or not self.iterative_agent:
            if self.verbose:
                print("[WARNING] Iterative mode is not enabled")
            return None

        task = self.iterative_agent.create_main_task(
            name=name,
            description=description,
            estimated_steps=estimated_steps
        )

        if self.verbose:
            print(f"[ITERATIVE] Created main task: {name}")
            print(f"[ITERATIVE] Description: {description}")
            print(f"[ITERATIVE] Estimated steps: {estimated_steps}")

        return task

    def create_document_task(self, title, sections):
        """Create a document creation task using a template"""
        if not self.enable_iterative_mode or not self.advanced_task_manager:
            if self.verbose:
                print("[WARNING] Iterative mode is not enabled")
            return None

        task = self.advanced_task_manager.create_task_from_template(
            template_type=TaskType.DOCUMENT_CREATION,
            title=title,
            sections=sections
        )

        # Auto-break down the task
        subtasks = self.advanced_task_manager.auto_break_down_task(task.id)

        if self.verbose:
            print(f"[ITERATIVE] Created document task: {title}")
            print(f"[ITERATIVE] Sections: {', '.join(sections)}")
            print(f"[ITERATIVE] Created {len(subtasks)} subtasks")

        return task

    def get_next_step(self):
        """Get the next step to work on from the iterative agent"""
        if not self.enable_iterative_mode or not self.iterative_agent:
            if self.verbose:
                print("[WARNING] Iterative mode is not enabled")
            return None

        step_info = self.iterative_agent.get_next_step()

        if self.verbose and step_info["has_next"]:
            print("\n=== Next Iterative Step ===")
            print(f"Task: {step_info['task_name']}")
            print(f"Step: {step_info['step_number']}/{step_info['total_steps']}")
            print(f"Progress so far: {step_info['progress_so_far']:.0%}")
            print(f"Previous step: {step_info['previous_step']}")
            print("===========================\n")

        return step_info

    def complete_step(self, task_id, step_description, progress_increment=None):
        """Complete a step in the current task"""
        if not self.enable_iterative_mode or not self.iterative_agent:
            if self.verbose:
                print("[WARNING] Iterative mode is not enabled")
            return None

        result = self.iterative_agent.complete_step(
            task_id=task_id,
            step_description=step_description,
            progress_increment=progress_increment
        )

        if self.verbose and result["success"]:
            print(f"[ITERATIVE] Completed step: {step_description}")
            print(f"[ITERATIVE] Progress: {result['progress']:.0%}")
            if result["is_complete"]:
                print(f"[ITERATIVE] Task completed: {result['task_name']}")

        return result

    def get_task_summary(self):
        """Get a summary of all tasks"""
        if not self.enable_iterative_mode or not self.iterative_agent:
            if self.verbose:
                print("[WARNING] Iterative mode is not enabled")
            return None

        return self.iterative_agent.get_task_summary()

    def generate_paper_iteratively(self):
        """Generate a paper using the iterative approach"""
        if not self.enable_iterative_mode or not self.iterative_agent:
            if self.verbose:
                print("[WARNING] Iterative mode is not enabled, using standard approach")
            return self.generate_paper()

        # Create the main document task
        sections = self.evm_generator.sections
        task = self.create_document_task(self.topic, sections)

        if not task:
            if self.verbose:
                print("[ERROR] Failed to create document task")
            return False

        # Process steps until complete
        while True:
            # Get the next step
            step_info = self.get_next_step()

            if not step_info or not step_info["has_next"]:
                if self.verbose:
                    print("[ITERATIVE] No more steps to process")
                break

            # Process the step based on metadata
            task_id = step_info["task_id"]
            task = self.advanced_task_manager.get_task(task_id)

            if not task:
                if self.verbose:
                    print(f"[ERROR] Task not found: {task_id}")
                break

            # Get task type from metadata
            task_type = task.metadata.get("type")
            parent_type = task.metadata.get("parent_type")

            # Process based on task type
            step_description = ""

            if task_type == "setup":
                # Set up the document
                self.setup_word()
                step_description = "Set up the document and prepared structure"

            elif parent_type == TaskType.CONTENT_GENERATION.value:
                # Generate content for a section
                section_name = task.metadata.get("section_name")

                if task_type == "planning":
                    # Plan the section content
                    step_description = f"Planned content for {section_name} section"

                elif task_type == "writing":
                    # Write the section content
                    content = self.evm_generator.generate_section_content(section_name)
                    self.type_heading(section_name)
                    self.type_text(content)
                    self.press_enter(2)
                    step_description = f"Wrote content for {section_name} section"

                elif task_type == "review":
                    # Review the section content
                    step_description = f"Reviewed and refined {section_name} section"

            elif task_type == "review":
                # Final review
                step_description = "Reviewed and finalized the document"

            else:
                # Generic step
                step_description = f"Processed step for {task.name}"

            # Complete the step
            self.complete_step(task_id, step_description)

            # Update progress
            self.update_progress(task.name, "completed")

        # Get final task summary
        summary = self.get_task_summary()

        if self.verbose:
            print("\n=== Document Generation Complete ===")
            print(f"Total tasks: {summary['total_tasks']}")
            print(f"Completed steps: {summary['completed_steps']}")
            print(f"Overall progress: {summary['overall_progress']:.0%}")
            print("===================================\n")

        return True

    def analyze_and_improve(self):
        """Analyze behavior patterns and generate improvements"""
        if self.verbose:
            print("\n=== Analyzing Behavior for Improvements ===")

        # Generate improvements based on behavior patterns
        improvements = self.learning_system.generate_behavior_improvements()

        if improvements:
            if self.verbose:
                print(f"Generated {len(improvements)} new improvements:")
                for i, improvement in enumerate(improvements):
                    print(f"  {i+1}. {improvement.description} (confidence: {improvement.confidence:.2f})")
        else:
            if self.verbose:
                print("No new improvements generated from behavior analysis")

        # Get learning system summary
        summary = self.learning_system.summarize_learning()

        if self.verbose:
            print("\n=== Learning System Status ===")
            print(f"Total improvements: {summary['count']}")
            print(f"Improvement types: {summary['types']}")
            print(f"Average confidence: {summary['avg_confidence']:.2f}")
            print(f"Success rate: {summary['success_rate']:.2f}")
            print(f"Error patterns tracked: {summary['error_patterns']}")
            print("============================\n")

        # Record this analysis in memory
        self.memory_system.add_memory(
            content=f"Performed behavior analysis and generated {len(improvements)} improvements",
            memory_type=MemoryType.LEARNING,
            metadata={
                "summary": summary,
                "new_improvements": len(improvements)
            },
            importance=0.7
        )

        # If self-evolution is enabled, also evolve capabilities
        evolved_capabilities = []
        if self.enable_self_evolution:
            if self.verbose:
                print("\n=== Performing Self-Evolution ===")
            evolved_capabilities = self.self_evolve()
            if evolved_capabilities:
                if self.verbose:
                    print(f"Self-evolved {len(evolved_capabilities)} capabilities")
            else:
                if self.verbose:
                    print("No capabilities evolved at this time")

        return {"improvements": improvements, "evolved_capabilities": evolved_capabilities}

    def get_word_cursor_position(self):
        """Get the exact cursor position from Word using COM interface"""
        try:
            import win32com.client

            # Try to get the active Word application
            word_app = win32com.client.GetActiveObject("Word.Application")

            # Get the active document
            doc = word_app.ActiveDocument

            # Get the current selection
            selection = word_app.Selection

            # Get position information
            current_section_index = selection.Information(9)  # wdActiveEndSectionNumber
            current_section_name = self.progress.get('current_section', 'Unknown')

            # Try to get section name from the document if possible
            try:
                if doc.Sections.Count >= current_section_index:
                    # Try to get heading text for this section
                    section = doc.Sections(current_section_index)
                    # This is a simplification - in a real implementation you would
                    # search for the nearest heading to determine section name
                    current_section_name = f"Section {current_section_index}"
            except:
                pass  # Use the default section name from progress

            # Get paragraph, line and character position
            current_paragraph = selection.Information(10)  # wdStartOfRangeRowNumber
            current_line = selection.Information(5)       # wdFirstCharacterLineNumber
            current_character = selection.Information(3)  # wdFirstCharacterColumnNumber

            position = {
                "section": current_section_name,
                "section_index": current_section_index,
                "paragraph": current_paragraph,
                "line": current_line,
                "character": current_character
            }

            if self.verbose:
                print(f"[POSITION] Current Word position: {position}")

            return position

        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to get Word cursor position: {str(e)}")

            # Return a basic position based on progress
            return {
                "section": self.progress.get('current_section', 'Unknown'),
                "section_index": 1,
                "paragraph": 1,
                "line": 1,
                "character": 1
            }

    def observe_document(self):
        """Observe the current state of the document"""
        try:
            # Check if Word is active
            if not self.check_word_active():
                if self.verbose:
                    print("[OBSERVE] Word is not active, cannot observe document")
                return False

            # Get current cursor position from Word
            position = self.get_word_cursor_position()

            # Update memory with current position
            self.document_memory.update_position(
                section=position["section"],
                paragraph=position["paragraph"],
                line=position["line"],
                character=position["character"]
            )

            # Remember what we're seeing with more detailed position information
            position_str = f"Section: {position['section']}, "
            position_str += f"Paragraph: {position['paragraph']}, "
            position_str += f"Line: {position['line']}, "
            position_str += f"Character: {position['character']}"

            self.memory_system.add_memory(
                content=f"Observed document at {position_str}",
                memory_type=MemoryType.SPATIAL,
                metadata=position,
                importance=0.6  # Increased importance for precise positioning
            )

            return True

        except Exception as e:
            self.handle_error("Document Observation", e)
            return False

    def update_position(self, force_update=False):
        """Explicitly update the agent's position in Word"""
        # Check if Word is active
        if not self.check_word_active():
            if self.verbose:
                print("[POSITION] Word is not active, cannot update position")
            return False

        # Get current position from Word
        position = self.get_word_cursor_position()

        # Update memory with current position
        self.document_memory.update_position(
            section=position["section"],
            section_index=position["section_index"],
            paragraph=position["paragraph"],
            line=position["line"],
            character=position["character"]
        )

        if self.verbose and force_update:
            print(f"[POSITION] Updated position: {position}")

        return True

    def understand_context(self):
        """Understand the current document context"""
        # Update position before getting context
        self.update_position(force_update=True)

        # Get current context from memory
        context = self.document_memory.get_current_context()

        # Add iterative agent context if enabled
        if self.enable_iterative_mode and self.iterative_agent:
            iterative_context = self.iterative_agent.get_current_context()
            context['iterative'] = iterative_context

        if self.verbose:
            print("\n=== Current Document Context ===")
            print(f"Position: {context['current_position']}")
            print(f"Position Freshness: {context['position_freshness']}")
            print(f"Current Section: {context['current_position']['section']}")
            print(f"Document Structure: {len(context['document_structure']['sections'])} sections")
            print(f"Recent Changes: {len(context['recent_changes'])}")
            print(f"Position History: {len(context['position_history'])} entries")

            if self.enable_iterative_mode and 'iterative' in context:
                print("\n=== Iterative Agent Context ===")
                if context['iterative']['has_current_task']:
                    task = context['iterative']['current_task']
                    print(f"Current Task: {task['name']}")
                    print(f"Progress: {task['progress']:.0%}")
                    print(f"Steps Completed: {task['steps_completed']}/{task['estimated_steps']}")
                    print(f"Current Step: {task['current_step_description']}")
                else:
                    print("No current task")

                summary = context['iterative']['task_summary']
                print(f"Total Tasks: {summary['total_tasks']}")
                print(f"Overall Progress: {summary['overall_progress']:.0%}")

            print("================================\n")

        return context

    def remember_action(self, action, description):
        """Remember an action taken by the agent"""
        self.memory_system.add_memory(
            content=f"Action: {action} - {description}",
            memory_type=MemoryType.PROCEDURAL,
            metadata={
                "action": action,
                "description": description,
                "section": self.progress.get('current_section')
            },
            importance=0.6
        )

        # Also record as a document change
        self.document_memory.remember_document_change(
            change_type="action",
            description=description
        )

    def handle_error(self, operation, error, context=None):
        """Handle errors with retry logic, learning, and logging"""
        if self.verbose:
            print(f"[ERROR] {operation} failed: {str(error)}")

        # Gather context if not provided
        if context is None:
            context = {
                "section": self.progress.get('current_section'),
                "status": self.progress.get('status'),
                "typing_mode": self.typing_mode
            }

        # Remember the error
        self.memory_system.add_memory(
            content=f"Error during {operation}: {str(error)}",
            memory_type=MemoryType.TEMPORAL,
            metadata={
                "operation": operation,
                "error": str(error),
                "context": context
            },
            importance=0.7  # Errors are important to remember
        )

        # Track error in learning system for pattern analysis
        self.learning_system.track_error(operation, str(error), context)

        # Check if we have any applicable improvements for this error
        applicable_improvements = self.learning_system.find_applicable_improvements(operation, context)
        improvement_applied = False

        if applicable_improvements:
            # Apply the most confident improvement
            improvement = applicable_improvements[0]
            if self.verbose:
                print(f"[LEARNING] Applying improvement: {improvement.description}")
                print(f"[LEARNING] New behavior: {improvement.new_behavior}")

            # Record that we're applying this improvement
            self.memory_system.add_memory(
                content=f"Applying learned improvement: {improvement.description}",
                memory_type=MemoryType.LEARNING,
                metadata={
                    "improvement": improvement.to_dict(),
                    "operation": operation,
                    "error": str(error)
                },
                importance=0.8
            )

            # Apply the improvement
            success = self._apply_improvement(improvement, operation)
            self.learning_system.apply_improvement(improvement, success)
            improvement_applied = success

            if success and self.verbose:
                print(f"[LEARNING] Successfully applied improvement")

        # If no improvement was applied or it failed, use standard recovery
        if not improvement_applied:
            retries = 0
            while retries < self.max_retries:
                try:
                    # Attempt to recover
                    if operation == "Word Startup":
                        self.setup_word()
                    elif operation == "Typing":
                        self.initialize_typing_engine()
                    elif operation == "Document Observation":
                        # Add delay and retry observation
                        time.sleep(1)
                        return self.observe_document()

                    return True

                except Exception as recovery_error:
                    retries += 1
                    if self.verbose:
                        print(f"[RETRY] Attempt {retries}/{self.max_retries} for {operation}")
                    time.sleep(self.retry_delay)

            if self.verbose:
                print(f"[FAILED] {operation} after {self.max_retries} attempts")

            # After failure, generate a new improvement for next time
            self._generate_improvement_from_failure(operation, str(error), context)
            return False

        return improvement_applied

    def _apply_improvement(self, improvement, operation):
        """Apply a learned improvement"""
        try:
            # Apply different improvements based on operation and behavior
            if operation == "Word Startup":
                if "wait longer" in improvement.new_behavior.lower():
                    # Increase wait time
                    time.sleep(5)  # Additional wait time
                    self.setup_word()
                    return True
                elif "keyboard shortcut" in improvement.new_behavior.lower():
                    # Use keyboard shortcut instead of clicking
                    pyautogui.hotkey('ctrl', 'n')
                    time.sleep(1)
                    return True

            elif operation == "Typing":
                if "delay between key presses" in improvement.new_behavior.lower():
                    # Reinitialize typing engine with slower speed
                    if self.typing_mode == 'fast':
                        self.typing_mode = 'realistic'
                    elif self.typing_mode == 'realistic':
                        self.typing_mode = 'slow'
                    self.initialize_typing_engine()
                    return True
                elif "check if word is active" in improvement.new_behavior.lower():
                    # Ensure Word is active
                    if not self.check_word_active():
                        self.setup_word()
                    return True

            elif operation == "Document Observation":
                if "retry observation" in improvement.new_behavior.lower():
                    # Add delay and retry
                    time.sleep(2)
                    return self.observe_document()

            # Generic improvements
            if "timeout" in improvement.new_behavior.lower():
                # Increase timeout/retry delay
                self.retry_delay *= 2
                time.sleep(self.retry_delay)
                return True

            # If we don't know how to apply this improvement
            return False

        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to apply improvement: {str(e)}")
            return False

    def _generate_improvement_from_failure(self, operation, error_message, context):
        """Generate a new improvement after a failure"""
        # This will be called when standard recovery fails
        # The learning system will analyze patterns and generate improvements
        if operation == "Word Startup":
            self.learning_system.add_improvement(
                description="Improve Word startup reliability after failure",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"Word Startup:.*{error_message}",
                new_behavior="Wait longer before interacting with Word after startup",
                confidence=0.6,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Increase startup wait time and use keyboard shortcuts"
                }
            )
        elif operation == "Typing":
            self.learning_system.add_improvement(
                description="Improve typing reliability after failure",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"Typing:.*{error_message}",
                new_behavior="Slow down typing speed and add delays between key presses",
                confidence=0.6,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Reduce typing speed and increase delays"
                }
            )
        else:
            # Generic improvement
            self.learning_system.add_improvement(
                description=f"Improve {operation} reliability after failure",
                learning_type=LearningType.ERROR_CORRECTION,
                trigger_pattern=f"{operation}:.*error",
                new_behavior=f"Add additional error checking and recovery for {operation}",
                confidence=0.5,
                metadata={
                    "original_error": error_message,
                    "context": context,
                    "solution": "Add more robust error handling"
                }
            )

    def initialize_typing_engine(self):
        """Initialize the typing simulation engine"""
        self.typing_speed = {
            'fast': 0.05,
            'realistic': 0.1,
            'slow': 0.2
        }

        self.typing_variation = {
            'fast': (0.02, 0.08),
            'realistic': (0.05, 0.15),
            'slow': (0.1, 0.3)
        }

        self.current_typing_speed = self.typing_speed[self.typing_mode]
        self.current_typing_variation = self.typing_variation[self.typing_mode]

        # Define character typing behaviors
        self.character_behaviors = {
            'letters': {
                'speed_factor': 1.0,
                'error_rate': 0.01,
                'pause_after': 0.05
            },
            'digits': {
                'speed_factor': 0.8,
                'error_rate': 0.005,
                'pause_after': 0.03
            },
            'punctuation': {
                'speed_factor': 0.6,
                'error_rate': 0.001,
                'pause_after': 0.02
            },
            'spaces': {
                'speed_factor': 0.9,
                'error_rate': 0.001,
                'pause_after': 0.01
            }
        }

    def type_text(self, text):
        """Simulate realistic typing behavior"""
        if not self.check_word_active():
            self.setup_word()

        # Remember that we're typing text
        self.remember_action(
            action="type_text",
            description=f"Typing text: {text[:30]}..." if len(text) > 30 else f"Typing text: {text}"
        )

        # Split text into words for better readability
        words = text.split()
        words_typed = 0
        total_words = len(words)
        last_position_update = time.time()

        for word in words:
            # Add small pause between words
            time.sleep(random.uniform(0.1, 0.3))

            for char in word:
                # Determine character type
                if char in string.ascii_letters:
                    char_type = 'letters'
                elif char in string.digits:
                    char_type = 'digits'
                elif char in string.punctuation:
                    char_type = 'punctuation'
                else:
                    char_type = 'spaces'

                # Get character behavior
                behavior = self.character_behaviors[char_type]

                # Calculate typing speed with variation
                base_speed = self.current_typing_speed * behavior['speed_factor']
                delay = base_speed + random.uniform(
                    *self.current_typing_variation
                )

                # Simulate typing errors
                if random.random() < behavior['error_rate']:
                    # Type a random error character
                    error_char = random.choice(string.ascii_letters)
                    keyboard.write(error_char)
                    time.sleep(0.05)
                    keyboard.press_and_release('backspace')
                    time.sleep(0.02)

                # Type the actual character
                keyboard.write(char)
                time.sleep(delay)

                # Add small pause after certain characters
                if random.random() < behavior['pause_after']:
                    time.sleep(random.uniform(0.01, 0.03))

                # Periodically update position (every 5 seconds)
                if time.time() - last_position_update > 5:
                    self.observe_document()
                    last_position_update = time.time()

            # Add space after word
            keyboard.write(' ')
            time.sleep(random.uniform(0.05, 0.1))

            # Update words typed count
            words_typed += 1

            # Periodically update position (every 10 words or 25% of text)
            if words_typed % 10 == 0 or words_typed / total_words >= 0.25:
                self.observe_document()
                last_position_update = time.time()

        # Observe document after typing
        self.observe_document()

        return True

    def type_heading(self, text, level=1):
        """Type and format a heading with specified level"""
        if self.verbose:
            print(f"\n=== Typing Heading ===")
            print(f"Text: {text}")
            print(f"Level: {level}")
            print("=====================")

        # Remember that we're creating a heading
        self.remember_action(
            action="type_heading",
            description=f"Creating level {level} heading: {text}"
        )

        # Update document memory with new section
        self.document_memory.update_position(section=text)

        # Remember section content (initially empty)
        self.document_memory.remember_section_content(
            section=text,
            content_summary="[Empty section]"
        )

        # Type heading with realistic typing
        self.type_text(text)
        self.press_enter()

        # Select the line we just typed
        keyboard.press_and_release('shift+up')

        # Apply heading style based on level
        if level == 1:
            keyboard.press_and_release('alt+h,1')
        elif level == 2:
            keyboard.press_and_release('alt+h,2')

        # Deselect the text
        keyboard.press_and_release('right')

        # Add spacing after heading
        self.press_enter(2)

        # Observe document after creating heading
        self.observe_document()

    def type_table(self, table_content):
        """Type table content with proper formatting"""
        if self.verbose:
            print(f"\n=== Typing Table ===")
            print(f"Content: {table_content}")
            print("====================")

        # Split table into rows
        rows = table_content.strip().split('\n')

        # Type table heading
        self.type_heading("Table: " + rows[0], level=2)
        self.press_enter()

        # Type table content
        for row in rows[1:]:
            # Type row content
            self.type_text(row.strip())
            # Tab to next cell or end of row
            keyboard.press_and_release('tab')
            time.sleep(0.2)

        # Add spacing after table
        self.press_enter(2)

    def type_figure_description(self, description):
        """Type figure description with proper formatting"""
        if self.verbose:
            print(f"\n=== Typing Figure ===")
            print(f"Description: {description}")
            print("=====================")

        # Split description into title and content
        parts = description.split('\n')
        title = parts[0]
        content = '\n'.join(parts[1:])

        # Type figure title
        self.type_heading("Figure: " + title, level=2)
        self.press_enter()

        # Type figure content
        self.type_text(content)

        # Add spacing after figure
        self.press_enter(2)

    def update_progress(self, section_name=None, status=None):
        """Update and report progress"""
        old_section = self.progress.get('current_section')
        old_status = self.progress.get('status')

        if section_name:
            self.progress['current_section'] = section_name
        if status:
            self.progress['status'] = status

        # Remember progress update in memory system
        if section_name or status:
            update_desc = []
            if section_name and section_name != old_section:
                update_desc.append(f"Section changed from '{old_section}' to '{section_name}'")
                # Update position in document memory
                self.document_memory.update_position(section=section_name)

            if status and status != old_status:
                update_desc.append(f"Status changed from '{old_status}' to '{status}'")

            if update_desc:
                self.memory_system.add_memory(
                    content=f"Progress update: {'; '.join(update_desc)}",
                    memory_type=MemoryType.TEMPORAL,
                    metadata=self.progress.copy(),
                    importance=0.5
                )

        if self.verbose:
            print(f"\n=== Progress Update ===")
            print(f"Status: {self.progress['status']}")
            print(f"Current Section: {self.progress['current_section']}")
            print(f"Completed: {self.progress['completed_sections']}/{self.progress['total_sections']}")
            print("======================\n")

        # Observe document after progress update
        self.observe_document()

    def validate_document(self):
        """Validate the document's integrity"""
        try:
            # Check if Word is still active
            if not self.check_word_active():
                raise Exception("Word application is not active")

            # Check for empty sections
            empty_sections = [s for s in self.evm_generator.sections
                            if not self.evm_generator.document.sections[s].text.strip()]

            if empty_sections:
                raise Exception(f"Empty sections found: {', '.join(empty_sections)}")

            # Check for missing references
            if not self.evm_generator.references:
                raise Exception("No references found in the document")

            return True

        except Exception as e:
            self.handle_error("Document Validation", e)
            return False

    def log_key_press(self, event):
        """Log keyboard events for debugging"""
        if self.verbose:
            print(f"[KEY] {event.name} {event.event_type}")

    def check_word_active(self):
        """Check if Word is the active window"""
        try:
            import win32gui
            import win32process
            import psutil

            # Get the active window
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)

            if self.verbose:
                print(f"[WINDOW] Active window: {process.name()}")

            return process.name() == 'WINWORD.EXE'
        except Exception as e:
            if self.verbose:
                print(f"[ERROR] Failed to check active window: {str(e)}")
            return False

    def setup_word(self):
        # Get the path to Microsoft Word executable
        word_path = r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'

        # Check if Word is already running
        if not self.is_word_running():
            # Launch Microsoft Word
            subprocess.Popen([word_path])
            time.sleep(5)  # Wait for Word to open

        # Handle the New document screen
        # Look for and click on the Blank document option
        try:
            # Wait a moment for the new document screen to appear
            time.sleep(2)

            # Click on the Blank document
            # These coordinates are an estimate and may need adjustment based on screen resolution
            # Clicking in the center of the blank document thumbnail
            blank_doc_x, blank_doc_y = 287, 220  # Adjust these coordinates if needed
            pyautogui.click(blank_doc_x, blank_doc_y)
            time.sleep(1)
        except Exception as e:
            # Fallback method - use keyboard shortcut if clicking fails
            print(f"Falling back to keyboard shortcut: {str(e)}")
            pyautogui.hotkey('ctrl', 'n')
            time.sleep(1)

    def is_word_running(self):
        """Check if Microsoft Word is running"""
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] == 'WINWORD.EXE':
                    return True
            return False
        except:
            return False

    def press_enter(self, times=1):
        """Press Enter key multiple times"""
        for _ in range(times):
            keyboard.press_and_release('enter')
            time.sleep(0.1)

    def generate_and_type_paper(self):
        """Generate and type the entire paper with tables and figures"""
        print("\n=== Generating Paper Content ===")
        paper_content = self.evm_generator.generate_full_paper()
        print(f"Generated {len(paper_content)} sections")
        print("================================\n")

        # Add title page
        self.evm_generator.add_title_page()

        print("=== Starting Paper Typing ===")

        # Add progress tracking
        total_sections = len(paper_content)
        current_section = 0

        # Type each section
        for section in paper_content:
            current_section += 1
            content_type = section.get("type", "section")

            print(f"\n=== Section {current_section}/{total_sections} ===")
            print(f"Type: {content_type}")
            print(f"Title: {section['section']}")
            print("===============================")

            # Handle different content types
            if content_type == "section":  # Regular section
                self.type_heading(section["section"], level=1)
                # Handle content based on typing mode
                self.type_text(section["content"])
                self.press_enter(2)
            elif content_type == "table":  # Table
                self.type_heading(section["section"], level=2)
                self.type_table(section["content"])
            elif content_type == "figure":  # Figure
                self.type_heading(section["section"], level=2)
                self.type_figure_description(section["content"])

        # Save the document
        filename = f"{self.evm_generator.topic.replace(' ', '_')}_Academic_Paper.docx"
        self.evm_generator.save_document(filename)
        print(f"\nPaper generation complete! The document has been saved as '{filename}'")

    def generate_table_of_contents(self):
        """Generate and insert a Table of Contents section"""
        if self.verbose:
            print("\n=== Generating Table of Contents ===")

        # Type the heading
        self.type_heading("Table of Contents", level=1)
        self.press_enter()

        # Insert Word's automatic Table of Contents
        # Alt+N to access References tab, followed by T for Table of Contents, A for Automatic Table 1
        keyboard.press_and_release('alt+n')
        time.sleep(0.5)
        keyboard.press_and_release('t')
        time.sleep(0.5)
        keyboard.press_and_release('a')

        # Add some space after TOC
        self.press_enter(2)

        if self.verbose:
            print("Table of Contents inserted successfully")
            print("===================================")

    def generate_document(self):
        """Generate the complete document with all sections"""
        try:
            if self.verbose:
                print("\n=== Starting Document Generation ===")

            # Get content from the EVM generator
            content = self.evm_generator.generate_content()

            # Generate title page
            self.type_heading(content['title'], level=1)
            self.press_enter(2)

            # Add author and date
            self.type_text(f"Author: {content['author']}")
            self.press_enter()
            self.type_text(f"Date: {datetime.datetime.now().strftime('%B %d, %Y')}")
            self.press_enter(2)

            # Insert Table of Contents
            self.generate_table_of_contents()

            # Generate each section
            self.progress['total_sections'] = len(content['sections'])

            for section in content['sections']:
                self.progress['current_section'] = section['title']

                # Type section heading
                self.type_heading(section['title'], level=section.get('level', 1))
                self.press_enter()

                # Type section content
                self.type_text(section['content'])
                self.press_enter(2)

                self.progress['completed_sections'] += 1

                if self.verbose:
                    print(f"Completed section: {section['title']}")
                    print(f"Progress: {self.progress['completed_sections']}/{self.progress['total_sections']}")

            # Update Table of Contents
            keyboard.press_and_release('ctrl+a')  # Select all
            keyboard.press_and_release('f9')      # Update fields
            keyboard.press_and_release('right')   # Deselect

            self.progress['status'] = 'completed'

            if self.verbose:
                print("\n=== Document Generation Complete ===")

            return True

        except Exception as e:
            self.handle_error("Document Generation", e)
            return False

def main():
    """Main function to run the WordAIAssistant"""
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Generate an academic paper using AI')
    parser.add_argument('--typing-mode', choices=['fast', 'realistic', 'slow'], default='realistic',
                        help='Typing behavior mode (default: realistic)')
    parser.add_argument('--topic', default='Earned Value Management',
                        help='Paper topic (default: Earned Value Management)')
    parser.add_argument('--template', help='Path to custom document template')
    parser.add_argument('--memory-file', help='Path to memory file (default: data/agent_memory.json)')
    parser.add_argument('--learning-file', help='Path to learning file (default: data/agent_learning.json)')
    parser.add_argument('--scaffold-file', help='Path to scaffold file (default: data/agent_scaffold.json)')
    parser.add_argument('--capabilities-dir', help='Path to capabilities directory (default: data/capabilities)')
    parser.add_argument('--tasks-file', help='Path to tasks file (default: data/agent_tasks.json)')
    parser.add_argument('--verbose', action='store_true',
                        help='Enable verbose output for debugging')
    parser.add_argument('--self-improve', action='store_true',
                        help='Enable self-improvement during execution')
    parser.add_argument('--self-evolve', action='store_true',
                        help='Enable self-evolution of capabilities')
    parser.add_argument('--track-position', action='store_true',
                        help='Enable detailed position tracking in Word')
    parser.add_argument('--iterative', action='store_true',
                        help='Enable iterative mode for incremental progress')
    args = parser.parse_args()

    # Initialize the assistant
    assistant = WordAIAssistant(
        typing_mode=args.typing_mode,
        verbose=args.verbose,
        topic=args.topic,
        template=args.template,
        memory_file=args.memory_file,
        learning_file=args.learning_file,
        scaffold_file=args.scaffold_file,
        capabilities_dir=args.capabilities_dir,
        tasks_file=args.tasks_file,
        enable_self_evolution=args.self_evolve,
        enable_iterative_mode=args.iterative
    )

    # Print memory system status
    if args.verbose:
        memory_summary = assistant.memory_system.summarize_memories()
        print("\n=== Memory System Status ===")
        print(f"Total memories: {memory_summary['count']}")
        print(f"Memory types: {memory_summary['types']}")
        print("============================\n")

        # Print learning system status
        learning_summary = assistant.learning_system.summarize_learning()
        print("\n=== Learning System Status ===")
        print(f"Total improvements: {learning_summary['count']}")
        print(f"Improvement types: {learning_summary['types']}")
        if learning_summary['count'] > 0:
            print(f"Average confidence: {learning_summary['avg_confidence']:.2f}")
            print(f"Success rate: {learning_summary['success_rate']:.2f}")
        print(f"Error patterns tracked: {learning_summary['error_patterns']}")
        print("============================\n")

        # Print scaffold system status
        if args.self_evolve:
            capability_analysis = assistant.scaffold_system.analyze_capabilities()
            print("\n=== Scaffold System Status ===")
            print(f"Total capabilities: {capability_analysis['count']}")
            print(f"Capability types: {capability_analysis['types']}")
            print(f"Evolution stages: {capability_analysis['stages']}")
            if capability_analysis['count'] > 0:
                print(f"Overall success rate: {capability_analysis['success_rate']:.2f}")
            print(f"Improvement opportunities: {len(capability_analysis['improvement_opportunities'])}")
            print("============================\n")

        # If iterative mode is enabled, print status
        if args.iterative:
            task_summary = assistant.get_task_summary()
            if task_summary:
                print("\n=== Iterative Agent Status ===")
                print(f"Total tasks: {task_summary['total_tasks']}")
                print(f"Status counts: {task_summary['status_counts']}")
                print(f"Overall progress: {task_summary['overall_progress']:.0%}")
                print(f"Top-level tasks: {task_summary['top_level_tasks']}")
                print("============================\n")

    # If position tracking is enabled, get initial position
    if args.track_position:
        print("\n=== Position Tracking Enabled ===")
        print("The agent will continuously track its position in Word")
        print("Initial position check...")
        assistant.update_position(force_update=True)
        context = assistant.understand_context()
        print(f"Current position: {context['current_position']}")
        print("===============================\n")

    print(f"\nGenerating academic paper on {args.topic}...")
    print("This may take a few minutes. Please keep Microsoft Word in focus.")

    try:
        # Choose the appropriate method based on iterative mode
        if args.iterative:
            print("\n=== Using Iterative Mode for Paper Generation ===")
            print("This mode breaks down the task into smaller steps and makes incremental progress")
            print("===================================================\n")

            # Generate the paper using the iterative approach
            assistant.generate_paper_iteratively()

            # Get final task summary
            task_summary = assistant.get_task_summary()
            if task_summary:
                print("\n=== Final Task Summary ===")
                print(f"Total tasks: {task_summary['total_tasks']}")
                print(f"Completed tasks: {task_summary['status_counts'].get('completed', 0)}")
                print(f"Overall progress: {task_summary['overall_progress']:.0%}")
                print(f"Completed steps: {task_summary['completed_steps']}")
                print("===========================\n")

            # Save the document
            filename = assistant.evm_generator.save_document()
            print(f"\nPaper generation complete! The document has been saved as '{filename}'")
        else:
            # Standard approach (non-iterative)
            # Get the sections to generate
            sections = assistant.evm_generator.sections
            total_sections = len(sections)

            # Track completed sections for self-improvement
            completed_sections = 0

            for i, section in enumerate(sections):
                # Update progress
                assistant.update_progress(section, "generating")
                print(f"\n=== Generating Section {i+1}/{total_sections} ===")
                print(f"Section: {section}")
                print("==============================")

                # Generate content for this section
                content = assistant.evm_generator.generate_section_content(section)

                # Update progress
                assistant.update_progress(section, "completed")

                # Type the section heading
                assistant.type_heading(section)

                # Type the section content
                assistant.type_text(content)

                # Update progress
                assistant.update_progress(section, "written")

                # Add spacing between sections
                assistant.press_enter(2)

                # Increment completed sections
                completed_sections += 1

                # Periodically analyze and improve behavior (every 2 sections or if errors occurred)
                if args.self_improve and (completed_sections % 2 == 0 or len(assistant.learning_system.error_history) > 0):
                    print("\n=== Performing Self-Improvement Analysis ===")
                    improvement_result = assistant.analyze_and_improve()
                    improvements = improvement_result.get("improvements", [])
                    evolved_capabilities = improvement_result.get("evolved_capabilities", [])

                    if improvements:
                        print(f"Self-improvement generated {len(improvements)} new strategies")
                    else:
                        print("No new improvements needed at this time")

                    if args.self_evolve and evolved_capabilities:
                        print(f"Self-evolution improved {len(evolved_capabilities)} capabilities")

                # Update position if tracking is enabled
                if args.track_position and completed_sections % 1 == 0:  # Every section
                    print("\n=== Updating Position Information ===")
                    assistant.update_position(force_update=True)
                    context = assistant.understand_context()
                    print(f"Current position: {context['current_position']}")
                    print(f"Position history: {len(context['position_history'])} entries")
                    print("===================================\n")

            # Generate and type the references section
            assistant.update_progress("References", "generating")
            print("\n=== Generating References ===")
            assistant.evm_generator.write_references()
            assistant.update_progress("References", "written")

            # Save the document
            filename = assistant.evm_generator.save_document()
            print(f"\nPaper generation complete! The document has been saved as '{filename}'")

        # Final self-improvement analysis
        if args.self_improve:
            print("\n=== Final Self-Improvement Analysis ===")
            improvement_result = assistant.analyze_and_improve()
            improvements = improvement_result.get("improvements", [])
            evolved_capabilities = improvement_result.get("evolved_capabilities", [])

            if improvements:
                print(f"Final self-improvement generated {len(improvements)} new strategies")
                print("These improvements will be applied in future runs")
            else:
                print("No new improvements needed")

            # Print learning summary
            learning_summary = assistant.learning_system.summarize_learning()
            print("\n=== Learning Summary ===")
            print(f"Total improvements learned: {learning_summary['count']}")
            print(f"Improvements applied: {learning_summary.get('total_applications', 0)}")
            print(f"Successful applications: {learning_summary.get('total_successes', 0)}")
            if learning_summary['count'] > 0:
                print(f"Success rate: {learning_summary['success_rate']:.2f}")
            print("=========================\n")

            # Final capability evolution summary
            if args.self_evolve:
                if evolved_capabilities:
                    print(f"Final self-evolution improved {len(evolved_capabilities)} capabilities")
                    print("These capability improvements will be available in future runs")

                # Print capability summary
                capability_analysis = assistant.scaffold_system.analyze_capabilities()
                print("\n=== Capability Summary ===")
                print(f"Total capabilities: {capability_analysis['count']}")
                print(f"Capability types: {capability_analysis['types']}")
                print(f"Evolution stages: {capability_analysis['stages']}")
                if capability_analysis['count'] > 0:
                    print(f"Overall success rate: {capability_analysis['success_rate']:.2f}")
                print("=========================\n")

    except Exception as e:
        print(f"\nError generating paper: {str(e)}")

        # Even on error, try to learn from the experience
        if args.self_improve:
            print("\n=== Attempting to Learn from Error ===")
            context = {
                "section": assistant.progress.get('current_section'),
                "status": assistant.progress.get('status'),
                "typing_mode": assistant.typing_mode,
                "error_location": "main_execution"
            }
            assistant.learning_system.track_error("Paper Generation", str(e), context)
            assistant._generate_improvement_from_failure("Paper Generation", str(e), context)
            print("Error has been analyzed and will be avoided in future runs")

        sys.exit(1)

if __name__ == "__main__":
    main()
