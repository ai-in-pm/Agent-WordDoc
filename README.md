# Agent WordDoc

This AI-powered system generates comprehensive academic papers by directly typing into Microsoft Word with human-like behavior, leveraging advanced NLP models and a modular architecture.

![Agent WordDoc Icon](src/images/agent_worddoc_icon.png)

## Enhanced Features

### Core Architecture
- **Modular Design**: Fully modularized codebase with proper separation of concerns
- **Asynchronous Processing**: Leverages asyncio for non-blocking operations
- **Plugin System**: Extensible plugin architecture for custom capabilities
- **Dependency Injection**: Factory patterns for better component management
- **Comprehensive Logging**: Advanced logging system with rotation and levels
- **Performance Monitoring**: Real-time metrics and performance tracking
- **Security Features**: Input validation, credential management, and API protection

### User Interface Options
- **GUI Interface**: User-friendly graphical interface with real-time progress updates
- **CLI Interface**: Command-line interface with comprehensive options
- **Interactive Mode**: Step-by-step guided configuration and execution
- **Service Mode**: Background service capability for automated operation

### Document Generation
- **Multiple Document Formats**: Support for Word, PDF, LaTeX, and plain text
- **Template System**: Customizable document templates and styles
- **Version Control**: Document history and version management
- **Real-time Editing**: Live document editing with real-time validation
- **Collaborative Editing**: Support for multiple concurrent editors

### Enhanced AI Capabilities
- **Advanced NLP Models**: Integration with cutting-edge AI models
- **Self-Improvement System**: Learns from mistakes and optimizes behavior
- **Self-Evolution**: Evolves capabilities based on usage patterns
- **Context Awareness**: Maintains context across document sections
- **Automatic Research**: Integrates research capabilities for citations

### Microsoft Word Integration

The Word AI Agent can directly interact with Microsoft Word to create and edit documents:

```bash
# Generate content and automatically create a Word document
python run_launcher.py --topic "Quantum Computing" --use-word

# Use a specific template and save to a specified file
python run_launcher.py --topic "Machine Learning" --template-id research-paper --use-word --word-doc "path/to/output.docx"

# Use a specific output path without specifying word-doc
python run_launcher.py --topic "Data Science" --use-word --output "path/to/save.docx"
```

#### Word Automation Requirements

To use the Word automation features, ensure you have:
- Microsoft Word installed
- Python dependencies: `pywin32`, `pyautogui`

Install the dependencies with:
```bash
pip install pywin32 pyautogui
```

#### Word Automation Features

- Automatically opens Microsoft Word
- Types content directly into Word with realistic typing patterns
- Applies document formatting based on selected templates
- Can insert proper document structure with headings and sections
- Save documents to specified locations

#### Robot Cursor

The AI Agent displays a distinctive robot cursor when it's controlling your mouse and keyboard:

```bash
# Enable robot cursor (enabled by default)
python run_launcher.py --topic "Quantum Computing" --use-word --robot-cursor

# Disable robot cursor
python run_launcher.py --topic "Quantum Computing" --use-word --no-robot-cursor

# Adjust cursor size (standard, large, extra_large)
python run_launcher.py --topic "Quantum Computing" --use-word --cursor-size large
```

This visual indicator makes it clear when the AI is actively controlling your system vs. when you have control.

### Interactive Mode

Launch in interactive mode for guided execution:

```bash
python run_launcher.py --interactive
```

### Service Mode

Run as a background service:

```bash
python run_launcher.py --service
```

### Available Options

```
--gui                 Run with GUI interface
--cli                 Run with command-line interface
--interactive         Run in interactive mode
--service             Run as a background service
--topic TOPIC         Paper topic (default: Earned Value Management)
--config CONFIG       Path to configuration file
--typing-mode MODE    Typing behavior mode (fast, realistic, slow)
--verbose             Enable verbose output
--iterative           Enable iterative processing
--self-improve        Enable self-improvement
--self-evolve         Enable self-evolution
--track-position      Enable position tracking
--robot-cursor        Show robot cursor (default: enabled)
--no-robot-cursor     Disable robot cursor
--use-autoit          Use AutoIt (default: enabled)
--no-autoit           Disable AutoIt
--delay DELAY         Delay before starting (seconds)
--max-retries MAX     Maximum retries on error
--retry-delay DELAY   Delay between retries (seconds)
--log-level LEVEL     Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
--plugins-dir DIR     Plugins directory
```

### Voice Command System

The Agent WordDoc now supports voice commands using ElevenLabs voice recognition and text-to-speech capabilities:

```bash
# Enable voice commands
python run_launcher.py --voice-commands

# Enable voice commands with spoken responses
python run_launcher.py --voice-commands --voice-speak-responses

# Launch with voice tutorial (explains available commands)
python run_launcher.py --voice-commands --voice-speak-responses --first-run

# Set voice command language (default: en-US)
python run_launcher.py --voice-commands --voice-language en-US
```

#### Available Voice Commands

**Basic Controls:**
- "start writing" - Begin writing process
- "stop writing" - Stop writing process
- "stop the agent from typing now" - Emergency stop (immediately halts typing)
- "pause" - Pause the agent
- "continue" or "resume" - Continue writing

**Document Creation:**
- "write about [topic]" - Start writing about a specific topic
- "change topic to [topic]" - Change the current topic
- "add section [name]" - Add a new section with specified name
- "add introduction" - Add an introduction section
- "add conclusion" - Add a conclusion section
- "add references" - Add a references section

**Navigation:**
- "go to top" or "go to start" - Navigate to document beginning
- "go to bottom" or "go to end" - Navigate to document end
- "go to section [name]" - Navigate to a specific section

**Editing:**
- "delete last paragraph" - Delete the previous paragraph
- "undo" - Undo last action
- "redo" - Redo last undone action

**Formatting:**
- "format as [style]" - Apply formatting
- "make [text] bold" - Apply bold formatting to text

**Help Commands:**
- "help" - Show available commands
- "help with [category] commands" - Get help with specific command categories
- "list commands" - List all available commands

**System Commands:**
- "save document" - Save the current document
- "set typing speed to [fast/realistic/slow]" - Change typing speed
- "enable/disable [feature]" - Toggle features like self-improvement
- "exit" - Exit the application

#### Voice Command Requirements

To use the voice command features, ensure you have:
- ElevenLabs API key (set in .env file)
- Python dependencies: `pyaudio`, `elevenlabs`

Install the dependencies with:
```bash
pip install pyaudio elevenlabs
```

### Word Interface Explorer

The Word Interface Explorer allows the AI Agent to physically demonstrate Microsoft Word's Home tab features with visible cursor movements and real-time interaction:

```bash
# Launch the Word Interface Explorer GUI
python src/bootstrap_training/word_interface/home_tab/gui.py
```

You can also use the provided shortcut:
```
C:\Users\USERNAME\OneDrive\Desktop\AI_Shortcuts\AgentWordDoc\AgentWordDoc_WordExplorer.lnk
```

#### Explorer Features

**Demonstration Modes:**
- **Full Interactive Demonstration**: Complete demonstration of all Home tab elements with physical cursor movements
- **Explore All Elements**: Document all Home tab elements with descriptions and examples
- **Demonstrate Specific Element**: Shows how to use a specific element (e.g., Bold, Italic, Underline)

**Execution Options:**
- **Start Exploration**: Launch the demonstration with a configurable delay
- **Execute Now**: Immediately run the demonstration with minimal delay

**User Verification System:**
- After each demonstration, the AI Agent asks if it successfully completed the command
- User provides verification by clicking 'Yes' (success) or 'No' (failure) 
- If verification fails, user can click the "Retry" button to let the AI Agent try again
- This ensures accurate feedback and learning from demonstration attempts

**Adaptive Retry System:**
- When failures occur, the AI Agent analyzes logs to identify issues
- Click "Retry" to have the AI learn from mistakes and adapt its approach
- The learning process is displayed through intuitive dialog boxes

**Visibility Options:**
- Show robot cursor during demonstrations (configurable size)
- Force element position calibration when needed
- Adjust startup delay before demonstrations begin

#### Explorer Requirements

To use the Word Interface Explorer, ensure you have:
- Microsoft Word installed
- Python dependencies: `pywin32`, `pyautogui`, `pillow`
- Optional: `pytesseract` for OCR capabilities (improves visual recognition)

Install the dependencies with:
```bash
pip install pywin32 pyautogui pillow pytesseract
```

For Tesseract OCR (optional):
1. Download and install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
2. Add the Tesseract installation directory to your PATH

#### Shortcuts

The following shortcuts are available in the `AI_Shortcuts/AgentWordDoc` folder:

- **AgentWordDoc.lnk** - Default settings with all enhancements
- **AgentWordDoc_Advanced.lnk** - Full-featured experience with advanced capabilities
- **AgentWordDoc_Fast.lnk** - Fast typing mode with all enhancements
- **AgentWordDoc_Business.lnk** - Pre-configured with Business Strategy topic
- **AgentWordDoc_Voice.lnk** - Voice-first experience with spoken responses
- **AgentWordDoc_VoiceTutorial.lnk** - Interactive voice tutorial for learning commands
- **AgentWordDoc_VisibleCursor.lnk** - Extra large robot cursor for maximum visibility
- **AgentWordDoc_WordExplorer.lnk** - Word Interface Explorer for demonstrating Home tab features

#### Robot Cursor

The AI Agent displays a distinctive robot cursor when it's controlling your mouse and keyboard:

```bash
# Enable robot cursor (enabled by default)
python run_launcher.py --topic "Quantum Computing" --use-word --robot-cursor

# Disable robot cursor
python run_launcher.py --topic "Quantum Computing" --use-word --no-robot-cursor

# Adjust cursor size (standard, large, extra_large)
python run_launcher.py --topic "Quantum Computing" --use-word --cursor-size large
```

This visual indicator makes it clear when the AI is actively controlling your system vs. when you have control.

## Directory Structure

```
word-ai-agent/
|-- src/
|   |-- core/            # Core system components
|   |-- agents/          # AI agent implementations
|   |-- services/        # Service layer components
|   |-- utils/           # Utility functions and tools
|   |-- handlers/        # Event and error handlers
|   |-- interfaces/      # User interface components
|   |-- config/          # Configuration management
|   |-- plugins/         # Plugin system components
|   `-- tests/           # Comprehensive test suite
|-- config/              # Configuration files
|-- data/                # Data files and templates
|-- docs/                # Documentation files
|-- logs/                # Log files
|-- output/              # Generated documents
|-- plugins/             # User plugins
|-- run_launcher.py      # Main entry point
|-- run_word_agent.py    # Legacy entry point
|-- setup.py             # Installation script
`-- requirements.txt     # Dependencies
```

## Installation

### Standard Installation

1. Clone the repository

2. Create a virtual environment:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   source .venv/bin/activate  # Unix/MacOS
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the setup script:
   ```bash
   python setup.py --install --configure
   ```

5. Enter your OpenAI API key when prompted or create a .env file with:
   ```
   OPENAI_API_KEY="your_api_key_here"
   ```

### Advanced Installation

For advanced features, install additional dependencies:

```bash
python setup.py --install --advanced --venv
```

## Usage

### GUI Mode

Launch with the graphical user interface:

```bash
python run_launcher.py --gui
```

### CLI Mode

Generate a paper on a specific topic:

```bash
python run_launcher.py --topic "Quantum Computing" --typing-mode realistic
```

Use a specific template for your paper:

```bash
python run_launcher.py --topic "Machine Learning" --template-id research-paper
```

Available template IDs include:
- `research-paper`: Standard academic research paper
- `literature-review`: Comprehensive literature review
- `case-study`: In-depth case analysis
- `thesis`: Thesis or dissertation format
- `technical-report`: Technical report format
- `journal-article`: Journal publication format
- `systematic-review`: Systematic literature review
- `conference-paper`: Conference submission format

### Interactive Mode

Launch in interactive mode for guided execution:

```bash
python run_launcher.py --interactive
```

### Service Mode

Run as a background service:

```bash
python run_launcher.py --service
```

### Available Options

```
--gui                 Run with GUI interface
--cli                 Run with command-line interface
--interactive         Run in interactive mode
--service             Run as a background service
--topic TOPIC         Paper topic (default: Earned Value Management)
--config CONFIG       Path to configuration file
--typing-mode MODE    Typing behavior mode (fast, realistic, slow)
--verbose             Enable verbose output
--iterative           Enable iterative processing
--self-improve        Enable self-improvement
--self-evolve         Enable self-evolution
--track-position      Enable position tracking
--robot-cursor        Show robot cursor (default: enabled)
--no-robot-cursor     Disable robot cursor
--use-autoit          Use AutoIt (default: enabled)
--no-autoit           Disable AutoIt
--delay DELAY         Delay before starting (seconds)
--max-retries MAX     Maximum retries on error
--retry-delay DELAY   Delay between retries (seconds)
--log-level LEVEL     Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
--plugins-dir DIR     Plugins directory
```

## Configuration

The system can be configured using YAML files in the config directory. Create environment-specific configurations in subfolders.

Example configuration:

```yaml
# API Configuration
api_key: ""  # Will be loaded from .env

# Typing Configuration
typing_mode: "realistic"  # Options: fast, realistic, slow

# Feature Flags
verbose: true
iterative: true
self_improve: true
self_evolve: true
track_position: true
robot_cursor: true
use_autoit: true

# Performance Configuration
delay: 3.0
max_retries: 3
retry_delay: 1.0

# Logging Configuration
log_level: "INFO"

# Paths
output_directory: "output"
data_directory: "data"
templates_directory: "data/templates"
memory_directory: "data/memory"
```

## Plugin System

Create custom plugins to extend the system's functionality.

Example plugin:

```python
# plugins/my_plugin.py

__version__ = "1.0.0"
__description__ = "My custom plugin"
__author__ = "Your Name"

def my_function(**kwargs):
    # Plugin logic here
    return result

def register_plugin(plugin_manager):
    # Register hooks
    plugin_manager.register_hook("before_processing", my_function)
```

## Testing

Run the test suite:

```bash
python -m pytest src/tests
```

Run specific tests:

```bash
python -m pytest src/tests/test_word_agent.py
```

## Documentation

Generate API documentation:

```bash
python -m src.utils.documentation_generator
```

Access the documentation in the docs directory.

## Security

- API keys are securely stored using environment variables or encrypted storage
- Input validation prevents injection attacks
- Rate limiting protects against API abuse
- Credential management secures sensitive information

## Performance

The system includes built-in performance monitoring. View metrics:

```python
from src.utils.performance_monitor import PerformanceMonitor

monitor = PerformanceMonitor()
metrics = monitor.get_metrics_summary()
print(metrics)
```

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch
3. Add your changes
4. Run tests
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.