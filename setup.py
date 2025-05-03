"""
Setup script for the Word AI Agent

Provides installation and configuration functionality.
"""

import os
import sys
import shutil
import argparse
import subprocess
from pathlib import Path

def parse_arguments():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description='Word AI Agent Setup')
    parser.add_argument('--install', action='store_true', help='Install the Word AI Agent')
    parser.add_argument('--upgrade', action='store_true', help='Upgrade existing installation')
    parser.add_argument('--configure', action='store_true', help='Configure the Word AI Agent')
    parser.add_argument('--api-key', help='OpenAI API key')
    parser.add_argument('--venv', action='store_true', help='Create virtual environment')
    parser.add_argument('--no-deps', action='store_false', dest='install_deps', help='Skip dependency installation')
    parser.add_argument('--advanced', action='store_true', help='Install advanced dependencies')
    
    return parser.parse_args()

def create_virtual_environment():
    """Create a virtual environment"""
    print("Creating virtual environment...")
    
    venv_path = Path(".venv")
    if venv_path.exists():
        print("Virtual environment already exists")
        return True
    
    try:
        subprocess.run([sys.executable, "-m", "venv", ".venv"], check=True)
        print("Virtual environment created successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error creating virtual environment: {e}")
        return False

def install_dependencies(advanced=False):
    """Install dependencies"""
    print("Installing dependencies...")
    
    # Determine pip executable
    pip_exec = ".venv/bin/pip" if os.name != "nt" else ".venv\\Scripts\\pip.exe"
    if not Path(pip_exec).exists():
        pip_exec = "pip"
    
    # Install basic dependencies
    try:
        subprocess.run([pip_exec, "install", "-r", "requirements.txt"], check=True)
        print("Basic dependencies installed successfully")
    except subprocess.CalledProcessError as e:
        print(f"Error installing basic dependencies: {e}")
        return False
    
    # Install advanced dependencies if requested
    if advanced:
        try:
            print("Installing advanced dependencies...")
            advanced_deps = [
                "numpy",
                "pandas",
                "scikit-learn",
                "matplotlib",
                "nltk",
                "pyyaml",
                "markdown",
                "cryptography",
                "pytest",
                "pytest-asyncio"
            ]
            subprocess.run([pip_exec, "install"] + advanced_deps, check=True)
            print("Advanced dependencies installed successfully")
        except subprocess.CalledProcessError as e:
            print(f"Error installing advanced dependencies: {e}")
            return False
    
    return True

def configure_api_key(api_key=None):
    """Configure API key"""
    env_file = Path(".env")
    
    # Get API key if not provided
    if not api_key:
        print("\nOpenAI API Key Configuration")
        print("---------------------------")
        print("The Word AI Agent requires an OpenAI API key.")
        api_key = input("Enter your OpenAI API key: ")
    
    # Validate API key (simple check)
    if not api_key or len(api_key) < 10:
        print("Invalid API key")
        return False
    
    # Write to .env file
    try:
        with open(env_file, "w") as f:
            f.write(f"OPENAI_API_KEY=\"{api_key}\"\n")
        print("API key configured successfully")
        return True
    except Exception as e:
        print(f"Error writing API key to .env file: {e}")
        return False

def create_directory_structure():
    """Create directory structure"""
    print("Creating directory structure...")
    
    directories = [
        "config",
        "logs",
        "docs",
        "output",
        "data",
        "data/templates",
        "data/memory"
    ]
    
    for directory in directories:
        path = Path(directory)
        if not path.exists():
            path.mkdir(parents=True, exist_ok=True)
            print(f"Created directory: {directory}")
    
    return True

def create_config_file():
    """Create default configuration file"""
    print("Creating default configuration file...")
    
    config_file = Path("config/default.yaml")
    if config_file.exists():
        print("Configuration file already exists")
        return True
    
    try:
        with open(config_file, "w") as f:
            f.write("""# Word AI Agent Configuration

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
""")
        print("Configuration file created successfully")
        return True
    except Exception as e:
        print(f"Error creating configuration file: {e}")
        return False

def install():
    """Install the Word AI Agent"""
    print("\nWord AI Agent Installation")
    print("=========================")
    
    # Create directory structure
    if not create_directory_structure():
        print("Installation failed: Could not create directory structure")
        return False
    
    # Create default configuration file
    if not create_config_file():
        print("Installation failed: Could not create configuration file")
        return False
    
    print("\nWord AI Agent installed successfully!")
    print("\nNext steps:")
    print("1. Install dependencies: python setup.py --install-deps")
    print("2. Configure API key: python setup.py --configure")
    print("3. Run the agent: python run_launcher.py")
    
    return True

def upgrade():
    """Upgrade existing installation"""
    print("\nWord AI Agent Upgrade")
    print("====================")
    
    # Backup configuration files
    print("Backing up configuration files...")
    config_path = Path("config")
    backup_path = Path("config_backup")
    
    if config_path.exists():
        if backup_path.exists():
            shutil.rmtree(backup_path)
        
        try:
            shutil.copytree(config_path, backup_path)
            print("Configuration files backed up successfully")
        except Exception as e:
            print(f"Error backing up configuration files: {e}")
            return False
    
    # Update directory structure
    if not create_directory_structure():
        print("Upgrade failed: Could not update directory structure")
        return False
    
    print("\nWord AI Agent upgraded successfully!")
    print("Your configuration files have been backed up to config_backup/")
    
    return True

def main():
    """Main function"""
    args = parse_arguments()
    
    # Create virtual environment if requested
    if args.venv:
        if not create_virtual_environment():
            return 1
    
    # Install dependencies if requested
    if args.install_deps:
        if not install_dependencies(args.advanced):
            return 1
    
    # Install the agent if requested
    if args.install:
        if not install():
            return 1
    
    # Upgrade the agent if requested
    if args.upgrade:
        if not upgrade():
            return 1
    
    # Configure API key if requested
    if args.configure or args.api_key:
        if not configure_api_key(args.api_key):
            return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
