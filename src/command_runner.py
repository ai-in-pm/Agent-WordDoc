"""
Command Runner for AI Agent

This module provides a command-line interface for running commands with realistic typing.
It uses the OS interaction module to simulate human-like typing and interaction.
"""

import os
import sys
import argparse
import time
from typing import List, Optional

from src.os_interaction import OSInteraction, TypingMode


def parse_arguments():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description='Run commands with realistic typing')
    parser.add_argument('command', nargs='*', help='Command to run')
    parser.add_argument('--typing-mode', choices=['fast', 'realistic', 'slow'], default='realistic',
                        help='Typing behavior mode (default: realistic)')
    parser.add_argument('--delay', type=float, default=1.0,
                        help='Delay in seconds before starting (default: 1.0)')
    parser.add_argument('--verbose', action='store_true',
                        help='Enable verbose output for debugging')
    parser.add_argument('--no-execute', action='store_true',
                        help='Type the command but do not execute it')
    
    return parser.parse_args()


def run_command(command: List[str], typing_mode: str = 'realistic', delay: float = 1.0, 
                verbose: bool = False, execute: bool = True) -> bool:
    """
    Run a command with realistic typing
    
    Args:
        command: Command to run as a list of strings
        typing_mode: Typing behavior mode ('fast', 'realistic', 'slow')
        delay: Delay in seconds before starting
        verbose: Whether to print debug information
        execute: Whether to execute the command after typing
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Initialize OS interaction
        os_interaction = OSInteraction(typing_mode=typing_mode, verbose=verbose)
        
        # Join command parts
        command_str = ' '.join(command)
        
        print(f"Will type command in {delay} seconds: {command_str}")
        time.sleep(delay)
        
        # Type and optionally execute the command
        return os_interaction.type_command(command_str, execute=execute)
    except Exception as e:
        print(f"Error running command: {str(e)}")
        return False


def run_python_module(module_path: str, args: str = "", typing_mode: str = 'realistic', 
                     delay: float = 1.0, verbose: bool = False) -> bool:
    """
    Run a Python module with realistic typing
    
    Args:
        module_path: Path to the Python module
        args: Command-line arguments
        typing_mode: Typing behavior mode ('fast', 'realistic', 'slow')
        delay: Delay in seconds before starting
        verbose: Whether to print debug information
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Initialize OS interaction
        os_interaction = OSInteraction(typing_mode=typing_mode, verbose=verbose)
        
        # Build command
        command = f"python -m {module_path}"
        if args:
            command += f" {args}"
        
        print(f"Will run Python module in {delay} seconds: {command}")
        time.sleep(delay)
        
        # Run the module
        return os_interaction.run_python_module(module_path, args, real_typing=True)
    except Exception as e:
        print(f"Error running Python module: {str(e)}")
        return False


def main():
    """Main function"""
    args = parse_arguments()
    
    if not args.command:
        print("No command specified. Use --help for usage information.")
        return
    
    # Run the command
    success = run_command(
        command=args.command,
        typing_mode=args.typing_mode,
        delay=args.delay,
        verbose=args.verbose,
        execute=not args.no_execute
    )
    
    if success:
        print("Command typed successfully.")
    else:
        print("Failed to type command.")
        sys.exit(1)


if __name__ == "__main__":
    main()
