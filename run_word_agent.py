"""
Run Word AI Agent with Real-Time Typing

This script runs the Word AI Agent with real-time typing, simulating human-like interaction.
It uses the OS interaction module to type commands and interact with the operating system.
"""

import os
import sys
import argparse
import time

from src.os_interaction import OSInteraction, TypingMode


def parse_arguments():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description='Run Word AI Agent with real-time typing')
    parser.add_argument('--topic', default='Earned Value Management',
                        help='Paper topic (default: Earned Value Management)')
    parser.add_argument('--typing-mode', choices=['fast', 'realistic', 'slow'], default='realistic',
                        help='Typing behavior mode (default: realistic)')
    parser.add_argument('--verbose', action='store_true',
                        help='Enable verbose output for debugging')
    parser.add_argument('--iterative', action='store_true',
                        help='Enable iterative mode for incremental progress')
    parser.add_argument('--self-improve', action='store_true',
                        help='Enable self-improvement during execution')
    parser.add_argument('--self-evolve', action='store_true',
                        help='Enable self-evolution of capabilities')
    parser.add_argument('--track-position', action='store_true',
                        help='Enable detailed position tracking in Word')
    parser.add_argument('--robot-cursor', action='store_true', default=True,
                        help='Show robot cursor during AI control (default: True)')
    parser.add_argument('--no-robot-cursor', action='store_false', dest='robot_cursor',
                        help='Disable robot cursor during AI control')
    parser.add_argument('--control-ui', action='store_true', default=True,
                        help='Show control UI with Pause/Stop/Continue buttons (default: True)')
    parser.add_argument('--no-control-ui', action='store_false', dest='control_ui',
                        help='Disable control UI')
    parser.add_argument('--use-autoit', action='store_true', default=True,
                        help='Use AutoIt for advanced Windows automation (default: True)')
    parser.add_argument('--no-autoit', action='store_false', dest='use_autoit',
                        help='Disable AutoIt integration')
    parser.add_argument('--delay', type=float, default=3.0,
                        help='Delay in seconds before starting (default: 3.0)')

    return parser.parse_args()


def run_word_agent(args):
    """Run the Word AI Agent with real-time typing"""
    try:
        # Initialize OS interaction
        os_interaction = OSInteraction(
            typing_mode=args.typing_mode,
            verbose=args.verbose,
            show_robot_cursor=args.robot_cursor,
            use_autoit=args.use_autoit,
            show_control_ui=args.control_ui
        )

        # Build command
        command = f"python -m src.word_ai_agent --topic \"{args.topic}\""

        # Add optional flags
        if args.verbose:
            command += " --verbose"
        if args.iterative:
            command += " --iterative"
        if args.self_improve:
            command += " --self-improve"
        if args.self_evolve:
            command += " --self-evolve"
        if args.track_position:
            command += " --track-position"

        print(f"Will run Word AI Agent in {args.delay} seconds with command:")
        print(command)
        print("\nPlease ensure you have a terminal/command prompt window focused.")

        if args.robot_cursor:
            print("The cursor will change to a robot icon when the AI Agent is in control.")

        if args.control_ui:
            print("Control UI is enabled. Use the Pause, Stop, and Continue buttons to control the AI Agent.")

        if args.use_autoit:
            print("Using AutoIt for advanced Windows automation (if available).")

        time.sleep(args.delay)

        # Launch command prompt or terminal
        if os.name == 'nt':  # Windows
            os_interaction.launch_application('cmd.exe')
            time.sleep(1)
        else:  # Unix-like
            os_interaction.launch_application('xterm')
            time.sleep(1)

        # Type and execute the command
        result = os_interaction.type_command(command)

        # Clean up resources
        os_interaction.cleanup()

        return result
    except Exception as e:
        print(f"Error running Word AI Agent: {str(e)}")
        return False


def main():
    """Main function"""
    args = parse_arguments()

    # Run the Word AI Agent
    success = run_word_agent(args)

    if success:
        print("Word AI Agent command typed successfully.")
    else:
        print("Failed to type Word AI Agent command.")
        sys.exit(1)


if __name__ == "__main__":
    main()
