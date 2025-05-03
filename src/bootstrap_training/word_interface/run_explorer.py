"""
Microsoft Word Interface Explorer Runner

This script runs the self-learning exploration of Microsoft Word's interface,
allowing the AI Agent to physically interact with and demonstrate Word features.
"""

import os
import sys
import argparse
import time
import asyncio
from pathlib import Path

# Ensure the src directory is in the Python path
src_path = Path(__file__).parent.parent.parent
sys.path.insert(0, str(src_path.parent))

from src.bootstrap_training.word_interface.home_tab.explorer import HomeTabExplorer
from src.core.logging import setup_logger, get_logger

# Initialize logger
logger = setup_logger(__name__)

def parse_arguments():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description='Microsoft Word Interface Explorer')
    
    # Mode options
    mode_group = parser.add_argument_group('Mode')
    mode_group.add_argument('--explore-all', action='store_true', 
                         help='Explore all Home tab elements')
    mode_group.add_argument('--demonstrate', action='store_true',
                         help='Run full interactive demonstration')
    mode_group.add_argument('--element', type=str,
                         help='Demonstrate a specific element (e.g., "Bold")')
    
    # Cursor options
    cursor_group = parser.add_argument_group('Cursor')
    cursor_group.add_argument('--robot-cursor', action='store_true', 
                           help='Show robot cursor during exploration')
    cursor_group.add_argument('--no-robot-cursor', action='store_true', default=True,
                           help='Hide robot cursor')
    cursor_group.add_argument('--cursor-size', choices=['standard', 'large', 'extra_large'],
                           default='large', help='Size of robot cursor')
    
    # Other options
    parser.add_argument('--calibrate', action='store_true',
                      help='Force calibration of element positions')
    parser.add_argument('--delay', type=float, default=3.0,
                      help='Delay before starting (seconds)')
    parser.add_argument('--no-dialogs', action='store_true',
                      help='Disable dialog boxes and show errors in Word document only')
    
    return parser.parse_args()

def main():
    """Main function"""
    args = parse_arguments()
    
    # Determine if robot cursor should be shown
    use_robot_cursor = args.robot_cursor and not args.no_robot_cursor
    
    try:
        # Show startup message
        print("\nMicrosoft Word Interface Explorer")
        print("================================\n")
        print(f"Starting in {args.delay} seconds...")
        print("Please ensure Microsoft Word is installed and ready.")
        print("The AI Agent will physically move the cursor to explore Word's interface.")
        time.sleep(args.delay)
        
        # Initialize the explorer
        explorer = HomeTabExplorer(
            use_robot_cursor=use_robot_cursor,
            cursor_size=args.cursor_size
        )
        
        # Connect to Word
        if not explorer.connect_to_word():
            error_msg = "\nFailed to connect to Microsoft Word. Please ensure Word is installed."
            print(error_msg)
            
            # If dialogs are enabled, show error dialog
            if not args.no_dialogs:
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Connection Failed", error_msg)
                root.destroy()
            
            return 1
        
        # Calibrate if requested or needed
        if args.calibrate or not explorer._home_tab_elements:
            print("\nCalibrating element positions...")
            explorer.calibrate_element_positions()
        
        # Execute the requested action
        if args.demonstrate:
            print("\nRunning interactive demonstration of all Home tab elements...")
            explorer.interactive_demonstration()
            print("\nDemonstration complete! A document has been saved to your Desktop.")
        
        elif args.explore_all:
            print("\nExploring all Home tab elements...")
            grouped_elements = explorer.explore_all_elements()
            print("\nExploration complete! The elements have been documented in Word.")
            
            # Print summary to console
            print("\nSummary of Home Tab Elements:")
            for group, elements in grouped_elements.items():
                print(f"\n{group} Group:")
                for element in elements:
                    print(f"  - {element}")
        
        elif args.element:
            element_name = args.element
            explanation = explorer.explain_element(element_name)
            if explanation:
                print(f"\n{explanation}")
                
                print(f"\nDemonstrating {element_name}...")
                success = explorer.demonstrate_element(element_name)
                
                if success:
                    print(f"\nDemonstration of {element_name} complete!")
                else:
                    error_msg = f"\nFailed to demonstrate {element_name}."
                    print(error_msg)
                    
                    # If dialogs are enabled, show error dialog
                    if not args.no_dialogs:
                        import tkinter as tk
                        from tkinter import messagebox
                        root = tk.Tk()
                        root.withdraw()
                        messagebox.showerror("Demonstration Failed", error_msg)
                        root.destroy()
                    
                    return 1
            else:
                error_msg = f"\nElement '{element_name}' not found."
                print(error_msg)
                
                # If dialogs are enabled, show error dialog
                if not args.no_dialogs:
                    import tkinter as tk
                    from tkinter import messagebox
                    root = tk.Tk()
                    root.withdraw()
                    messagebox.showerror("Element Not Found", error_msg)
                    root.destroy()
                
                return 1
        
        else:
            # Default to a simple demonstration of one element
            demo_element = "Bold"  # Default element to demonstrate
            print(f"\nNo specific action requested. Demonstrating {demo_element} as an example...")
            
            success = explorer.demonstrate_element(demo_element)
            
            if success:
                print(f"\nDemonstration of {demo_element} complete!")
                print("\nTry running with --explore-all or --demonstrate for a full demonstration.")
            else:
                error_msg = f"\nFailed to demonstrate {demo_element}."
                print(error_msg)
                
                # If dialogs are enabled, show error dialog
                if not args.no_dialogs:
                    import tkinter as tk
                    from tkinter import messagebox
                    root = tk.Tk()
                    root.withdraw()
                    messagebox.showerror("Demonstration Failed", error_msg)
                    root.destroy()
                
                return 1
        
        # Clean up
        explorer.close()
        return 0
        
    except Exception as e:
        error_msg = f"\nError in Word Interface Explorer: {e}"
        logger.error(error_msg)
        print(error_msg)
        
        # If dialogs are enabled, show error dialog
        if hasattr(args, 'no_dialogs') and not args.no_dialogs:
            try:
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Explorer Error", f"An error occurred: {str(e)}")
                root.destroy()
            except Exception:
                pass  # If we can't show dialog, just continue
        
        return 1

if __name__ == "__main__":
    sys.exit(main())
