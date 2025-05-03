import os
import sys
import tempfile
import shutil
import subprocess
from pathlib import Path
import urllib.request
import zipfile
import winreg
import ctypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def set_environment_variable(name, value):
    """Set environment variable at user level"""
    try:
        # Set for current process
        os.environ[name] = value
        
        # Set permanently at user level (doesn't require admin)
        subprocess.run(
            f'setx {name} "{value}"', 
            shell=True, 
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        print(f"[SUCCESS] Set {name}={value} in environment variables")
        return True
    except Exception as e:
        print(f"[ERROR] Failed to set environment variable {name}: {e}")
        return False

def get_tesseract_path():
    """Try to find Tesseract installation path"""
    try:
        # Try common installation paths
        common_paths = [
            r"C:\Program Files\Tesseract-OCR",
            r"C:\Program Files (x86)\Tesseract-OCR",
            os.path.join(os.environ.get('LOCALAPPDATA', ''), "Tesseract-OCR"),
            os.path.join(os.environ.get('APPDATA', ''), "Tesseract-OCR"),
        ]
        
        # Check if tesseract.exe exists in any of these paths
        for path in common_paths:
            exe_path = os.path.join(path, "tesseract.exe")
            if os.path.isfile(exe_path):
                return path
        
        # Try to check registry
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Tesseract-OCR") as key:
                return winreg.QueryValueEx(key, "InstallDir")[0]
        except Exception:
            pass
            
        # Try system PATH
        for path_dir in os.environ.get('PATH', '').split(';'):
            exe_path = os.path.join(path_dir, "tesseract.exe")
            if os.path.isfile(exe_path):
                return path_dir
                
        return None
    except Exception as e:
        print(f"Error finding Tesseract: {e}")
        return None

def manual_install():
    """Guide for manual installation"""
    print("\n=== Manual Installation Guide ===\n")
    print("Since automatic installation isn't working, please follow these steps:")
    print("1. Download the Tesseract OCR installer from one of these links:")
    print("   - https://github.com/UB-Mannheim/tesseract/wiki")
    print("   - https://tesseract-ocr.github.io/tessdoc/Installation.html")
    print("\n2. During installation:")
    print("   - Make sure to check 'Add to PATH'")
    print("   - The default installation location (C:\\Program Files\\Tesseract-OCR) is recommended")
    print("   - Select additional languages if needed")
    
    print("\n3. After installation is complete:")
    print("   - Close and reopen any command prompt windows")
    print("   - Run this script again to verify the installation")
    
    print("\nChecking if you've already installed Tesseract...")
    
    # Check if Tesseract is already installed and in PATH
    tesseract_path = get_tesseract_path()
    if tesseract_path:
        print(f"[SUCCESS] Found Tesseract at: {tesseract_path}")
        return tesseract_path
    else:
        print("[INFO] Tesseract not found. Please install it following the steps above.")
        return None

def copy_source_to_tesseract():
    """Copy all files from our tesseract-main to the installed Tesseract"""
    tesseract_path = get_tesseract_path()
    if not tesseract_path:
        print("[ERROR] Cannot copy files - Tesseract installation not found")
        return False
    
    try:
        # Source directory
        source_dir = os.path.join(
            Path(__file__).parent.absolute(), 
            "tesseract-main"
        )
        
        if not os.path.isdir(source_dir):
            print(f"[WARNING] Source directory not found at {source_dir}")
            return False
        
        # Copy tessdata files specifically
        source_tessdata = os.path.join(source_dir, "tessdata")
        if os.path.isdir(source_tessdata):
            dest_tessdata = os.path.join(tesseract_path, "tessdata")
            os.makedirs(dest_tessdata, exist_ok=True)
            
            # Copy language files
            file_count = 0
            for file in os.listdir(source_tessdata):
                source_file = os.path.join(source_tessdata, file)
                if os.path.isfile(source_file) and file.endswith('.traineddata'):
                    dest_file = os.path.join(dest_tessdata, file)
                    shutil.copy2(source_file, dest_file)
                    print(f"Copied language file: {file}")
                    file_count += 1
            
            if file_count > 0:
                print(f"[SUCCESS] Successfully copied {file_count} language files to {dest_tessdata}")
            else:
                print("[INFO] No language files found to copy")
        
        # Try to copy all other relevant directories and files
        print("[INFO] Checking for other relevant files to copy...")
        for item in os.listdir(source_dir):
            source_item = os.path.join(source_dir, item)
            dest_item = os.path.join(tesseract_path, item)
            
            # Skip if it's tessdata (already handled) or non-essential directories
            if item == "tessdata" or item in [".git", ".github", "docs", "test", "unittest"]:
                continue
                
            try:
                if os.path.isdir(source_item):
                    # For directories like 'include', copy if not existing
                    if not os.path.exists(dest_item):
                        print(f"Copying directory {item}...")
                        shutil.copytree(source_item, dest_item)
                elif os.path.isfile(source_item) and item.lower().endswith(('.dll', '.exe', '.h', '.lib')):
                    # Only copy binary/library files if they don't exist
                    if not os.path.exists(dest_item):
                        print(f"Copying file {item}...")
                        shutil.copy2(source_item, dest_item)
            except Exception as e:
                print(f"[WARNING] Couldn't copy {item}: {e}")
        
        return True
    
    except Exception as e:
        print(f"[ERROR] Error copying source files: {e}")
        return False

def test_tesseract():
    """Test if Tesseract is working properly"""
    try:
        # Try to run tesseract --version
        result = subprocess.run(
            ["tesseract", "--version"], 
            capture_output=True, 
            text=True,
            check=True
        )
        
        print("\n[SUCCESS] Tesseract is working properly!")
        print(f"Version info:\n{result.stdout}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Tesseract test failed with error: {e.stderr}")
        return False
    except FileNotFoundError:
        print("[ERROR] Tesseract not found in PATH. Please make sure it's installed correctly.")
        return False
    except Exception as e:
        print(f"[ERROR] Error testing Tesseract: {e}")
        return False

def update_code_with_path():
    """Update the document_awareness.py file with the correct Tesseract path"""
    tesseract_path = get_tesseract_path()
    if not tesseract_path:
        print("[ERROR] Cannot update code - Tesseract installation not found")
        return False
    
    try:
        doc_awareness_file = os.path.join(
            Path(__file__).parents[1],
            "bootstrap_training",
            "word_interface",
            "document_awareness.py"
        )
        
        if not os.path.exists(doc_awareness_file):
            print(f"[ERROR] Cannot find the document_awareness.py file at {doc_awareness_file}")
            return False
        
        # Read the file
        with open(doc_awareness_file, 'r') as f:
            content = f.read()
        
        # Update the content with explicit Tesseract path
        tesseract_exe = os.path.join(tesseract_path, "tesseract.exe")
        
        # First look for specific pattern to update
        if 'pytesseract.pytesseract.tesseract_cmd' in content:
            updated_content = content.replace(
                'HAS_OCR = pytesseract is not None',
                f'HAS_OCR = pytesseract is not None\n\n# Set explicit path to Tesseract executable\nif HAS_OCR:\n    pytesseract.pytesseract.tesseract_cmd = r"{tesseract_exe}"'
            )
        else:
            # If pattern not found, find a good place to insert it
            # Look for import statements related to pytesseract
            import_lines = [i for i, line in enumerate(content.split('\n')) 
                          if 'import pytesseract' in line or 'from pytesseract' in line]
            
            if import_lines:
                # Get the line index after the last import
                line_idx = max(import_lines)
                lines = content.split('\n')
                
                # Insert our tesseract path code after that import
                lines.insert(line_idx + 1, f"\n# Set explicit path to Tesseract executable")
                lines.insert(line_idx + 2, f"if 'pytesseract' in locals() or 'pytesseract' in globals():")
                lines.insert(line_idx + 3, f"    pytesseract.pytesseract.tesseract_cmd = r\"{tesseract_exe}\"")
                
                updated_content = '\n'.join(lines)
            else:
                # No good place found, add to the start of the file with comment
                updated_content = (f"# Set explicit path to Tesseract executable\n"
                               f"try:\n"
                               f"    import pytesseract\n"
                               f"    pytesseract.pytesseract.tesseract_cmd = r\"{tesseract_exe}\"\n"
                               f"except ImportError:\n"
                               f"    pass  # pytesseract will be imported later\n\n"
                              ) + content
        
        # Backup the original file
        backup_path = doc_awareness_file + '.bak'
        shutil.copy2(doc_awareness_file, backup_path)
        print(f"[INFO] Created backup at {backup_path}")
        
        # Write the updated content
        with open(doc_awareness_file, 'w') as f:
            f.write(updated_content)
        
        print(f"[SUCCESS] Updated {doc_awareness_file} with Tesseract path: {tesseract_exe}")
        return True
    
    except Exception as e:
        print(f"[ERROR] Error updating code with Tesseract path: {e}")
        return False

def unpackage_tesseract():
    """Directly extract bundled Tesseract files to Program Files"""
    try:
        print("[INFO] Trying to use bundled Tesseract files...")
        
        # Source directory with embedded tesseract files
        source_dir = os.path.join(
            Path(__file__).parent.absolute(), 
            "tesseract-main"
        )
        
        # Default install location
        dest_dir = r"C:\Program Files\Tesseract-OCR"
        
        # Create destination directory if it doesn't exist
        os.makedirs(dest_dir, exist_ok=True)
        
        # Copy all relevant files from source to destination
        file_count = 0
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                if file.endswith(('.exe', '.dll', '.traineddata')):
                    src_file = os.path.join(root, file)
                    
                    # Get relative path from source_dir
                    rel_path = os.path.relpath(src_file, source_dir)
                    
                    # Create destination path
                    dst_file = os.path.join(dest_dir, rel_path)
                    
                    # Make sure destination directory exists
                    os.makedirs(os.path.dirname(dst_file), exist_ok=True)
                    
                    # Copy the file
                    shutil.copy2(src_file, dst_file)
                    print(f"Copied {rel_path}")
                    file_count += 1
        
        if file_count > 0:
            print(f"[SUCCESS] Copied {file_count} files to {dest_dir}")
            
            # Add to PATH if not already there
            if dest_dir not in os.environ.get('PATH', '').split(';'):
                set_environment_variable('PATH', os.environ.get('PATH', '') + f";{dest_dir}")
            
            return dest_dir
        else:
            print("[WARNING] No essential files found in the source directory")
            return None
    
    except Exception as e:
        print(f"[ERROR] Error unpacking Tesseract: {e}")
        return None

def main():
    print("=== Tesseract OCR Installation for AI Agent ===\n")
    
    # Check if Tesseract is already installed
    tesseract_path = get_tesseract_path()
    if tesseract_path:
        print(f"[SUCCESS] Tesseract OCR is already installed at: {tesseract_path}")
        # Just update paths and continue
    else:
        print("[INFO] Tesseract OCR is not installed or not found.")
        
        # Try to unpack bundled tesseract files
        tesseract_path = unpackage_tesseract()
        
        if not tesseract_path:
            # If unpacking failed, guide user through manual installation
            print("[INFO] Using manual installation approach.")
            tesseract_path = manual_install()
    
    if tesseract_path:
        # Copy any additional files from source
        copy_source_to_tesseract()
        
        # Update document_awareness.py with the correct path
        update_code_with_path()
        
        # Test if Tesseract is working
        test_tesseract()
        
        print("\n=== Installation Complete ===")
        print("You can now use the Word Interface Explorer with OCR capabilities!")
        return 0
    else:
        print("\n[WARNING] Tesseract installation process completed but unable to detect Tesseract.")
        print("After installing Tesseract manually, please run this script again.")
        return 1

if __name__ == "__main__":
    import ctypes
    # Skip admin check in non-interactive mode
    try:
        sys.exit(main())
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        sys.exit(1)
