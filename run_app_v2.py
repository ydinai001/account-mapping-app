#!/usr/bin/env python3
"""
Launcher script for the Multi-Project Account Mapping Application v2
Provides better error handling and user feedback
"""

import sys
import os
import traceback
import tkinter as tk
from tkinter import messagebox
import subprocess
import platform
import warnings

# Suppress openpyxl header/footer parsing warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def activate_macos_app():
    """Force macOS to activate the tkinter app for proper mouse functionality"""
    if platform.system() == "Darwin":  # macOS only
        try:
            script = '''
            tell application "System Events"
                set frontmost of first process whose frontmost is true to false
                set frontmost of first process whose name contains "Python" to true
            end tell
            '''
            subprocess.run(['osascript', '-e', script], check=False, capture_output=True)
        except Exception:
            pass

def check_dependencies():
    """Check if all required dependencies are available"""
    try:
        import pandas as pd
        import openpyxl
        return True, "All dependencies available"
    except ImportError as e:
        return False, f"Missing dependency: {e}"

def main():
    """Main launcher function"""
    
    try:
        # Check dependencies first
        deps_ok, deps_msg = check_dependencies()
        if not deps_ok:
            error_msg = f"Dependency Error: {deps_msg}\n\nPlease install required packages:\npip install pandas openpyxl"
            
            try:
                root = tk.Tk()
                root.withdraw()  # Hide main window
                messagebox.showerror("Dependency Error", error_msg)
                root.destroy()
            except:
                pass
            
            return 1
        
        # Import and create the application
        from main_v2 import MultiProjectAccountMappingApp
        
        root = tk.Tk()
        
        # Set application icon if available
        try:
            # You can add an icon file here if desired
            pass
        except:
            pass
        
        # Create and run the application
        _ = MultiProjectAccountMappingApp(root)
        
        # Python 3.13 works without additional mouse fixes
        
        # Start the main loop
        root.mainloop()
        
        return 0
        
    except Exception as e:
        error_msg = f"Error starting application: {str(e)}"
        print(f"\nERROR: {error_msg}")
        print("\nFull traceback:")
        traceback.print_exc()
        
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Application Error", f"{error_msg}\n\nSee console for details.")
            root.destroy()
        except:
            pass
        
        return 1

if __name__ == "__main__":
    sys.exit(main())