#!/usr/bin/env python3
"""
Build script for Account Mapping Tool
Creates a standalone macOS application
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def clean_build_dirs():
    """Clean previous build artifacts"""
    print("🧹 Cleaning previous build artifacts...")
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"  ✓ Removed {dir_name}/")

def check_requirements():
    """Check if all required files exist"""
    print("\n📋 Checking requirements...")
    required_files = [
        'run_app_v2.py',
        'main_v2.py',
        'project_manager.py',
        'requirements.txt',
        'account_mapping.spec'
    ]
    
    for file in required_files:
        if not os.path.exists(file):
            print(f"  ✗ Missing required file: {file}")
            return False
        print(f"  ✓ Found {file}")
    return True

def build_application():
    """Build the application using PyInstaller"""
    print("\n🔨 Building application with PyInstaller...")
    print("  This may take a few minutes...")
    
    # Use Python 3.13 specifically
    python_exe = '/usr/local/bin/python3.13'
    
    # Build command
    cmd = [
        python_exe, '-m', 'PyInstaller',
        '--clean',
        '--noconfirm',
        'account_mapping.spec'
    ]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  ✗ Build failed!")
            print(f"  Error: {result.stderr}")
            return False
        print("  ✓ Build completed successfully!")
        return True
    except Exception as e:
        print(f"  ✗ Build error: {e}")
        return False

def create_dmg():
    """Create a DMG installer for macOS"""
    print("\n📦 Creating DMG installer...")
    
    app_path = "dist/Account Mapping Tool.app"
    dmg_name = "AccountMappingTool-v2.5.dmg"
    
    if not os.path.exists(app_path):
        print(f"  ✗ App bundle not found at {app_path}")
        return False
    
    # Remove old DMG if it exists
    if os.path.exists(f"dist/{dmg_name}"):
        os.remove(f"dist/{dmg_name}")
    
    # Create a temporary folder for DMG contents
    dmg_folder = "dist/dmg_contents"
    if os.path.exists(dmg_folder):
        shutil.rmtree(dmg_folder)
    os.makedirs(dmg_folder)
    
    # Copy app to DMG folder
    shutil.copytree(app_path, f"{dmg_folder}/Account Mapping Tool.app")
    
    # Create a symlink to Applications
    os.symlink('/Applications', f"{dmg_folder}/Applications")
    
    # Create README file for DMG
    readme_content = """Account Mapping Tool v2.5
========================

Installation:
1. Drag "Account Mapping Tool" to the Applications folder
2. Double-click to run from Applications

First Run:
- You may need to right-click and select "Open" the first time
- This is due to macOS security (Gatekeeper)

Requirements:
- macOS 10.13 or later
- No Python installation required

For help and documentation, see README_v2.md
"""
    
    with open(f"{dmg_folder}/README.txt", 'w') as f:
        f.write(readme_content)
    
    # Create DMG using hdiutil
    cmd = [
        'hdiutil', 'create',
        '-volname', 'Account Mapping Tool',
        '-srcfolder', dmg_folder,
        '-ov',
        '-format', 'UDZO',
        f"dist/{dmg_name}"
    ]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  ✗ DMG creation failed!")
            print(f"  Error: {result.stderr}")
            return False
        print(f"  ✓ DMG created: dist/{dmg_name}")
        
        # Clean up temporary folder
        shutil.rmtree(dmg_folder)
        return True
    except Exception as e:
        print(f"  ✗ DMG creation error: {e}")
        return False

def print_instructions():
    """Print post-build instructions"""
    print("\n" + "="*50)
    print("✅ BUILD COMPLETE!")
    print("="*50)
    print("\n📁 Output files:")
    print("  • Application: dist/Account Mapping Tool.app")
    print("  • Installer: dist/AccountMappingTool-v2.5.dmg")
    print("\n📝 Distribution instructions:")
    print("  1. Share the DMG file with users")
    print("  2. Users drag the app to Applications folder")
    print("  3. First run: Right-click → Open (bypasses Gatekeeper)")
    print("\n⚠️  Note: The app is not code-signed.")
    print("  Users may see a security warning on first launch.")
    print("  To sign the app, you need an Apple Developer certificate.")

def main():
    """Main build process"""
    print("🚀 Account Mapping Tool - Build Script")
    print("="*50)
    
    # Check we're in the right directory
    if not os.path.exists('main_v2.py'):
        print("❌ Error: Please run this script from the account-mapping-app directory")
        sys.exit(1)
    
    # Clean previous builds
    clean_build_dirs()
    
    # Check requirements
    if not check_requirements():
        print("\n❌ Build aborted: Missing requirements")
        sys.exit(1)
    
    # Build application
    if not build_application():
        print("\n❌ Build failed")
        sys.exit(1)
    
    # Create DMG
    if not create_dmg():
        print("\n⚠️  Warning: DMG creation failed, but app bundle is ready")
    
    # Print instructions
    print_instructions()

if __name__ == "__main__":
    main()