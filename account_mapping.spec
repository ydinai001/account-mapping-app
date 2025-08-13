# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Account Mapping Tool v2.4
Creates a standalone macOS application bundle
"""

import os
import sys
from pathlib import Path

block_cipher = None

# Get the absolute path to the project directory
project_dir = os.path.abspath('.')

# Define the main script
main_script = 'run_app_v2.py'

# Collect all Python files
python_files = [
    'run_app_v2.py',
    'main_v2.py', 
    'project_manager.py',
]

# Data files to include (samples, configs, etc.)
data_files = [
    ('range_settings.json', '.'),
    ('README_v2.md', '.'),
    ('requirements.txt', '.'),
]

# Add sample files if they exist
sample_files = [
    'sample_source_pl.xlsx',
    'sample_rolling_pl.xlsx',
]

for sample in sample_files:
    if os.path.exists(sample):
        data_files.append((sample, 'samples'))

# Hidden imports that PyInstaller might miss
hidden_imports = [
    'pandas',
    'pandas._libs',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.skiplist',
    'numpy',
    'numpy.random',
    'numpy.random._pickle',
    'numpy.random._common',
    'numpy.random._bounded_integers',
    'numpy.random._mt19937',
    'numpy.random._pcg64',
    'numpy.random._philox',
    'numpy.random._sfc64',
    'numpy.random._generator',
    'numpy.random.bit_generator',
    'numpy.random.mtrand',
    'numpy.random.numpy',
    'openpyxl',
    'openpyxl.cell._writer',
    'xlrd',
    'msoffcrypto',
    'msoffcrypto.method.ecma376_agile',
    'tkinter',
    'tkinter.ttk',
    'tkinter.messagebox',
    'tkinter.filedialog',
    'json',
    'datetime',
    'collections',
    'pathlib',
    'warnings',
    'subprocess',
    'platform',
    'traceback',
    'pickle',
    '_pickle',
    'pkg_resources.extern',
]

a = Analysis(
    [main_script],
    pathex=[project_dir],
    binaries=[],
    datas=data_files,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'test',
        'unittest',
        'pygame',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AccountMappingTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AccountMappingTool',
)

app = BUNDLE(
    coll,
    name='Account Mapping Tool.app',
    icon=None,  # We'll add an icon later if needed
    bundle_identifier='com.accountmapping.tool',
    version='2.5.0',
    info_plist={
        'NSPrincipalClass': 'NSApplication',
        'NSAppleScriptEnabled': False,
        'CFBundleName': 'Account Mapping Tool',
        'CFBundleDisplayName': 'Account Mapping Tool',
        'CFBundleGetInfoString': 'Account Mapping Tool v2.5',
        'CFBundleIdentifier': 'com.accountmapping.tool',
        'CFBundleVersion': '2.5.0',
        'CFBundleShortVersionString': '2.5.0',
        'NSHumanReadableCopyright': 'Copyright © 2025. All rights reserved.',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.13.0',
        'NSRequiresAquaSystemAppearance': False,
        'LSApplicationCategoryType': 'public.app-category.business',
    },
)