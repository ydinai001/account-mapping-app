# Account Mapping Tool - Installation Guide

## 📦 Installation Package Contents

You should have received the following file:
- `AccountMappingTool-v2.5.dmg` - The installation disk image

## 🚀 Installation Steps

### Step 1: Open the DMG File
1. Double-click `AccountMappingTool-v2.5.dmg`
2. A new window will open showing the app and Applications folder

### Step 2: Install the Application
1. Drag "Account Mapping Tool" to the Applications folder (shortcut provided)
2. Wait for the copy to complete
3. Eject the DMG by clicking the eject button in Finder

### Step 3: First Launch
Due to macOS security (Gatekeeper), the first launch requires special steps:

1. Open your Applications folder
2. Find "Account Mapping Tool"
3. **Right-click** on the app icon
4. Select "Open" from the context menu
5. Click "Open" in the security dialog

After the first launch, you can open the app normally by double-clicking.

## 💻 System Requirements

- **macOS**: 10.13 (High Sierra) or later
- **Memory**: 4GB RAM minimum
- **Storage**: 100MB free space
- **Python**: Not required (bundled in app)

## 🎯 Quick Start

After installation:
1. Launch the app from Applications
2. Click "Browse" to load your Source P&L Excel file
3. Click "Scan Projects" to detect all projects
4. Select a project from the dropdown
5. Follow the 4-step workflow

## ⚠️ Troubleshooting

### "App can't be opened because it is from an unidentified developer"
- Right-click the app and select "Open" instead of double-clicking
- This is a one-time security check

### App doesn't launch
- Ensure you're running macOS 10.13 or later
- Try restarting your Mac
- Re-download and reinstall the DMG

### Excel files not loading
- Ensure Excel files are not password-protected
- Check that files are in .xlsx or .xls format
- Close Excel before processing files

## 📁 Application Data Location

The app stores its data in:
- **Settings & Data**: `~/Documents/AccountMappingTool/`
  - `project_settings.json` - All project configurations and mappings
  - `range_settings.json` - Default Excel range specifications
  - `range_memory.json` - Per-project range persistence
  - `backups/` - Local backup storage
  - `restored_files/` - Excel files restored from backups

### Loading Existing Backups
The app now includes a "Browse for Backup" feature:
1. Click "Load Backup" from the main window
2. If no local backups exist, click "Yes" to browse
3. Navigate to your backup location (e.g., external drive, network share)
4. Select the backup folder to restore

**Note**: Backups from the development version can be loaded into the standalone app using the browse feature.

## 🗑️ Uninstallation

To remove the app:
1. Quit the Account Mapping Tool if running
2. Open Applications folder
3. Drag "Account Mapping Tool" to Trash
4. Empty Trash

To remove all app data:
```bash
rm -rf ~/Documents/AccountMappingTool/
```

## 📞 Support

For issues or questions:
- Refer to the README_v2.md documentation
- Check the built-in help within the application
- Contact your system administrator

## 📄 License

This software is proprietary and licensed for authorized use only.

---

**Version**: 2.5  
**Release Date**: August 2025  
**Compatibility**: macOS 10.13+