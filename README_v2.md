# Account Mapping Tool v2.5 - Multi-Project Edition

A professional Python Tkinter desktop application for mapping account descriptions between multiple P&L projects in Excel workbooks with intelligent pattern recognition, manual editing capabilities, and automated monthly statement generation.

> ⚠️ **Important: Python 3.13 Required**  
> This application requires Python 3.13 specifically for proper mouse event handling in Tkinter.  
> Other Python versions may cause mouse clicks to not respond consistently.  
> **Recommended command**: `/usr/local/bin/python3.13 run_app_v2.py`

## 📋 Table of Contents
- [Overview](#overview)
- [Core Features](#core-features)
- [Application Architecture](#application-architecture)
- [Workflow & Logic](#workflow--logic)
- [Data Management](#data-management)
- [Installation & Setup](#installation--setup)
- [User Guide](#user-guide)
- [Technical Details](#technical-details)
- [Troubleshooting](#troubleshooting)
- [Version History](#version-history)

## Overview

The Account Mapping Tool v2 is a sophisticated desktop application designed for financial professionals working with multi-property real estate P&L statements. It automates the complex process of mapping account descriptions between source P&L files and rolling P&L templates, handling multiple projects simultaneously with intelligent account matching and manual override capabilities.

### Key Business Problem Solved
- **Manual Mapping Elimination**: Automates the tedious process of matching hundreds of account descriptions
- **Multi-Project Management**: Handles multiple properties/projects in a single workflow
- **Data Integrity**: Preserves Excel formulas and prevents data loss
- **Time Savings**: Reduces hours of manual work to minutes

## Core Features

### 🏢 Multi-Project Support
- **Automatic Project Detection**: Scans Excel workbooks and extracts project names from cell A1
- **Project Isolation**: Each project maintains separate settings, mappings, and data
- **Quick Switching**: Dropdown menu for instant project navigation
- **Bulk Operations**: Export all projects or process individually

### 🤖 Intelligent Mapping Engine
- **Fuzzy String Matching**: Uses advanced algorithms to match similar account descriptions
- **Confidence Scoring**: High/Medium/Low confidence ratings based on similarity
- **Pattern Recognition**: Detects account numbers, categories, and common variations
- **Learning Capability**: Preserves manual edits and applies them consistently

### 📊 Excel Integration
- **Formula Preservation**: Maintains Excel formulas when updating cells
- **SUM Formula Generation**: Creates intelligent formulas that preserve existing values
  - Empty cells: Writes new value directly
  - Cells with values: Creates `=existing_value + new_value`
  - Cells with formulas: Creates `=(existing_formula) + new_value`
- **Multi-Sheet Support**: Handles complex workbooks with multiple sheets
- **Subtotal Handling**: Displays subtotal accounts with amounts but no mapping
- **Password Protection**: Supports encrypted Excel files
- **Format Flexibility**: Works with .xlsx and .xls formats

### 💾 Data Persistence
- **Automatic Saving**: All changes saved immediately
- **Project Settings**: JSON-based configuration storage
- **Backup System**: Automatic timestamped backups with retention policy
- **Session Recovery**: Restores last state on restart

### 🎨 Advanced UI Features
- **Pop-out Windows**: Detachable mapping and preview windows with intelligent centering
  - All windows open centered on main application screen
  - Consistent sizing (1400x850) for Step 2 and Step 3 windows
  - No initial repositioning flash - windows appear directly in place
  - Windows stay on same monitor as main application
- **Zoom Controls**: Adjustable interface scaling (Cmd+/- or Ctrl+/-)
- **Real-time Filtering**: Search and filter mappings instantly
- **Bulk Editing**: Select multiple accounts for batch operations
- **Progress Indicators**: Visual feedback for long operations
- **Enhanced Scrolling**: Mouse wheel works anywhere in the application window

## Application Architecture

### File Structure
```
account-mapping-app/
├── Core Application Files
│   ├── main_v2.py              # Main application (5800+ lines)
│   ├── run_app_v2.py           # Application launcher with error handling
│   └── project_manager.py      # Project management and persistence
│
├── Configuration Files
│   ├── project_settings.json   # Multi-project configurations & mappings
│   ├── range_settings.json     # Default Excel range specifications
│   └── range_memory.json       # Per-project range persistence
│
├── Documentation
│   ├── README_v2.md            # This comprehensive documentation
│   ├── CLAUDE.md               # AI assistant guidance
│   └── CLAUDE_CODE_HISTORY.md  # Development session history
│
├── Testing & Samples
│   ├── test_performance.py     # Performance profiling utilities
│   ├── sample_source_pl.xlsx   # Sample multi-project source file
│   └── sample_rolling_pl.xlsx  # Sample rolling P&L template
│
├── Dependencies
│   ├── requirements.txt        # Python package dependencies
│   └── .gitignore             # Git ignore rules
│
└── Generated Directories
    ├── backups/                # Automatic project backups
    └── restored_files/         # Files recovered from backups
```

### Component Architecture

#### 1. **Main Application (`main_v2.py`)**
- **Class**: `MultiProjectAccountMappingApp`
- **Lines**: 5800+ lines of production code
- **Responsibilities**:
  - GUI management and event handling
  - Workflow orchestration (4-step process)
  - Excel file processing
  - Mapping generation and management
  - UI state management

#### 2. **Project Manager (`project_manager.py`)**
- **Classes**: `Project`, `ProjectManager`
- **Responsibilities**:
  - Project data encapsulation
  - Settings persistence (JSON)
  - Project switching logic
  - Backup and restore functionality

#### 3. **Launcher (`run_app_v2.py`)**
- **Purpose**: Safe application startup
- **Features**:
  - Dependency checking
  - Error handling
  - Graceful failure recovery

## Workflow & Logic

### 4-Step Workflow Process

#### Step 1: Upload Source and Rolling P&L Files
1. **Source P&L Upload**:
   - User selects multi-project Excel workbook
   - System scans all sheets
   - Extracts project names from cell A1 of each sheet
   - Creates project list automatically

2. **Project Selection**:
   - User selects active project from dropdown
   - UI updates to show project-specific data
   - Previous project data is preserved

3. **Rolling P&L Configuration**:
   - User uploads rolling P&L template
   - Selects appropriate sheet for current project
   - System validates sheet existence

#### Step 2: Map Account Descriptions

##### Data Flow in Step 2:
```
1. EXTRACTION PHASE
   Source Excel → extract_range_data() → Source Accounts List
   Rolling Excel → extract_range_data() → Rolling Accounts List
   
2. MAPPING GENERATION
   Source + Rolling Lists → create_intelligent_mappings() → Mappings Dict
   └── Fuzzy String Matching (difflib.SequenceMatcher)
   └── Confidence Scoring (High >80%, Medium >60%, Low >40%)
   └── Subtotals get empty mappings (no rolling account)
   
3. AMOUNT EXTRACTION
   Source Excel → extract_monthly_amounts() → Monthly Amounts Dict
   └── Finds target month column (searches for date patterns)
   └── Extracts amounts for ALL accounts (including subtotals)
   └── Caches results in source_amounts_cache
   
4. UI PRESENTATION
   populate_mapping_tree()
   ├── Regular Accounts: Checkbox + Description + Amount + Mapping + Confidence
   ├── Subtotals: No Checkbox + Bold Description + Amount + No Mapping
   └── Preserves Excel source order exactly
```

##### Key Components:
1. **Range Configuration**:
   - User specifies source account range (e.g., A8:F200)
   - User specifies rolling account range (e.g., A1:A100)
   - Live preview shows extracted data
   - Ranges saved to `project_settings.json` and `range_memory.json`

2. **Automatic Mapping Generation**:
   - System extracts accounts from both ranges (including subtotals)
   - Applies fuzzy matching algorithm (difflib)
   - Generates confidence scores based on similarity
   - Creates initial mappings (subtotals get empty mappings)

3. **Subtotal Handling** (NEW in v2.2):
   - Subtotals identified by "total" keyword in description
   - Extracted with their amounts but no mapping capability
   - Display in bold without checkboxes
   - Maintain exact Excel source order

4. **Manual Editing**:
   - Double-click to edit individual mappings
   - Bulk edit for multiple selections
   - Search and filter capabilities
   - All edits marked as `"user_edited": true`
   - Saved immediately to `project_settings.json`

#### Step 3: Review Mapped Accounts & Generate Statement

##### Data Flow in Step 3:
```
1. MONTHLY DATA EXTRACTION
   Source Excel + Mappings → extract_monthly_amounts()
   └── Reads from target month column
   └── Uses cached data if available
   └── Returns: {"Account": Amount} dict
   
2. AGGREGATION PHASE
   Monthly Data + Mappings → aggregate_by_mappings()
   └── Groups amounts by rolling account
   └── Skips accounts with empty mappings (subtotals)
   └── Returns: {"Rolling Account": Total Amount} dict
   
3. PREVIEW PREPARATION
   Rolling Excel + Aggregated Data → prepare_preview_data()
   ├── Reads previous month column
   ├── Matches with aggregated current month
   └── Returns: [(Account, Previous, Current)] list
   
4. DATA PERSISTENCE
   All data saved to project_settings.json:
   ├── monthly_data (raw amounts)
   ├── aggregated_data (grouped totals)
   ├── preview_data (for display)
   └── target_month (detected column header)
```

##### Key Components:
1. **Preview Generation**:
   - Extracts amounts from source P&L target month column
   - Aggregates by mapped rolling account categories
   - Shows 3-column preview (Account, Previous Month, Current Month)
   - Detects target month automatically from column headers

2. **Data Validation**:
   - Verifies all mappings are complete
   - Checks for missing amounts
   - Validates target cells in rolling P&L
   - Subtotals excluded from aggregation (empty mappings)

3. **Statement Generation** (Enhanced in v2.3):
   - Creates SUM formulas instead of overwriting values
   - Preserves both existing values and formulas
   - For empty cells: Writes new value directly
   - For cells with values: Creates formula `=existing_value + new_value`
   - For cells with formulas: Creates formula `=(existing_formula) + new_value`
   - Handles negative numbers with proper parentheses
   - Respects Excel formula length limits (8,192 characters)

#### Step 4: Save Final Rolling P&L with Actual Data
1. **Export Options**:
   - Export current project
   - Export all projects with data
   - Custom filename with timestamp

2. **Finalization**:
   - Updates all formulas
   - Saves to specified location
   - Creates backup automatically

### Automatic Workflow Detection
The application intelligently detects when all requirements are met and can automatically:
- Generate mappings when ranges are specified
- Create monthly statements when mappings exist
- Skip redundant steps when data is already processed

### New Account Detection
- **On Startup**: Checks for new accounts not in existing mappings
- **Automatic Addition**: Adds new accounts with intelligent matching
- **Order Preservation**: Maintains source file account order
- **Manual Override Protection**: Never overwrites user edits

## Data Management & Architecture

### Data Storage Hierarchy

The application uses a sophisticated multi-tier data management system:

```
┌─────────────────────────────────────────────────────┐
│                  USER INTERFACE                      │
│            (Tkinter GUI - main_v2.py)               │
└─────────────────────────────────────────────────────┘
                          ↕
┌─────────────────────────────────────────────────────┐
│               APPLICATION MEMORY                     │
│  • current_project (active Project object)          │
│  • projects dict (all Project objects)              │
│  • Cache layers (DataFrames, amounts, accounts)     │
└─────────────────────────────────────────────────────┘
                          ↕
┌─────────────────────────────────────────────────────┐
│              PROJECT MANAGER                         │
│         (project_manager.py)                        │
│  • Handles project switching                        │
│  • Manages persistence                              │
│  • Coordinates saves/loads                          │
└─────────────────────────────────────────────────────┘
                          ↕
┌─────────────────────────────────────────────────────┐
│             PERSISTENT STORAGE                       │
│  • project_settings.json (main data)                │
│  • range_memory.json (range persistence)            │
│  • range_settings.json (default ranges)             │
│  • Excel files (source data)                        │
│                                                      │
│  Storage Location:                                   │
│  • Development: ./account-mapping-app/              │
│  • Standalone: ~/Documents/AccountMappingTool/      │
└─────────────────────────────────────────────────────┘
```

### Project Settings Structure
```json
{
  "source_workbook_path": "/path/to/source.xlsx",
  "rolling_workbook_path": "/path/to/rolling.xlsx",
  "current_project": "Columbia Villas",
  "projects": {
    "Columbia Villas": {
      "name": "Columbia Villas",
      "source_sheet": "COVI",
      "rolling_sheet": "Columbia",
      "source_range": "A8:F200",
      "rolling_range": "A1:A100",
      "source_file_path": "/path/to/source.xlsx",
      "mapping_file_path": "/path/to/mapping.json",
      "mappings": {
        "8000 Rental Income": {
          "rolling_account": "Rental Revenue",
          "confidence": "High",
          "similarity": 85.5,
          "user_edited": false
        },
        "8540 Other Exterior Replacement": {
          "rolling_account": "Other Income",
          "confidence": "Manual",
          "similarity": 100.0,
          "user_edited": true
        },
        "Total Rental Income": {
          "rolling_account": "",
          "confidence": "None",
          "similarity": 0,
          "user_edited": false
        }
      },
      "monthly_data": {
        "8000 Rental Income": 50000.00,
        "8540 Other Exterior Replacement": 5020.00,
        "Total Rental Income": 55020.00
      },
      "aggregated_data": {
        "Rental Revenue": 50000.00,
        "Other Income": 5020.00
      },
      "preview_data": [],
      "target_month": "Jun-25",
      "step4_completed": false,
      "workflow_state": {
        "step1_complete": true,
        "step2_complete": true,
        "step3_complete": false,
        "step4_complete": false,
        "has_generated_mappings": true,
        "has_generated_monthly": false
      },
      "ui_state": {
        "filter_value": "",
        "sort_value": "",
        "zoom_level": 1.0,
        "checkbox_states": {},
        "last_active_step": 2
      }
    }
  }
}
```

### Data Loading Sequence (On Startup)

1. **Application Launch** (`run_app_v2.py`)
   ```python
   ProjectManager.__init__()
   ├── load_settings() → Reads project_settings.json
   ├── load_range_memory() → Reads range_memory.json
   └── Creates Project objects from saved data
   ```

2. **Project Restoration** 
   ```python
   Project.from_dict()
   ├── Restores mappings (OrderedDict)
   ├── Restores monthly_data (amounts)
   ├── Restores aggregated_data
   ├── Restores workflow_state
   └── Restores ui_state
   ```

3. **New Account Detection**
   ```python
   check_and_add_new_accounts()
   ├── Extracts current source accounts
   ├── Compares with existing mappings
   ├── Adds new accounts with fuzzy matching
   └── Preserves source file order
   ```

### Cache Management System

The application implements a multi-level caching strategy for performance:

#### 1. **DataFrame Cache** (`_excel_cache`)
- **Purpose**: Avoid re-reading Excel files
- **Key Format**: `"{file_path}:{sheet_name}"`
- **Lifetime**: Application session
- **Size**: Unlimited during session

#### 2. **Source Amounts Cache** (`source_amounts_cache`)
- **Purpose**: Store extracted monthly amounts
- **Key Format**: `"{project_name}:{file_path}:{range}"`
- **Lifetime**: Cleared when ranges change
- **Usage**: Step 2 display, Step 3 aggregation

#### 3. **Rolling Accounts Cache** (`rolling_accounts_cache`)
- **Purpose**: Quick access to target accounts
- **Key Format**: `"{project_name}:{rolling_range}"`
- **Lifetime**: Session or until range changes
- **Usage**: Dropdown populations, validations

#### 4. **Target Month Cache** (`target_month_cache`)
- **Purpose**: Remember which column has current month data
- **Key Format**: `"{project_name}:{file_path}"`
- **Lifetime**: Session
- **Usage**: Speeds up repeated extractions

### Mapping Priority System
1. **User Manual Edits** (Highest Priority)
   - Always preserved across sessions
   - Never overwritten by automatic generation
   - Marked with `"user_edited": true`
   - Confidence set to "Manual"

2. **Project Settings File** (`project_settings.json`)
   - Primary source of truth
   - Loaded on startup
   - Updated after every change
   - Contains all project data

3. **External Mapping Files** (Lowest Priority)
   - Only loaded if no mappings exist
   - Used for initial import
   - Not loaded if project has existing mappings
   - Prevented from overwriting (fix in v2.2)

## Installation & Setup

### System Requirements
- **Python**: 3.13 specifically (required for consistent mouse event handling)
  - ⚠️ **Important**: Other Python versions may cause mouse clicks to not respond consistently
  - Tkinter event handling works best with Python 3.13
- **Operating System**: Windows, macOS, or Linux
- **Memory**: 4GB RAM minimum (8GB recommended)
- **Display**: 1280x720 minimum resolution

### Installation Steps

#### First, verify Python 3.13 is installed:
```bash
# Check if Python 3.13 is available
python3.13 --version
# or on macOS with Homebrew:
/usr/local/bin/python3.13 --version
```

#### Option 1: Direct Execution (Recommended)
```bash
# 1. Clone the repository
git clone https://github.com/ydinai001/account-mapping-app.git
cd account-mapping-app

# 2. Install dependencies directly with Python 3.13
/usr/local/bin/python3.13 -m pip install -r requirements.txt

# 3. Run the application with full Python 3.13 path
/usr/local/bin/python3.13 run_app_v2.py
```

#### Option 2: Using Virtual Environment
```bash
# 1. Clone the repository
git clone https://github.com/ydinai001/account-mapping-app.git
cd account-mapping-app

# 2. Create virtual environment with Python 3.13
/usr/local/bin/python3.13 -m venv .venv

# 3. Activate virtual environment
source .venv/bin/activate  # macOS/Linux
# or
.venv\Scripts\activate  # Windows

# 4. Install dependencies
pip install -r requirements.txt

# 5. Run the application
python run_app_v2.py
```

### Dependencies
- **pandas**: Excel file processing and data manipulation
- **openpyxl**: Excel file reading/writing with formula support
- **tkinter**: GUI framework (included with Python)
- **msoffcrypto-tool**: Password-protected Excel support (optional)

## User Guide

### Getting Started
1. **Launch Application**: Run `/usr/local/bin/python3.13 run_app_v2.py`
   - For macOS users: Use the full path to ensure Python 3.13 is used
   - For virtual environment: Activate venv first, then run `python run_app_v2.py`
2. **Upload Source P&L**: Click "Browse" and select your multi-project workbook
3. **Scan Projects**: Click "Scan Projects" to detect all available projects
4. **Select Project**: Choose a project from the dropdown menu
5. **Upload Rolling P&L**: Select your rolling P&L template file
6. **Configure Ranges**: Set source and rolling account ranges
7. **Generate Mappings**: Click "Generate Automatic Mappings"
8. **Review and Edit**: Double-click any mapping to modify
9. **Generate Statement**: Click "Generate Monthly Statement"
10. **Export Results**: Save the final rolling P&L with actual data

### Keyboard Shortcuts
- **Cmd/Ctrl + Plus**: Zoom in
- **Cmd/Ctrl + Minus**: Zoom out
- **Cmd/Ctrl + 0**: Reset zoom
- **Double-click**: Edit mapping
- **Right-click**: Context menu

### Tips for Best Results
1. **Consistent Naming**: Use consistent account names across projects
2. **Range Selection**: Include all accounts in your range specifications
3. **Manual Review**: Always review automatic mappings before generating statements
4. **Regular Backups**: The app creates automatic backups, but keep your own too
5. **Save Mappings**: Export mappings for reuse in future periods

## Technical Details

### Performance Optimizations
- **Lazy Loading**: Data loaded only when needed
- **Caching Strategy**: Multi-level caching for repeated operations
  - Rolling preview uses separate cache without formula evaluation
  - DataFrame cache prevents redundant Excel file reads
- **Batch Processing**: Bulk operations for better performance
- **Memory Management**: Efficient data structure usage
- **Window Geometry**: Optimized positioning calculations with `update_idletasks()`
- **Preview Performance**: Rolling range preview loads without expensive month detection

### Error Handling
- **File Access Errors**: Graceful handling of locked/missing files
- **Data Validation**: Comprehensive input validation
- **Recovery Options**: Automatic backup restoration
- **User Feedback**: Clear error messages with solutions

### Security Features
- **No Network Access**: Completely offline operation
- **Local Storage Only**: All data stored locally
- **Password Support**: Handles encrypted Excel files
- **Data Isolation**: Projects cannot access each other's data

## Troubleshooting

### Common Issues and Solutions

#### "No projects found in workbook"
- **Cause**: Cell A1 doesn't contain project names
- **Solution**: Ensure each sheet has project name in cell A1

#### "Account 8540 not showing after restart"
- **Cause**: Old mapping file overwriting updates
- **Solution**: Fixed in latest version - mappings now persist correctly

#### "Cannot write to rolling P&L"
- **Cause**: File is open in Excel
- **Solution**: Close the file in Excel before generating statement

#### "Mappings not saving"
- **Cause**: Permission issues or file conflicts
- **Solution**: Check file permissions and ensure write access

#### "Performance is slow"
- **Cause**: Large Excel files or many projects
- **Solution**: 
  - Reduce range sizes to necessary cells only
  - Process projects individually
  - Close other applications

### Debug Mode
For troubleshooting, uncomment debug lines in:
- `project_manager.py`: Lines 527-528, 535-543 (save verification)
- `main_v2.py`: Various `pass` statements replaced with print statements

## Version History

### v2.5 (August 2025) - Current
- **Feature**: Browse for backup functionality - load backups from any location
- **Enhancement**: Automatic writable directory detection for standalone app
- **Fix**: Resolved read-only file system error when loading backups
- **Feature**: App data now stored in ~/Documents/AccountMappingTool/ for standalone version
- **Enhancement**: Backup restoration works from USB drives, network locations, or any folder
- **Improvement**: Better error handling and validation for backup files

### v2.4 (August 2025)
- **Enhancement**: Universal window centering system for all pop-out windows
- **Fix**: Pop-out windows now open on same screen as main application
- **Enhancement**: Consistent window sizing (1400x850) for Step 2 and Step 3
- **Fix**: Eliminated initial corner flash for all pop-out windows
- **Performance**: Optimized rolling range preview loading speed
- **UX**: Improved scrolling to work in gaps between subwindows
- **Feature**: Added `center_window_on_parent()` method for unified dialog positioning
- **Enhancement**: All pop-out windows use transient relationship with main window

### v2.3 (August 2025)
- **Feature**: Subtotals now included in Step 2 with amounts
- **Enhancement**: Extract and display all accounts including totals
- **Fix**: Modified `extract_range_data()` to include subtotal accounts
- **Fix**: Updated `extract_monthly_amounts()` to extract amounts for subtotals
- **Enhancement**: Subtotals display in bold without checkboxes or mapping options
- **Feature**: Step 3 export now creates SUM formulas instead of overwriting
- **Enhancement**: Preserves existing values and formulas when exporting
- **Feature**: Intelligent formula creation with negative number handling
- **Documentation**: Comprehensive README update with data flow architecture

### v2.2 (August 2025)
- **Fix**: Prevented old mapping files from overwriting updated mappings
- **Feature**: Automatic detection and addition of new accounts
- **Fix**: Object reference consistency between current_project and manager
- **Enhancement**: Improved save verification and debugging
- **Feature**: Main window scrollbar for easier navigation
- **Fix**: Context-aware scrolling to prevent subwindow scroll conflicts

### v2.1 (July 2025)
- **Feature**: Pop-out windows for Steps 2 and 3
- **Enhancement**: Advanced filtering and search capabilities
- **Fix**: Source amount display during filtering
- **Feature**: Bulk editing with checkbox selection
- **Enhancement**: Bold category headings and improved spacing

### v2.0 (July 2025)
- **Major**: Multi-project support architecture
- **Feature**: Automatic project detection from workbooks
- **Enhancement**: Project isolation and switching
- **Feature**: Persistent project settings

### v1.x Series
- Initial single-project version
- Basic mapping functionality
- Manual editing capabilities

## Development Notes

### Code Organization
- **Modular Design**: Separation of concerns between UI and logic
- **Event-Driven**: Tkinter event handling for responsive UI
- **State Management**: Centralized project state management
- **Error Boundaries**: Try-catch blocks at all entry points

### Future Enhancements (Planned)
- [ ] Cloud storage integration
- [ ] Mapping templates library
- [ ] Advanced reporting features
- [ ] API for external integrations
- [ ] Multi-user collaboration

## Support & Contributing

### Getting Help
- Check this README for comprehensive documentation
- Review CLAUDE.md for development guidance
- Check existing issues on GitHub

### Contributing
1. Work on `testing` branch only
2. Follow existing code style
3. Add tests for new features
4. Update documentation
5. Submit pull request to `testing` branch

### License
Proprietary - All rights reserved

---

**Last Updated**: August 13, 2025  
**Version**: 2.5  
**Maintainer**: Account Mapping Development Team