# Account Mapping Tool v2 - Multi-Project Edition

A professional desktop application for mapping account descriptions between multiple P&L projects in Excel workbooks with intelligent pattern recognition, manual editing capabilities, and monthly statement generation.

## 🆕 **What's New in Version 2**

### **Multi-Project Support**
- **Excel Workbooks**: Upload workbooks containing multiple project sheets
- **Project Isolation**: Each project maintains separate settings and data
- **Project Switching**: Easy navigation between projects via dropdown menu
- **Automatic Detection**: Scans workbooks and extracts project names from cell A1

### **Enhanced User Interface**
- **Project Header**: Shows current project name in window title and header
- **Context-Aware UI**: Interface adapts based on selected project
- **Smart Workflow**: Steps activate progressively as requirements are met
- **Project Menu**: Top-right dropdown for quick project switching

### **Advanced File Handling**
- **Encrypted Files**: Supports password-protected Excel workbooks
- **Sheet Mapping**: Link projects to specific worksheets in rolling P&L
- **Persistent Settings**: Project-specific configurations saved automatically
- **Multiple Formats**: Supports both .xlsx and .xls files

## 🎯 **Multi-Project Workflow**

### **Step 1: Project Discovery & Setup**
1. **Upload Source P&L Workbook**
   - Click "Browse..." to select your multi-project Source P&L workbook
   - Click "Scan Projects" to automatically detect all projects
   - System reads project names from cell A1 of each sheet

2. **Project Selection**
   - Select a project from the dropdown menu in the top-right
   - Only selected project data is shown in the interface
   - All other steps become available after project selection

3. **Rolling P&L Configuration**
   - Upload Rolling P&L workbook (after project selection)
   - Select the appropriate sheet for the current project
   - Each project can map to different sheets in the rolling workbook

### **Step 2: Project-Specific Mapping**
- **Range Configuration**: Set account description ranges per project
- **Mapping Generation**: Create intelligent mappings for selected project only
- **Data Isolation**: Each project maintains separate mappings and settings

### **Step 3: Save & Manage**
- **Project-Specific Storage**: Mappings saved separately for each project
- **Reset Option**: Clear mappings for current project only
- **Project Switching**: Switch between projects without losing data

### **Step 4: Monthly Statement Generation**
- **Project Context**: Generate statements for currently selected project
- **Target Workbook**: Write data back to the original rolling P&L file
- **Conflict Resolution**: Handle existing data with user-friendly dialogs

## 📁 **File Structure & Organization**

### **Project Directory Structure**
```
account-mapping-app/
├── main_v2.py              # Main application (5800+ lines)
├── run_app_v2.py           # Application launcher with error handling
├── project_manager.py      # Project management utilities
├── test_performance.py     # Performance testing utilities
├── requirements.txt        # Python dependencies
├── README_v2.md           # This documentation
├── CLAUDE_CODE_HISTORY.md # Development history
├── project_settings.json   # Multi-project configuration
├── range_settings.json     # Excel range configurations
├── range_memory.json       # Range memory persistence
├── sample_source_pl.xlsx   # Sample source P&L file
├── sample_rolling_pl.xlsx  # Sample rolling P&L file
└── backups/               # Automatic project backups (8 snapshots retained)
```

### **Source P&L Workbook Format**
```
Source P&L.xlsx
├── Sheet 1 (e.g., "COVI")
│   ├── A1: "Columbia Villas"     ← Project name
│   ├── A8:F200: Account data     ← Configurable range
│   └── Column I: Amount data
├── Sheet 2 (e.g., "DYVI")
│   ├── A1: "Dyersdale Village"   ← Project name
│   └── ... (similar structure)
└── ... (additional project sheets)
```

### **Rolling P&L Workbook Format**
```
Rolling P&L.xlsx
├── "Columbia" sheet              ← Maps to Columbia Villas project
├── "Dyersdale" sheet            ← Maps to Dyersdale Village project
├── "Pine Terrace" sheet         ← Maps to Pine Terrace project
├── "Summerglen" sheet           ← Maps to Summer Glen project
├── "WCL" sheet                  ← Maps to West Campus Lofts project
└── ... (other sheets)
```

## 🚀 **Getting Started**

### **Installation**
```bash
# 1. Navigate to project directory
cd account-mapping-app

# 2. Set up virtual environment (recommended)
python3 -m venv .venv
source .venv/bin/activate  # macOS/Linux
# or .venv\Scripts\activate  # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the multi-project application
python3 run_app_v2.py
```

### **Sample Data Testing**
Use the provided sample files to test multi-project functionality:
- **Source P&L.xlsx**: Contains 5 projects (Columbia Villas, Dyersdale Village, Pine Terrace, Summer Glen, West Campus Lofts)
- **Rolling P&L.xlsx**: Contains corresponding project sheets

## 🔧 **Project Management Features**

### **Automatic Project Detection**
- Scans all sheets in source workbook
- Extracts project names from cell A1
- Creates project list automatically
- Handles various Excel formats and encryption

### **Project-Specific Data Storage**
```json
{
  "source_workbook_path": "/path/to/source.xlsx",
  "rolling_workbook_path": "/path/to/rolling.xlsx",
  "current_project": "Columbia Villas",
  "projects": {
    "Columbia Villas": {
      "source_sheet": "COVI",
      "rolling_sheet": "Columbia",
      "source_range": "A8:F200",
      "rolling_range": "A1:A100",
      "mappings": { ... }
    },
    "Dyersdale Village": {
      "source_sheet": "DYVI",
      "rolling_sheet": "Dyersdale",
      ...
    }
  }
}
```

### **Smart UI State Management**
- **Context Switching**: Interface updates when projects are switched
- **Progressive Activation**: Steps unlock as requirements are met
- **Data Isolation**: No cross-contamination between projects
- **Session Recovery**: Restores last selected project on restart

## 📊 **Advanced Features**

### **Conflict Resolution**
When writing data to rolling P&L, if existing data is found:
- **Show Current Data**: Display what's already in the cell
- **Show New Data**: Display what will be written
- **User Choice**: Keep old, overwrite, or add both values
- **Excel Preservation**: Maintain formulas and formatting

### **Enhanced Error Handling**
- **File Validation**: Check workbook accessibility and format
- **Sheet Verification**: Ensure required sheets exist
- **Data Validation**: Verify range specifications and data integrity
- **Graceful Recovery**: Handle errors without data loss

### **Performance Optimization**
- **Lazy Loading**: Load project data only when selected
- **Efficient Storage**: Minimal memory footprint per project
- **Fast Switching**: Quick transitions between projects
- **Background Processing**: Non-blocking operations where possible

### **Backup & Recovery System**
- **Automatic Backups**: Creates timestamped backups during critical operations
- **Backup Contents**: Preserves project settings, mappings, and Excel files
- **Retention Policy**: Maintains recent backups for easy recovery
- **Restore Functionality**: Built-in project restoration from backups
- **Backup Naming**: Clear timestamps and month indicators (e.g., Jun_2025_Actual)

## 🔄 **Migrating from v1**

### **Key Differences**
- **Single vs Multiple**: v1 handles one project, v2 handles multiple
- **File Structure**: v2 expects workbooks with multiple sheets
- **UI Layout**: v2 adds project selection and context awareness
- **Data Storage**: v2 uses project-isolated storage format

### **Backward Compatibility**
- v1 single-sheet files can be used by treating them as single-project workbooks
- Existing mapping files can be imported for individual projects
- Same core mapping algorithms and logic

## 🐛 **Troubleshooting**

### **Common Issues**

**"No projects found"**
- Ensure cell A1 contains project names in each sheet
- Verify Excel file format is supported (.xlsx or .xls)
- Check for file encryption and provide password if needed

**"Rolling sheet not found"**
- Verify rolling workbook contains expected sheet names
- Check sheet name matching (case-sensitive)
- Ensure rolling workbook is accessible

**"Project switching not working"**
- Ensure projects were properly loaded from source workbook
- Check project settings file (project_settings.json)
- Restart application to reload projects

### **File Format Support**
- **.xlsx files**: Full support with openpyxl
- **.xls files**: Supported via xlrd
- **Encrypted files**: Supported via msoffcrypto-tool
- **Password protection**: Automatic detection and handling

## 📈 **Version History**

- **v2.1** (July 2025): Backup system improvements, file cleanup, Excel formula preservation
- **v2.0**: Multi-project support, enhanced UI, project management
- **v1.5**: Virtual environment setup, enhanced documentation
- **v1.4**: UI improvements, font sizing, reset functionality
- **v1.3**: Zoom controls and window resizing
- **v1.2**: Confidence display fixes
- **v1.1**: Monthly P&L statement generation
- **v1.0**: Initial single-project release

### **Recent Updates (July 2025)**
- **Formula Preservation**: Maintains Excel formulas when updating rolling P&L
- **Export Improvements**: Fixed "Export All Projects" functionality
- **Auto-Export**: Streamlined workflow with automatic finalization
- **File Management**: Cleaned up duplicate files and old backups
- **Performance**: Added test_performance.py for optimization testing

## 📋 **Feature Inventory**

### **UI Components & Windows**

#### **Main Interface Elements**
- **Project Selector**: Dropdown menu for quick project switching
- **File Upload Areas**: Drag-and-drop or browse functionality for Excel files
- **Range Input Fields**: Smart validation with preview capabilities
- **Status Bar**: Real-time feedback and progress indicators
- **Treeview Tables**: Interactive data grids with sorting/filtering

#### **Dialog Windows**
- **Range Preview Dialog**: Live visualization of selected Excel ranges
- **Column Selection Dialog**: Manual column mapping for exports
- **Edit Mapping Dialog**: Individual account mapping editor
- **Backup Selection Menu**: Choose from timestamped backups
- **Progress Dialogs**: Track bulk operations with cancel options

#### **Pop-out Windows**
- **Step 2 Mapping Window**: Full-featured mapping interface
- **Step 3 Preview Window**: Statement preview with export controls
- **Synchronized Updates**: Real-time data sync between windows
- **Resizable/Movable**: Flexible window positioning

### **Business Logic Modules**

#### **Account Mapping Engine** (`generate_automatic_mappings()`)
- **Fuzzy String Matching**: Intelligent pattern recognition
- **Confidence Scoring**: High/Medium/Low scoring system
- **Caching System**: Performance-optimized matching
- **Pattern Recognition**: Handles variations in account names
- **Manual Override**: User can adjust automated suggestions

#### **Excel Processing** (`_load_excel_with_cache()`, `extract_accounts()`)
- **Formula Evaluation**: Reads calculated values using openpyxl
- **Multi-sheet Support**: Handles complex workbook structures
- **Range Extraction**: Precise data selection from spreadsheets
- **Column Detection**: Smart header matching for data alignment
- **Format Preservation**: Maintains Excel formulas and formatting

#### **Project Management** (`project_manager.py`)
- **Multi-project Isolation**: Independent settings per project
- **State Persistence**: Saves all project data and UI state
- **Sheet Association**: Links projects to rolling P&L sheets
- **Settings Management**: JSON-based configuration storage

### **Data Storage & Integrations**

#### **File-Based Storage**
```json
{
  "project_settings.json": "Multi-project configurations",
  "range_settings.json": "Excel range specifications",
  "range_memory.json": "Persistent range memory",
  "[project]_mapping.json": "Individual project mappings"
}
```

#### **Caching Systems**
- **DataFrame Cache**: Loaded Excel files with timestamp validation
- **Fuzzy Match Cache**: Pre-computed similarity scores
- **Rolling Accounts Cache**: Quick access for edit dialogs
- **Target Month Cache**: Speeds up statement regeneration
- **Source Amounts Cache**: Avoids repeated extractions

#### **Excel Integration**
- **Read Operations**: openpyxl with formula evaluation
- **Write Operations**: Preserves existing formulas
- **Format Support**: .xlsx and .xls files
- **Password Protection**: Handles encrypted workbooks

### **Advanced Features**

#### **Keyboard Shortcuts**
- **Space**: Toggle checkbox selection
- **Enter**: Edit selected mapping
- **Arrow Keys**: Navigate through items
- **Cmd/Ctrl +/-**: Zoom in/out
- **Double-click**: Quick edit mode

#### **Bulk Operations**
- **Select All/None**: Quick selection management
- **Bulk Edit**: Modify multiple mappings at once
- **Export All Projects**: Single-click full export
- **Auto-process**: Complete workflow automation
- **Progress Tracking**: Visual feedback for long operations

#### **Performance Features**
- **Lazy Loading**: Data loaded only when needed
- **Smart Caching**: Timestamp-based cache validation
- **Deferred Filtering**: Timer-based filter execution
- **Optimized Updates**: Minimal UI redraws
- **Background Processing**: Non-blocking operations

### **Workflow Steps**

#### **Step 1: File Upload & Setup**
- Automatic project detection from workbook
- Rolling sheet selection per project
- Range specification with memory
- File validation and error handling

#### **Step 2: Account Mapping**
- Automatic mapping generation
- Manual adjustment capabilities
- Confidence score display
- Filter/sort functionality
- Bulk selection tools

#### **Step 3: Monthly Statement**
- Target month identification
- Historical data extraction
- Preview before generation
- Column selection options

#### **Step 4: Export & Finalize**
- Individual project export
- Bulk export all projects
- Auto-export workflow
- Progress monitoring

### **UI State Management**

#### **Persistent State**
- Filter values per project
- Sort preferences
- Zoom levels
- Window geometry
- Selection states

#### **Real-time Updates**
- Status messages
- Button state management
- Progress indicators
- Data synchronization
- Modified state tracking

### **Backup & Recovery**

#### **Automatic Backups**
- Timestamped snapshots
- Project data preservation
- Settings backup
- Metadata storage
- 8-backup retention policy

#### **Restore Features**
- Complete system restore
- Individual project restore
- Settings recovery
- UI state restoration
- Selective restoration

## 🔮 **Future Enhancements**

- **Batch Processing**: Process multiple projects simultaneously
- **Template Management**: Save and reuse project templates
- **Advanced Reporting**: Cross-project analysis and reporting
- **Database Integration**: Store projects in database instead of JSON
- **Cloud Sync**: Synchronize projects across multiple devices

---

**Note**: This is Version 2 of the Account Mapping Tool. For single-project usage, the original `main.py` application remains available and fully functional.