# CLAUDE.md - Account Mapping App Development Guide

This file provides guidance to Claude Code (claude.ai/code) when working with the Account Mapping App repository.

## 🚨 CRITICAL: Start Here

### **BEFORE ANY DEVELOPMENT WORK:**

When starting from a fresh chat or new session, **ALWAYS** follow this sequence:

1. **Read README_v2.md FIRST** - This contains comprehensive documentation about:
   - Current application state and architecture
   - Complete feature inventory and workflows
   - Data management and persistence logic
   - Recent bug fixes and enhancements
   - Technical implementation details

2. **Review Recent Changes** - Check git log for latest updates:
   ```bash
   git log --oneline -10
   ```

3. **Verify Branch** - Ensure you're on the correct branch:
   ```bash
   git branch --show-current
   ```

### **Understanding the Application:**
The README_v2.md file contains everything you need to understand:
- **Architecture**: 3-file structure (main_v2.py, project_manager.py, run_app_v2.py)
- **Workflow**: 4-step process for mapping accounts
- **Data Flow**: How mappings are created, stored, and persisted
- **Key Features**: Multi-project support, intelligent mapping, manual overrides
- **Recent Fixes**: Account persistence issues, mapping priority system

## 🔒 Branch Protection Rules

### **ALWAYS FOLLOW THESE RULES:**

1. **DEFAULT BRANCH**: Always work on the `testing` branch unless explicitly instructed otherwise
2. **MAIN BRANCH PROTECTION**: 
   - ⛔ **NEVER** commit directly to `main` branch without explicit user confirmation
   - ⛔ **NEVER** push to `main` branch without user approval
   - ⚠️  Always ask: "Do you want me to commit this to the main branch?" before any main branch operations
3. **BRANCH WORKFLOW**:
   ```bash
   # ALWAYS check current branch before any work
   git branch --show-current
   
   # If not on testing, switch immediately
   git checkout testing
   
   # Pull latest changes before starting work
   git pull origin testing
   ```

### **Safe Git Operations Checklist:**
- [ ] Verify current branch is `testing` before making changes
- [ ] Create commits on `testing` branch only
- [ ] Get explicit permission before merging to `main`
- [ ] Use pull requests for `testing` → `main` merges when possible
- [ ] Always provide clear commit messages describing changes

## 📁 Repository Information

**Repository**: https://github.com/ydinai001/account-mapping-app  
**Main Branch**: `main` (production - protected)  
**Development Branch**: `testing` (active development)  
**Language**: Python 3.13+  
**Framework**: Tkinter (Desktop GUI)

## 🏗️ Project Structure

```
account-mapping-app/
├── main_v2.py              # Main application (5800+ lines)
├── run_app_v2.py           # Application launcher
├── project_manager.py      # Project management utilities
├── test_performance.py     # Performance testing
├── requirements.txt        # Python dependencies
├── README_v2.md           # COMPREHENSIVE DOCUMENTATION (READ FIRST!)
├── CLAUDE.md              # This file - AI guidance
├── .gitignore             # Git ignore rules
├── project_settings.json   # Project configurations
├── range_settings.json     # Excel range settings
├── range_memory.json       # Range persistence
├── sample_source_pl.xlsx   # Sample source data
├── sample_rolling_pl.xlsx  # Sample rolling data
└── backups/               # Auto-backup directory
```

## 🎯 Key Application Concepts

### **Multi-Project Architecture**
- Each project is isolated with its own settings and mappings
- Projects are detected from Excel sheet names (cell A1)
- Project switching preserves all data

### **Mapping Priority System**
1. **User Manual Edits** (Highest) - Never overwritten
2. **project_settings.json** (Primary) - Main source of truth
3. **External mapping files** (Lowest) - Only loaded if no mappings exist

### **Data Persistence**
- All changes saved immediately to `project_settings.json`
- Automatic backups created during critical operations
- Session state restored on restart

### **Recent Critical Fixes (August 2025)**
- **Issue**: New accounts (8540, 8505) weren't persisting after restart
- **Cause**: Old mapping files overwriting updated mappings
- **Solution**: Modified `load_mappings_from_saved_file()` to skip if mappings exist
- **Result**: New accounts now persist correctly

## 🔧 Development Guidelines

### **Before Making Changes:**
1. **ALWAYS** read README_v2.md for current state
2. **Verify** you're on `testing` branch:
   ```bash
   git checkout testing
   git pull origin testing
   ```

3. **File Modifications:**
   - Read files before editing using the Read tool
   - Preserve existing code style and conventions
   - Maintain backward compatibility
   - Update version numbers in comments when making significant changes

4. **Testing Changes:**
   - Use sample Excel files for testing
   - Verify all 4 workflow steps still function
   - Test with multiple projects
   - Check backup/restore functionality

### **Commit Guidelines:**

**CRITICAL: Always wait for user to test changes before committing!**

```bash
# WORKFLOW:
# 1. Make changes
# 2. Inform user changes are ready for testing
# 3. WAIT for user to test
# 4. ONLY commit after user confirms the changes work

# On testing branch (default)
git add .
git commit -m "feat: description of feature"  # or fix:, docs:, refactor:
git push origin testing

# For main branch (requires permission)
# ASK USER FIRST: "Should I merge these changes to main?"
# Only proceed if user confirms
```

### **Commit Message Format:**
- `feat:` New feature
- `fix:` Bug fix
- `docs:` Documentation changes
- `refactor:` Code refactoring
- `test:` Test additions/changes
- `chore:` Maintenance tasks

## 🚀 Common Development Tasks

### **Starting Development Session:**
```bash
# 1. Read comprehensive documentation
cat README_v2.md

# 2. Verify branch
git branch --show-current

# 3. If not on testing, switch
git checkout testing

# 4. Pull latest changes
git pull origin testing

# 5. Check app status
python3 run_app_v2.py
```

### **Adding New Features:**
1. Review README_v2.md for current architecture
2. Create feature on `testing` branch
3. Test thoroughly with sample data
4. Commit with descriptive message
5. Push to testing
6. **ASK USER** before merging to main

### **Bug Fixes:**
1. Check README_v2.md troubleshooting section
2. Reproduce issue on `testing` branch
3. Implement fix
4. Verify fix doesn't break other features
5. Commit and push to testing
6. Create PR or **ASK USER** about merging

## ⚠️ Important Reminders

### **Never Do Without Permission:**
- Push to `main` branch
- Merge to `main` branch
- Delete any branches
- Force push to any branch
- Reset commit history
- Change repository settings
- **NEVER commit changes to git before the user has tested them**

### **Always Do:**
- Read README_v2.md at session start
- Work on `testing` branch by default
- Pull before starting work
- **Wait for user to test changes before committing**
- Test changes before committing
- Write clear commit messages
- Ask for confirmation on major changes
- Preserve existing functionality

## 📊 Configuration Files

### **Files to Handle Carefully:**
- `project_settings.json` - Contains user's project data and mappings
- `range_settings.json` - Excel range configurations
- `range_memory.json` - Per-project range memory
- `*.json` mapping files - User's account mappings
- Files in `backups/` - User's backup data

### **Safe to Modify:**
- Python source files (*.py)
- Documentation files (*.md)
- requirements.txt
- .gitignore

## 🔄 Workflow States

The app maintains several states that should be preserved:
1. **Project Selection** - Current active project
2. **File Paths** - Source and rolling P&L paths
3. **Mappings** - Account mappings per project (stored in project_settings.json)
4. **UI State** - Filter values, sort preferences, zoom level
5. **Backup State** - Automatic backup snapshots

## 🐛 Known Considerations

- Excel files with formulas need special handling (use openpyxl)
- Password-protected Excel files require msoffcrypto
- Large Excel files (1000+ rows) may need performance optimization
- Tkinter has platform-specific rendering differences
- Old mapping files should not overwrite project_settings.json

## 🤝 Collaboration Rules

1. **Documentation First**: Always read README_v2.md before starting work
2. **Pull Requests**: Always create PRs from `testing` to `main`
3. **Code Review**: Major changes should be reviewed before merging
4. **Documentation Updates**: Update README_v2.md for user-facing changes
5. **Testing**: Ensure all sample workflows complete successfully

## 📝 Session History

When continuing work on this project:
1. **First**: Read README_v2.md for comprehensive understanding
2. Check CLAUDE_CODE_HISTORY.md for previous session context
3. Review recent commits on both branches
4. Verify no uncommitted changes exist
5. Continue from where the last session ended

## 🚦 Quick Status Check

Run these commands at the start of each session:
```bash
# Check comprehensive documentation
head -50 README_v2.md

# Check current branch
git branch --show-current

# Check for uncommitted changes
git status

# Check for unpushed commits
git log origin/testing..testing --oneline

# Verify app launches
python3 run_app_v2.py
```

## 📚 Key Functions and Methods to Know

### **main_v2.py** - Critical Functions:
- `generate_mappings()` - Creates automatic account mappings
- `load_mappings_from_saved_file()` - Loads external mapping files (now skips if mappings exist)
- `check_and_add_new_accounts()` - Detects and adds new accounts
- `populate_mapping_tree()` - Updates UI with mapping data
- `extract_monthly_amounts()` - Gets amounts from source P&L

### **project_manager.py** - Core Classes:
- `Project` - Encapsulates project data and settings
- `ProjectManager` - Manages multiple projects and persistence
- `save_settings()` - Persists all project data to JSON
- `load_settings()` - Loads project data on startup

## 🎓 Learning from Past Issues

### **Account Persistence Issue (August 2025)**
- **Problem**: New accounts not persisting after restart
- **Investigation**: Traced through save/load cycle
- **Discovery**: Old mapping files overwriting updates
- **Solution**: Skip loading mapping files if mappings exist
- **Lesson**: Always check for multiple data sources

---

**Remember**: 
1. **ALWAYS read README_v2.md first** for complete understanding
2. Work on `testing` branch and ASK before touching `main`
3. The README contains all architectural details, workflows, and current state
4. When in doubt, refer to the comprehensive documentation