# CLAUDE.md - Account Mapping App

This file provides guidance to Claude Code (claude.ai/code) when working with the Account Mapping App repository.

## 🔒 CRITICAL: Branch Protection Rules

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
├── README_v2.md           # User documentation
├── CLAUDE.md              # This file - AI guidance
├── .gitignore             # Git ignore rules
├── project_settings.json   # Project configurations
├── range_settings.json     # Excel range settings
├── range_memory.json       # Range persistence
├── sample_source_pl.xlsx   # Sample source data
├── sample_rolling_pl.xlsx  # Sample rolling data
└── backups/               # Auto-backup directory
```

## 🎯 Project Overview

A professional Python Tkinter desktop application for mapping account descriptions between P&L spreadsheets with intelligent pattern recognition and monthly statement generation.

### **Core Features:**
- Multi-project support with isolation
- 4-step workflow (Upload → Map → Review → Generate)
- Intelligent account mapping with confidence scoring
- Excel file processing with formula preservation
- Automatic backup and restore functionality
- Advanced UI with zoom controls and popup windows

## 🔧 Development Guidelines

### **Before Making Changes:**
1. **ALWAYS** verify you're on `testing` branch:
   ```bash
   git checkout testing
   git pull origin testing
   ```

2. **File Modifications:**
   - Read files before editing using the Read tool
   - Preserve existing code style and conventions
   - Maintain backward compatibility
   - Update version numbers in comments when making significant changes

3. **Testing Changes:**
   - Use sample Excel files for testing
   - Verify all 4 workflow steps still function
   - Test with multiple projects
   - Check backup/restore functionality

### **Commit Guidelines:**
```bash
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
# 1. Verify branch
git branch --show-current

# 2. If not on testing, switch
git checkout testing

# 3. Pull latest changes
git pull origin testing

# 4. Check app status
python3 run_app_v2.py
```

### **Adding New Features:**
1. Create feature on `testing` branch
2. Test thoroughly with sample data
3. Commit with descriptive message
4. Push to testing
5. **ASK USER** before merging to main

### **Bug Fixes:**
1. Reproduce issue on `testing` branch
2. Implement fix
3. Verify fix doesn't break other features
4. Commit and push to testing
5. Create PR or **ASK USER** about merging

## ⚠️ Important Reminders

### **Never Do Without Permission:**
- Push to `main` branch
- Merge to `main` branch
- Delete any branches
- Force push to any branch
- Reset commit history
- Change repository settings

### **Always Do:**
- Work on `testing` branch by default
- Pull before starting work
- Test changes before committing
- Write clear commit messages
- Ask for confirmation on major changes
- Preserve existing functionality

## 📊 Configuration Files

### **Files to Handle Carefully:**
- `project_settings.json` - Contains user's project data
- `range_settings.json` - Excel range configurations
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
3. **Mappings** - Account mappings per project
4. **UI State** - Filter values, sort preferences, zoom level
5. **Backup State** - Automatic backup snapshots

## 🐛 Known Considerations

- Excel files with formulas need special handling
- Password-protected Excel files require msoffcrypto
- Large Excel files (1000+ rows) may need performance optimization
- Tkinter has platform-specific rendering differences

## 🤝 Collaboration Rules

1. **Pull Requests**: Always create PRs from `testing` to `main`
2. **Code Review**: Major changes should be reviewed before merging
3. **Documentation**: Update README_v2.md for user-facing changes
4. **Testing**: Ensure all sample workflows complete successfully

## 📝 Session History

When continuing work on this project:
1. Check CLAUDE_CODE_HISTORY.md for previous session context
2. Review recent commits on both branches
3. Verify no uncommitted changes exist
4. Continue from where the last session ended

## 🚦 Quick Status Check

Run these commands at the start of each session:
```bash
# Check current branch
git branch --show-current

# Check for uncommitted changes
git status

# Check for unpushed commits
git log origin/testing..testing --oneline

# Verify app launches
python3 run_app_v2.py
```

---

**Remember**: When in doubt, work on `testing` branch and ASK before touching `main`!