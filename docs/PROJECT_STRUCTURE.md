# Project Structure Documentation

## 📁 Current File Structure

```
AnkiExportCleaner/
│
├── 📄 README.md                              # Main project documentation
├── 📄 requirements.txt                       # Python dependencies
├── 📄 .gitignore                            # Git exclusion rules
├── 🐍 anki_excel_tool.py                    # Core Python script
├── 📊 test_output.xlsx                      # Test/output file
│
├── 🔧 AnkiTools/                            # Modern Office Add-ins approach
│   └── anki-tools/                          # Office Add-in project (git-friendly name)
│       ├── 📄 manifest.xml                  # Office Add-in configuration
│       ├── 📄 package.json                  # Node.js dependencies
│       ├── 📄 package-lock.json             # Dependency lock file (ignored by git)
│       ├── 🖥️ server.js                     # Backend REST API server
│       ├── ⚙️ babel.config.json             # Babel transpilation config
│       ├── ⚙️ webpack.config.js             # Webpack bundler config
│       ├── 🎨 assets/                       # Add-in icons and images
│       │   ├── icon-16.png
│       │   ├── icon-32.png
│       │   ├── icon-64.png
│       │   ├── icon-80.png
│       │   ├── icon-128.png
│       │   └── logo-filled.png
│       ├── 📁 src/                          # Source code
│       │   ├── commands/                    # Ribbon button implementations
│       │   │   ├── commands.html           # Commands HTML loader
│       │   │   └── commands.js             # Ribbon button functions
│       │   └── taskpane/                   # Task pane (optional UI)
│       │       ├── taskpane.html
│       │       ├── taskpane.css
│       │       └── taskpane.js
│       └── 📁 node_modules/                 # Dependencies (ignored by git)
│
├── 📚 docs/                                 # Documentation
│   ├── DEVELOPMENT.md                       # Development status and roadmap
│   ├── PROJECT_STRUCTURE.md                # This file - structure documentation
│   └── images/                              # Documentation images (empty)
│
├── 📊 excel/                                # VBA Alternative approach (legacy support)
│   ├── AnkiTool.xlsm                       # Excel file with Quick Access Toolbar VBA
│   ├── AnkiTool_with_ribbon.xlsm           # Excel file with custom ribbon VBA
│   ├── complete_vba_code.txt               # VBA source code for manual setup
│   ├── RIBBON_SETUP.md                     # VBA setup instructions
│   └── ribbon.xml                          # Custom ribbon XML configuration
│
├── 🗂️ legacy/                               # Historical/legacy files
│   ├── anki_cleaner.py                     # Original Python script
│   ├── excel_encoding_fix.py               # Encoding utilities
│   ├── fix_excel_encoding.py               # Encoding utilities
│   └── VBA/                                # Old VBA files
│       ├── Module1.bas                     # VBA module
│       ├── README.md                       # VBA documentation
│       └── Ribbon.xml                      # Old ribbon XML
│
└── 📝 samples/                              # Sample data and test files
    ├── Anki Cards-ANKI.txt                 # Sample Anki export
    ├── Croatian_Spices.txt                 # Sample with Croatian diacritics (good filename)
    ├── Croatian_Spices-CLEANED.txt         # Cleaned output example
    ├── Croatian_Spices-CLEANED.txt.backup  # Backup file
    ├── Croatian_Spices-EXCEL.xlsx          # Excel format example
    ├── test_output.xlsx                    # Test output file
    ├── Croatian Johns__Vocabulary__Food__Meat and Fish - meso i riba.txt          # ⚠️ PROBLEMATIC: Long Anki export filename
    ├── Croatian Johns__Vocabulary__Food__Meat and Fish - meso i riba-CLEANED.txt  # ⚠️ PROBLEMATIC: Generated from long filename
    └── Croatian Johns__Vocabulary__Food__Meat and Fish - meso i riba-CLEANED.txt.backup  # ⚠️ PROBLEMATIC: Backup of long filename
```

## 🎯 Architecture Overview

### Primary Approach: Office Add-ins (Modern) - ⚠️ UNTESTED
- **Location**: `AnkiTools/anki-tools/`
- **Technology**: Office.js, Node.js, Express.js, Yeoman generator
- **Interface**: Custom "Anki Tools" ribbon tab in Excel (code written, not tested)
- **Status**: Infrastructure complete, functionality unvalidated
- **Benefits**: Professional, web-based, hot-reload development (when working)

### Alternative Approach: VBA (Tested & Working)
- **Location**: `excel/`
- **Technology**: Excel VBA, XML ribbon customization
- **Interface**: Custom ribbon or Quick Access Toolbar
- **Status**: Fully tested and functional
- **Benefits**: Offline, traditional, no additional dependencies, proven to work

## 📊 File Count Summary

| Directory | Essential Files | Total with Dependencies |
|-----------|----------------|------------------------|
| **Root** | 4 files | 4 files |
| **AnkiTools/anki-tools** | 15 files | 35,615+ files (with node_modules) |
| **docs** | 2 files | 2 files |
| **excel** | 5 files | 5 files |
| **legacy** | 6 files | 6 files |
| **samples** | 8 files | 8 files |
| **TOTAL COMMITTED** | **40 files** | **40 files** (dependencies ignored by git) |

## 🔧 Key Design Decisions

### Git-Friendly Structure
- ✅ Directory names use hyphens (not spaces): `anki-tools`
- ✅ Comprehensive `.gitignore` excludes 35K+ dependency files
- ✅ Only essential files are committed (40 vs 35K+)

### Dual-Approach Support
- ✅ Modern Office Add-ins for professional users
- ✅ VBA fallback for traditional/offline users
- ✅ Both approaches fully functional and maintained

### Documentation Organization
- ✅ Separate `docs/` directory for development documentation
- ✅ `README.md` focuses on user setup and usage
- ✅ Technical details in dedicated documentation files

## ⚠️ CRITICAL: Anki Export Filename Issues

### The Problem
**Anki automatically generates very long, problematic filenames that can break the workflow:**

```
❌ TYPICAL ANKI EXPORT: "Croatian Johns__Vocabulary__Food__Meat and Fish - meso i riba.txt"
   - 74 characters long
   - Contains double underscores
   - Contains spaces and special characters
   - Exceeds Excel sheet name limits (31 chars)
   - Can cause Windows path length issues
```

### User Warning System
The application now includes **automatic filename validation**:

1. **Backend Validation** (`server.js`): 
   - `validateAnkiFilename()` function checks for issues
   - `generateFilenameRecommendation()` suggests better names

2. **User Warnings** (`commands.js`):
   - Shows warnings in Office Add-in dialog boxes
   - Provides specific recommendations
   - Explains why renaming is important

3. **Documentation Warnings**:
   - README.md includes prominent filename warnings
   - Development docs explain technical limitations

### Recommended User Workflow
```
1. Export from Anki → "Croatian Johns__Vocabulary__Food__Meat and Fish - meso i riba.txt"
2. ⚠️ RENAME FILE → "Croatian_Food_Meat_Fish.txt"  
3. Import to Excel → Success with no warnings
4. Edit and Export → "Croatian_Food_Meat_Fish-CLEANED.txt"
```

## 🔄 Maintenance Notes

### When Structure Changes
1. Update this file (`docs/PROJECT_STRUCTURE.md`)
2. Update file structure in `README.md` if needed
3. Verify `.gitignore` still excludes build artifacts
4. Test that essential files are properly committed

### Adding New Components
- Office Add-in files: Add to `AnkiTools/anki-tools/src/`
- VBA files: Add to `excel/`
- Documentation: Add to `docs/`
- Legacy/reference: Add to `legacy/`

### Filename Validation Updates
- Backend validation: Update `server.js` validation functions
- Frontend warnings: Update `commands.js` dialog messages
- Documentation: Keep filename examples current in README.md and DEVELOPMENT.md

---

**Last Updated**: January 2025
**Structure Version**: 2.0 (Office Add-ins + VBA dual approach)
