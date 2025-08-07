# Anki Export Cleaner - Excel Integration

A Python utility with Excel integration for editing Anki flashcard exports directly in Excel.

## 🎯 Purpose

Streamlined workflow combining HTML cleaning with Excel integration. Users work in Excel with proper encoding handling and export back to Anki-compatible format.

## ✨ Features

### ✅ Working Solution (VBA Approach)
- **HTML Content Extraction**: Removes HTML tags while preserving media links
- **Automatic Encoding**: UTF-8 handling throughout
- **Complete Column Preservation**: Maintains all original Anki columns
- **Audio Reference Preservation**: Maintains `[sound:filename.mp3]` references
- **File Naming**: Exports with -CLEANED suffix
- **Custom Excel Ribbon**: Professional "Anki Tools" tab with import/export buttons
- **Immediate Setup**: No development servers required

### ❌ Failed Approach (Office Add-ins)
- **Modern Excel Integration**: Office Add-ins with custom ribbon interface (built but failed)
- **Complex Setup**: Required 3 different servers running simultaneously
- **Protocol Issues**: HTTPS/HTTP mismatches between frontend and backend
- **CORS Problems**: Cross-origin request issues in development
- **Server Dependencies**: Overly complex for simple file processing task

## ⚠️ Current Status

**RECOMMENDED**: Use the **VBA approach** in `excel/AnkiTool_with_ribbon.xlsm` - it's fully tested and works immediately.

**NOT RECOMMENDED**: The Office Add-ins approach in `AnkiTools/` was built but failed due to:
- Complex multi-server setup requirements
- HTTPS/HTTP protocol mismatches
- CORS issues between frontend and backend
- No real advantage over VBA for this use case

## 🚀 Quick Start

### Prerequisites
- Python 3.6+
- Microsoft Excel (desktop version)
- Dependencies: `pip install -r requirements.txt`

### Setup
1. **Clone repository**:
   ```bash
   git clone https://github.com/JohnLayne/AnkiExportCleaner.git
   cd AnkiExportCleaner
   ```

2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Open the Excel file**:
   - Open `excel/AnkiTool_with_ribbon.xlsm`
   - Look for the "Anki Tools" ribbon tab
   - Ready to use immediately!

### Usage

1. **Export from Anki**: File → Export → "Notes in Plain Text (.txt)" → Check all boxes
2. **⚠️ IMPORTANT - Rename Long Filenames**: Anki exports often have very long names that can cause issues:
   ```
   ❌ BAD:  "Croatian Johns__Vocabulary__Food__Meat and Fish - meso i riba.txt"
   ✅ GOOD: "Croatian_Food_Meat_Fish.txt"
   ```
   **Recommendation**: Keep filenames under 50 characters, use only letters, numbers, underscores, and hyphens.

3. **Import to Excel**: Click "Import Anki" button in the "Anki Tools" ribbon tab
4. **Edit in Excel**: Make changes in familiar Excel interface
5. **Export back to Anki**: Click "Export Anki" button in the ribbon
6. **Import to Anki**: Use exported file to re-import (existing cards will be updated, not duplicated)

## 📁 File Structure

```
AnkiExportCleaner/
├── README.md                 # This file
├── anki_excel_tool.py        # Core Python script
├── requirements.txt          # Python dependencies
├── .gitignore               # Git exclusions
│
├── excel/                   # ✅ WORKING VBA APPROACH
│   ├── AnkiTool_with_ribbon.xlsm  # Ready-to-use Excel file
│   ├── AnkiTool.xlsm        # Alternative VBA file
│   ├── complete_vba_code.txt # VBA code for manual setup
│   └── ribbon.xml           # Custom ribbon configuration
│
├── AnkiTools/               # ❌ FAILED OFFICE ADD-INS APPROACH
│   └── anki-tools/          # Office Add-in files (not recommended)
│       ├── manifest.xml     # Add-in configuration
│       ├── package.json     # Node.js dependencies
│       ├── server.js        # Backend REST API server
│       └── src/             # Add-in source code
│
├── docs/                    # Documentation
│   └── DEVELOPMENT.md       # Current development status
│
└── samples/                 # Sample files
    └── Croatian_Spices.txt  # Sample input with diacritics
```

## 🔧 How It Works

### Workflow
1. **VBA Import**: Ribbon button calls Python script
2. **HTML Cleaning**: Python removes HTML tags, preserves media links
3. **Excel Conversion**: Converts to Excel format with formatting
4. **Excel Editing**: Users work in native Excel format
5. **VBA Export**: Ribbon button converts back to Anki format with UTF-8
6. **Anki Import**: Clean file ready for Anki

### Key Components

#### anki_excel_tool.py
- Enhanced AnkiCleaner class with Excel export
- Command line support for VBA integration
- Professional Excel formatting
- Comprehensive error handling

#### Excel Integration
- **AnkiTool_with_ribbon.xlsm**: Ready-to-use Excel file with custom ribbon
- **Module1.bas**: VBA functions for import/export
- **Ribbon.xml**: Custom "Anki Tools" ribbon tab

## 🛠️ Technical Details

### Dependencies
- **Python Standard Library**: Core functionality
- **openpyxl**: Excel file handling
- **chardet**: Encoding detection
- **tkinter**: File dialogs

### Excel Integration
- Native Excel format (no encoding issues)
- VBA automation for streamlined process
- Custom ribbon for professional integration
- Automatic UTF-8 conversion

## 📊 Complete Workflow

```
Anki Export (.txt)
       ↓
VBA "Import from Anki" button
       ↓
Python script (anki_excel_tool.py)
       ↓
Clean Excel file (.xlsx) ready for editing
       ↓
User edits in Excel
       ↓
VBA "Export to Anki" button
       ↓
Anki-compatible .txt file with UTF-8 encoding
```

## 🔮 Future Enhancements

- [ ] Batch processing for multiple files
- [ ] Advanced Excel formatting options
- [ ] Custom field validation
- [ ] Audio file management
- [ ] Template system for different note types

## 🤝 Contributing

1. Fork the repository
2. Create feature branch: `git checkout -b feature/new-feature`
3. Make changes following PEP 8 guidelines
4. Test thoroughly with real Anki exports
5. Submit pull request

For development status and roadmap, see [docs/development.md](docs/development.md).

## 📞 Support

For issues:
1. Check [GitHub Issues](https://github.com/JohnLayne/AnkiExportCleaner/issues)
2. Create issue with detailed information
3. Include sample data (with sensitive info removed)

## 🙏 Acknowledgments

This project builds upon the excellent work of several open-source communities:

### Anki Development Team
- **Anki** - The powerful spaced repetition flashcard system that makes learning efficient and effective
- **AnkiWeb** - The community platform for sharing and discovering Anki decks
- **AnkiDroid** - The Android app that brings Anki to mobile devices

### Microsoft Office Platform
- **VBA (Visual Basic for Applications)** - The automation language that enables Excel integration
- **Office Ribbon XML** - The framework for custom Excel ribbon interfaces

### Python Ecosystem
- **openpyxl** - The library that enables Excel file creation and manipulation
- **Python Standard Library** - Core functionality for file processing and encoding

### Development Tools
- **Git** - Version control system
- **GitHub** - Code hosting and collaboration platform

---

**Happy studying with Excel integration! 📚✨** 