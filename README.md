# Anki Export Cleaner - Excel Integration

A Python utility with Excel VBA integration for editing Anki flashcard exports directly in Excel.

## 🎯 Purpose

Streamlined workflow combining HTML cleaning with Excel integration. Users work in Excel with proper encoding handling and export back to Anki-compatible format.

## ✨ Features

- **HTML Content Extraction**: Removes HTML tags while preserving media links
- **Excel Integration**: Convert cleaned files to Excel format for easy editing
- **Quick Access Toolbar**: VBA-powered import/export buttons (manual setup)
- **Automatic Encoding**: UTF-8 handling throughout
- **Complete Column Preservation**: Maintains all original Anki columns
- **Audio Reference Preservation**: Maintains `[sound:filename.mp3]` references

## 🚀 Quick Start

### Prerequisites
- Python 3.6+
- Microsoft Excel (with VBA enabled)
- Dependencies: `pip install -r requirements.txt`

### Setup
1. **Clone and switch to branch**:
   ```bash
   git clone https://github.com/JohnLayne/AnkiExportCleaner.git
   cd AnkiExportCleaner
   git checkout excel-integration
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up Excel integration**:
   - Open `excel/AnkiTool.xlsm` in Excel
   - Enable macros when prompted
   - **Import VBA code**: 
     - Press Alt+F11 (opens VBA editor)
     - Double-click Module11 in left panel
     - Copy all code from `excel/complete_vba_code.txt`
     - Paste to replace existing code
     - Save Excel file
   - **Set up Quick Access Toolbar**:
     - Right-click Quick Access Toolbar (top-left)
     - Select "Customize Quick Access Toolbar"
     - Choose "Macros" from dropdown
     - Add: ImportFromAnki, ExportToAnki, ValidateAnkiFormat, ShowAnkiHelp
     - Rename buttons as desired (e.g., "Import Anki", "Export Anki")
     - Click "OK" to save

### Usage

1. **Export from Anki**: File → Export → "Notes in Plain Text (.txt)" → Check all boxes
2. **Import to Excel**: Click "Import from Anki" button in Quick Access Toolbar
3. **Edit in Excel**: Make changes in familiar Excel interface
4. **Export back to Anki**: Click "Export to Anki" button in Quick Access Toolbar
5. **Import to Anki**: Use exported file to re-import

## 📁 File Structure

```
excel-integration/
├── README.md                 # This file
├── anki_excel_tool.py        # Main Python script
├── requirements.txt          # Dependencies
├── .gitignore               # Git exclusions
│
├── excel/                   # Excel files
│   ├── AnkiTool.xlsm        # Ready-to-use Excel file
│   └── complete_vba_code.txt # VBA code for copy-paste
│
├── docs/                    # Documentation
│   └── development.md       # Development status
│
├── samples/                 # Sample files
│   └── Croatian_Spices.txt  # Sample input
│
└── legacy/                  # Old files
    ├── anki_cleaner.py      # Original script
    └── VBA/                 # Old VBA files
```

## 🔧 How It Works

### Workflow
1. **VBA Import**: Quick Access Toolbar button calls Python script
2. **HTML Cleaning**: Python removes HTML tags, preserves media links
3. **Excel Conversion**: Converts to Excel format with formatting
4. **Excel Editing**: Users work in native Excel format
5. **VBA Export**: Quick Access Toolbar button converts back to Anki format with UTF-8
6. **Anki Import**: Clean file ready for Anki

### Key Components

#### anki_excel_tool.py
- Enhanced AnkiCleaner class with Excel export
- Command line support for VBA integration
- Professional Excel formatting
- Comprehensive error handling

#### Excel Integration
- **AnkiTool.xlsm**: Ready-to-use Excel file with instructions
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

---

**Happy studying with Excel integration! 📚✨** 