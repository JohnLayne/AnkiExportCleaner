# Anki Export Cleaner - Excel Integration

A Python utility with Excel integration for editing Anki flashcard exports directly in Excel.

## 🎯 Purpose

Streamlined workflow combining HTML cleaning with Excel integration. Users work in Excel with proper encoding handling and export back to Anki-compatible format.

## ✨ Features

### ✅ Working Solution (VBA + Python)
- **HTML Content Extraction**: Removes HTML tags while preserving media links
- **Automatic Encoding**: UTF-8 handling throughout
- **Complete Column Preservation**: Maintains all original Anki columns
- **Audio Reference Preservation**: Maintains `[sound:filename.mp3]` references
- **File Naming**: Exports with -CLEANED suffix
- **Custom Excel Ribbon**: Professional "Anki Tools" tab with import/export buttons
- **Immediate Setup**: No development servers required
- **Optimized for Real Usage**: Hardcoded paths enable quick deployment in production environment

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
   - Open `AnkiTool_Exporter.xlsm`
   - Look for the "Anki Tools" ribbon tab
   - Ready to use immediately!

### Alternative Setup: Manual VBA Installation

If you prefer to add the functionality to your own Excel file:

1. **Open Excel** and create a new workbook or open an existing one
2. **Enable Developer Tab**: File → Options → Customize Ribbon → Check "Developer"
3. **Open VBA Editor**: Developer tab → Visual Basic
4. **Insert Module**: Right-click on your workbook → Insert → Module
5. **Copy VBA Code**: 
   - Open `complete_vba_code.txt` in a text editor
   - Copy all the code
   - Paste it into the VBA module
6. **Save as Macro-Enabled**: File → Save As → Excel Macro-Enabled Workbook (.xlsm)
7. **Add Ribbon Buttons**: Use the functions in the VBA code to create custom ribbon buttons

## 📁 File Structure

```
AnkiExportCleaner/
├── README.md                 # This file
├── anki_excel_tool.py        # Core Python script
├── complete_vba_code.txt     # VBA source code for manual installation
├── AnkiTool_Exporter.xlsm    # Ready-to-use Excel file
├── requirements.txt          # Python dependencies
├── .gitignore               # Git exclusions
│
├── docs/                    # Documentation
│   └── DEVELOPMENT.md       # Technical development status
│
├── samples/                 # Sample files
│   ├── input/              # Raw Anki export files
│   ├── output/             # Processed output files
│   └── problematic/        # Examples of problematic filenames
│
└── tests/                  # Unit tests (future)
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
- **AnkiTool_Exporter.xlsm**: Ready-to-use Excel file with custom ribbon
- **complete_vba_code.txt**: VBA source code for manual installation or customization
- **Custom Ribbon**: "Anki Tools" ribbon tab with import/export buttons

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

### Hardcoded Paths - Intentional Design

The application uses hardcoded paths for optimal performance in a real production environment:

- **Python Script Path**: `C:\Users\JohnL\DevProjects\AnkiExportCleaner\anki_excel_tool.py`
- **Default File Location**: `C:\Users\JohnL\OneDrive\Media\Croatian Language\ANKI_EXPORT_ADDED_PRONUNCIATION\`

**Why Hardcoded Paths?**
- **Speed**: No path resolution overhead during file operations
- **Reliability**: Eliminates path-related errors in production
- **User Experience**: Direct access to commonly used folders
- **Performance**: Faster file operations without dynamic path calculations

**For Custom Deployment:**
Edit the paths in `complete_vba_code.txt` to match your environment:
- Update `GetProjectRoot()` function for your project location
- Modify `defaultPath` variables for your preferred file locations

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

### Recommended User Workflow
```
1. Export from Anki → "Croatian Johns__Vocabulary__Food__Meat and Fish - meso i riba.txt"
2. ⚠️ RENAME FILE → "Croatian_Food_Meat_Fish.txt"  
3. Import to Excel → Success with no warnings
4. Edit and Export → "Croatian_Food_Meat_Fish-CLEANED.txt"
```

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

### Python Ecosystem
- **openpyxl** - The library that enables Excel file creation and manipulation
- **Python Standard Library** - Core functionality for file processing and encoding

### Development Tools
- **Git** - Version control system
- **GitHub** - Code hosting and collaboration platform

---

**Happy studying with Excel integration! 📚✨** 