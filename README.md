# Anki Export Cleaner - Excel Integration

A Python utility with Excel VBA integration that provides a streamlined workflow for editing Anki flashcard exports directly in Excel with proper encoding handling.

## 🎯 Purpose

This branch provides an enhanced workflow that combines the HTML cleaning capabilities of the main branch with direct Excel integration. Users can work in Excel with proper encoding handling and export back to Anki-compatible format with a single click.

## ✨ Features

- **HTML Content Extraction**: Removes all HTML tags and formatting from content fields
- **Direct Excel Integration**: Convert cleaned files to Excel format for easy editing
- **Custom Excel Ribbon**: VBA-powered ribbon buttons for import/export operations
- **Automatic Encoding Handling**: No more manual encoding fixes needed
- **Complete Column Preservation**: Maintains all original Anki columns
- **Anki Compatibility**: Preserves all required Anki import headers
- **Audio Reference Preservation**: Maintains `[sound:filename.mp3]` references
- **User-Friendly Interface**: Simple file dialogs and Excel ribbon integration

## 🚀 Quick Start

### Prerequisites

- Python 3.6 or higher
- Microsoft Excel (with VBA enabled)
- tkinter (usually included with Python)
- chardet (for encoding detection)
- openpyxl (for Excel file handling)

### Installation

1. **Clone the repository and switch to this branch**:
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
   - Copy `anki_excel_tool.py` to your working directory
   - Create a new Excel file with macros enabled (.xlsm)
   - Import the VBA code from `VBA/Module1.bas`
   - Set up the custom ribbon using `VBA/Ribbon.xml`
   - See `VBA/README.md` for detailed setup instructions

### Usage

1. **Export from Anki**:
   - Open Anki and select your deck
   - Go to File → Export
   - Choose "Notes in Plain Text (.txt)"
   - Check all boxes (HTML, tags, deck names, etc.)
   - Export the file

2. **Convert to Excel**:
   - Run `python anki_excel_tool.py`
   - Select your exported `.txt` file
   - The script will create an `.xlsx` file ready for Excel

3. **Edit in Excel**:
   - Open the generated `.xlsx` file
   - Make your changes in the familiar Excel interface
   - No encoding issues to worry about

4. **Export back to Anki**:
   - Click the "Export to Anki" button in the custom ribbon
   - Choose your save location
   - The file will be automatically converted to Anki-compatible format

5. **Import to Anki**:
   - Use the exported file to re-import to Anki
   - All formatting, audio references, and diacritics will be preserved

## 📁 File Structure

```
excel-integration/
├── anki_excel_tool.py      # Enhanced version with Excel export
├── excel_encoding_fix.py   # Encoding fix utility (backup)
├── VBA/                    # VBA code files
│   ├── Module1.bas         # Main VBA functions
│   ├── Ribbon.xml          # Custom ribbon definition
│   └── README.md           # VBA setup instructions
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── DEVELOPMENT.md          # Development status and roadmap
├── .gitignore              # Git exclusions
├── Croatian_Spices.txt     # Sample input file
└── Croatian_Spices-CLEANED.txt  # Sample output file
```

## 🔧 How It Works

### Enhanced Workflow
1. **HTML Cleaning**: Uses the proven HTML cleaning logic from the main branch
2. **Excel Conversion**: Converts cleaned data to Excel format with proper formatting
3. **Excel Editing**: Users work in native Excel format (no encoding issues)
4. **VBA Export**: Custom ribbon button handles conversion back to Anki format
5. **Anki Import**: Clean, properly encoded file ready for Anki

### Key Components

#### anki_excel_tool.py
- **Enhanced AnkiCleaner class**: Adds Excel export functionality
- **Excel formatting**: Applies proper formatting for better editing experience
- **Header preservation**: Maintains all Anki import headers
- **Error handling**: Comprehensive error handling throughout

#### VBA Components
- **Custom ribbon**: "Import from Anki" and "Export to Anki" buttons
- **VBA macros**: Handle file conversion and encoding
- **User-friendly interface**: Simple one-click operations
- **Encoding handling**: Automatic UTF-8 conversion

#### VBA Components
- **Module1.bas**: Main VBA functions for import/export
- **Ribbon.xml**: Custom ribbon definition
- **Error handling**: User-friendly error messages

## 🛠️ Technical Details

### Dependencies
- **Python Standard Library**: Core functionality
- **openpyxl**: Excel file handling and formatting
- **chardet**: Encoding detection
- **tkinter**: File dialogs
- **pathlib**: Modern path handling
- **dataclasses**: Structured data representation

### Excel Integration
- **Native Excel format**: No encoding issues during editing
- **VBA automation**: Streamlined import/export process
- **Custom ribbon**: Professional integration with Excel
- **Automatic encoding**: Handles UTF-8 conversion automatically

## 📊 Sample Workflow

### Step 1: Anki Export
```
Raw Anki export with HTML formatting
↓
anki_excel_tool.py
↓
Clean Excel file (.xlsx) ready for editing
```

### Step 2: Excel Editing
```
Open in Excel → Make changes → Save
(No encoding issues, familiar interface)
```

### Step 3: Export to Anki
```
Click "Export to Anki" ribbon button
↓
Anki-compatible .txt file with proper encoding
```

## 🔮 Future Enhancements

- [ ] Batch processing for multiple files
- [ ] Advanced Excel formatting options
- [ ] Custom field validation
- [ ] Audio file management
- [ ] Template system for different note types
- [ ] Integration with Anki Connect API
- [ ] Cloud storage integration
- [ ] Multi-language support

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test with various Anki export formats
5. Submit a pull request

### Code Style
- Follow PEP 8 guidelines
- Use type hints for all functions
- Add comprehensive docstrings
- Include error handling for all operations
- Test VBA macros thoroughly

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

## 🙏 Acknowledgments

- Anki community for the excellent flashcard software
- Python community for the robust standard library
- Excel VBA community for integration techniques
- All contributors and users of this tool

## 📞 Support

If you encounter any issues or have questions:
1. Check the [Issues](https://github.com/JohnLayne/AnkiExportCleaner/issues) page
2. Create a new issue with detailed information
3. Include sample data if possible (with sensitive information removed)

---

**Happy studying with Excel integration! 📚✨** 