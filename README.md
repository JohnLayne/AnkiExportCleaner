# Anki Export Cleaner

A Python utility that cleans and formats Anki flashcard exports for better Excel compatibility while maintaining Anki re-import functionality.

## 🎯 Purpose

When exporting Anki decks as "Notes in Plain Text" with all options enabled (HTML, tags, deck names, etc.), the resulting `.txt` file contains HTML formatting that makes it difficult to work with in Excel. This script cleans the HTML formatting while preserving all essential data and maintaining compatibility with Anki re-import.

## ✨ Features

- **HTML Content Extraction**: Removes all HTML tags and formatting from content fields
- **Complete Column Preservation**: Maintains all original Anki columns (GUID, NoteType, Deck, etc.)
- **Anki Compatibility**: Preserves all required Anki import headers for seamless re-import
- **Excel-Friendly Output**: Produces clean, tab-separated data perfect for Excel import
- **Audio Reference Preservation**: Maintains `[sound:filename.mp3]` references for Anki
- **User-Friendly Interface**: Simple file dialog for selecting input files
- **Overwrite Protection**: Asks for permission before overwriting existing files
- **Detailed Logging**: Provides feedback on processing status and any skipped entries
- **Multiline Record Support**: Handles complex Anki exports with multiline HTML content

## 🚀 Quick Start

### Prerequisites

- Python 3.6 or higher
- tkinter (usually included with Python)

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/JohnLayne/AnkiExportCleaner.git
   cd AnkiExportCleaner
   ```

2. Run the script:
   ```bash
   python anki_cleaner.py
   ```

### Usage

1. **Export from Anki**:
   - Open Anki and select your deck
   - Go to File → Export
   - Choose "Notes in Plain Text (.txt)"
   - Check all boxes (HTML, tags, deck names, etc.)
   - Export the file

2. **Clean the Export**:
   - Run `anki_cleaner.py`
   - Select your exported `.txt` file when prompted
   - The script will create a `-CLEANED.txt` version

3. **Import to Excel**:
   - Open Excel
   - Go to Data → From Text/CSV
   - Select your `-CLEANED.txt` file
   - Choose "Tab" as delimiter
   - Import the data

4. **Edit in Excel** (optional):
   - Make your changes in Excel
   - **Save normally** (Ctrl+S) - keeps the `.txt` extension
   - **Note**: Excel may corrupt the UTF-8 encoding

5. **Fix Excel Encoding** (if edited):
   - Run `python fix_excel_encoding.py`
   - Select your Excel-edited file
   - The script will detect and fix any encoding issues

6. **Re-import to Anki**:
   - Use the cleaned (and optionally fixed) file to re-import to Anki
   - All formatting and audio references will be preserved

## 📁 File Structure

```
AnkiExportCleaner/
├── anki_cleaner.py          # Main script
├── fix_excel_encoding.py    # Excel encoding fix utility
├── requirements.txt         # Dependencies
├── README.md               # This file
├── .gitignore              # Git exclusions
├── Croatian_Spices.txt     # Sample input file
└── Croatian_Spices-CLEANED.txt  # Sample output file
```

## 🔧 How It Works

### Input Format
The script expects Anki exports with the following structure:
```
#separator:tab
#html:true
#guid column:1
#notetype column:2
#deck column:3
#tags column:9
GUID    NoteType    Deck    HTML_Content    English_Content    Audio_Reference    ...
```

### Processing Steps
1. **Header Preservation**: Collects and preserves all Anki import headers
2. **Multiline Record Parsing**: Handles complex records that span multiple lines
3. **HTML Extraction**: Extracts clean text from HTML content while preserving media links
4. **Field Cleaning**: Removes all HTML formatting while preserving content and audio references
5. **Output Generation**: Creates a clean, tab-separated file with proper Anki format

### Output Format
The cleaned file maintains the same structure but with clean text:
```
#separator:tab
#html:true
#guid column:1
#notetype column:2
#deck column:3
#tags column:9
GUID    NoteType    Deck    Clean_Content    Clean_English    Audio_Reference    ...
```

## 🛠️ Technical Details

### Dependencies
- **Python Standard Library**: Most functionality uses standard library
- **chardet**: For encoding detection and conversion (optional, for Excel fix)
- **tkinter**: For file dialog (included with Python)
- **csv**: For tab-separated output handling
- **re**: For HTML parsing and text cleaning

### Excel Encoding Issue
Excel has a known issue with UTF-8 encoding when saving tab-delimited text files:
- **Problem**: Excel saves files as Windows-1252 or similar encoding instead of UTF-8
- **Symptom**: Croatian diacritics (č, ć, đ, š, ž) appear as garbled characters
- **Solution**: Use `fix_excel_encoding.py` to detect and fix encoding issues after Excel editing
- **Workflow**: Regular save (Ctrl+S) in Excel works fine - the encoding fix script handles any corruption

### Key Functions

- `clean_text()`: Removes HTML entities and normalizes whitespace
- `extract_td_content()`: Extracts content from HTML while preserving media links
- `parse_anki_line()`: Handles multiline records with quoted fields
- `is_new_record()`: Detects the start of new Anki records
- `main()`: Orchestrates the entire cleaning process with file overwrite protection

## 📊 Sample Data

### Before (Raw Anki Export)
```
v,Cc7]K_>Z	JohnsLanguageNote	Croatian Johns::Vocabulary::Food::Spices - začinima	"<table><tbody><tr><td>(mr)&nbsp;bosiljak</td></tr></tbody></table>"	basil	[sound:bosiljak.mp3]
```

### After (Cleaned Output)
```
v,Cc7]K_>Z	JohnsLanguageNote	Croatian Johns::Vocabulary::Food::Spices - začinima	(mr) bosiljak	basil	[sound:bosiljak.mp3]
```

## 🔮 Future Enhancements

- [ ] Command-line interface for batch processing
- [ ] Support for different Anki note types
- [ ] CSV output option for direct Excel compatibility
- [ ] Audio file extraction and management
- [ ] GUI interface with preview functionality
- [ ] Batch processing for multiple files
- [ ] Configuration file for custom field mappings
- [ ] Progress bar for large files
- [ ] Backup creation before processing

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

## 🙏 Acknowledgments

- Anki community for the excellent flashcard software
- Python community for the robust standard library
- All contributors and users of this tool

## 📞 Support

If you encounter any issues or have questions:
1. Check the [Issues](https://github.com/JohnLayne/AnkiExportCleaner/issues) page
2. Create a new issue with detailed information
3. Include sample data if possible (with sensitive information removed)

---

**Happy studying! 📚✨** 