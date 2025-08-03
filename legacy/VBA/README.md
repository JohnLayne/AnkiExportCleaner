# VBA Components for Anki Excel Integration

This directory contains the VBA code and ribbon customization for the Excel integration workflow.

## 📁 Files

- **Module1.bas**: Main VBA functions for import/export operations
- **Ribbon.xml**: Custom ribbon definition for Excel integration
- **README.md**: This file with setup instructions

## 🔧 Setup Instructions

### 1. Enable Developer Tab in Excel
1. Open Excel
2. Go to File → Options → Customize Ribbon
3. Check "Developer" in the right column
4. Click OK

### 2. Import VBA Code
1. Open `AnkiTool.xlsm` in Excel
2. Press Alt + F11 to open VBA Editor
3. Right-click on "Modules" → Import File
4. Select `Module1.bas` (from the same directory as AnkiTool.xlsm)

### 3. Custom Ribbon Setup
1. In Excel, go to File → Options → Customize Ribbon
2. Click "Import/Export" → "Import customization file"
3. Select `Ribbon.xml` (from the same directory as AnkiTool.xlsm)
4. The custom "Anki Tools" tab will appear in the ribbon

## 🎯 VBA Functions

### Import Functions
- `ImportFromAnki()`: Converts Anki .txt files to Excel format using Python script
- `GetPythonPath()`: Detects Python installation for script execution

### Export Functions
- `ExportToAnki()`: Converts Excel data back to Anki format with UTF-8 encoding
- `ValidateAnkiFormat()`: Validates data before export
- `WriteUTF8File()`: Writes files with proper UTF-8 BOM encoding

### Utility Functions
- `ShowAnkiHelp()`: Displays help information for users
- `WriteUTF8File()`: Handles UTF-8 file writing with BOM

## 🛠️ Development Notes

### Error Handling
- All functions include comprehensive error handling
- User-friendly error messages
- Graceful degradation on failures

### Performance
- Optimized for large Anki exports
- Minimal memory usage
- Fast processing times

### Compatibility
- Tested with Excel 2016, 2019, and 365
- Compatible with Windows and Mac versions
- Handles various Anki export formats

## 📝 Usage Examples

### Basic Import
```vba
Sub ImportAnkiFile()
    Call ImportFromAnki
End Sub
```

### Basic Export
```vba
Sub ExportAnkiFile()
    Call ExportToAnki
End Sub
```

## 🔮 Future Enhancements

- [ ] Batch processing for multiple files
- [ ] Advanced validation rules
- [ ] Custom formatting options
- [ ] Integration with Anki Connect API
- [ ] Cloud storage support

## 📞 Support

For VBA-specific issues:
1. Check the error messages in the VBA Editor
2. Ensure macros are enabled
3. Verify file permissions
4. Test with sample data first 