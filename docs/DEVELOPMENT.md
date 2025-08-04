# Development Status & Roadmap

## 📊 Current Status

### ✅ Completed
- **Core Functionality**: HTML cleaning, column preservation, UTF-8 encoding
- **Excel Integration**: Python script with Excel export, VBA module
- **Documentation**: README.md, file structure reorganization
- **Excel File**: AnkiTool.xlsm with working VBA functions
- **VBA Integration**: Complete VBA code with Import/Export functions working
- **Testing**: Import function working, Export function working with UTF-8 encoding
- **UI Setup**: Manual Quick Access Toolbar approach implemented (more reliable than custom ribbon)
- **File Cleanup**: Removed redundant files and duplicate Excel files

### 🔄 In Progress
- **Final Testing**: Complete end-to-end workflow validation with Anki import
- **User Validation**: Testing with real Anki exports and Croatian diacritics

### 🚧 Pending
- **Installation Package**: Complete user installation package
- **Enhanced Features**: Batch processing, advanced validation, templates
- **Future Custom Ribbon**: May revisit custom ribbon approach for advanced users

## 🎯 Roadmap

### Phase 1: Core Stability (Current)
- [x] Complete VBA integration with Import/Export functions
- [x] Fix Export function UTF-8 encoding issues
- [x] Add instruction line filtering for clean Anki exports
- [x] **Implement Quick Access Toolbar setup** (manual approach)
- [x] Clean up redundant files and duplicates
- [ ] Complete end-to-end testing with Anki import
- [ ] Test VBA integration thoroughly
- [ ] Validate with real user data
- [ ] Fix any discovered issues

### Phase 2: User Experience (Next)
- [ ] Create installation script/package
- [ ] Improve error messages and user feedback
- [ ] Create user guide with screenshots

### Phase 3: Advanced Features (Future)
- [ ] Batch processing capabilities
- [ ] Advanced Excel formatting options
- [ ] Custom field validation rules
- [ ] Template system for different note types
- [ ] **Custom Ribbon Implementation**: Revisit automated ribbon setup for power users

## 🛠️ Technical Notes

### Architecture
```
Anki Export (.txt) → Python Script → Excel File → VBA Export → Anki Import (.txt)
```

### Key Components
- **anki_excel_tool.py**: Main Python script with Excel export
- **excel/AnkiTool.xlsm**: Ready-to-use Excel file with VBA functions
- **excel/complete_vba_code.txt**: Complete VBA code for copy-paste
- **Manual Quick Access Toolbar**: User sets up buttons manually in Excel (most reliable)

### Dependencies
- Python 3.6+, openpyxl, chardet
- Microsoft Excel with VBA enabled

### Known Limitations
- **Long filenames**: Anki exports can have very long filenames that may exceed system limits
- **Special characters**: Filenames with special characters (/, \, :, *, ?, ", <, >, |) can cause issues
- **Excel sheet names**: Limited to 31 characters and cannot contain certain characters
- **Path length**: Windows has 260-character path limit (can be extended with registry changes)

## 🔧 VBA Code Maintenance

### Current VBA Code Location
- **Source**: `excel/complete_vba_code.txt` (maintained by AI)
- **Usage**: Copy and paste into Excel VBA editor

### VBA Functions
1. **ImportFromAnki()**: Converts Anki .txt to Excel .xlsx
2. **ExportToAnki()**: Converts Excel data back to Anki .txt format with UTF-8 encoding
3. **ValidateAnkiFormat()**: Checks required fields
4. **ShowAnkiHelp()**: Displays help information
5. **CheckRequirements()**: Verifies Python installation and script availability
6. **DebugPaths()**: Shows detected paths for troubleshooting

### Maintenance Workflow
1. **AI updates**: `excel/complete_vba_code.txt` when fixes needed
2. **User copies**: From text file to Excel VBA editor
3. **Testing**: Verify functions work correctly

### Recent Fixes
- ✅ **Script path**: Fixed to use absolute project path
- ✅ **Export function**: Fixed UTF-8 encoding using ADODB.Stream
- ✅ **Error handling**: Added better error messages and line numbers
- ✅ **Instruction filtering**: Added logic to exclude Excel instruction lines from Anki export
- ✅ **UI Approach**: Switched from problematic custom ribbon to reliable manual Quick Access Toolbar
- ✅ **File cleanup**: Removed duplicate files and redundant setup scripts
- ✅ **Filename consistency**: Fixed Excel sheet naming to use input filename instead of hardcoded "Anki Cards"
- ✅ **VBA improvements**: Fixed hardcoded paths, improved Python detection, added utility functions
- ✅ **Overflow errors**: Fixed VBA Shell command overflow issues by changing Integer to Long
- ✅ **Path detection**: Enhanced GetProjectRoot() and GetScriptPath() functions for better portability
- ✅ **System requirements**: Added CheckRequirements() function for troubleshooting
- ✅ **Debug functions**: Added DebugPaths() for path verification
- ✅ **Export naming**: Fixed ExportToAnki() to suggest correct filenames based on sheet names

## ⚠️ Filename Handling & Limitations

### Anki Export Filename Issues
Anki exports can create files with very long names and special characters that may cause problems:

**Common Issues:**
- **Long filenames**: Anki deck names can be very descriptive and long
- **Special characters**: Characters like `/`, `\`, `:`, `*`, `?`, `"`, `<`, `>`, `|` are invalid in Windows filenames
- **Unicode characters**: Non-ASCII characters may cause encoding issues
- **Excel sheet name limits**: Excel sheets are limited to 31 characters

**Current Handling:**
- Python script truncates sheet names to 31 characters
- Removes "-EXCEL" suffix from sheet names
- VBA suggests filenames based on sheet names
- UTF-8 encoding handles most Unicode characters

**Recommendations:**
- **Rename files**: Consider renaming long Anki exports to shorter, simpler names
- **Avoid special characters**: Use only letters, numbers, spaces, hyphens, and underscores
- **Keep names short**: Aim for filenames under 50 characters for best compatibility

**Example:**
```
Original: "Croatian Johns__Vocabulary__Food__Meat and Fish - meso i riba.txt"
Better: "Croatian_Food_Meat_Fish.txt"
```

## 🎛️ Quick Access Toolbar Setup

### Manual Setup Process
1. **Open Excel** with `excel/AnkiTool.xlsm`
2. **Right-click Quick Access Toolbar** (top-left)
3. **Select "Customize Quick Access Toolbar"**
4. **Choose "Macros"** from dropdown
5. **Add macros**: ImportFromAnki, ExportToAnki, ValidateAnkiFormat, ShowAnkiHelp
6. **Rename buttons** as desired (e.g., "Import Anki", "Export Anki")
7. **Click "OK"** to save

### Advantages of Manual Setup
- ✅ **No VBA compilation issues**
- ✅ **No file corruption risks**
- ✅ **Easy to customize** button names and icons
- ✅ **Persistent** across Excel sessions
- ✅ **No additional dependencies**

## 📅 Next Steps

### Priority 1: Testing
1. **Test complete workflow**:
   - Import Anki .txt file using Quick Access Toolbar button
   - Edit data in Excel
   - Export back to Anki .txt using Quick Access Toolbar button
   - Import into Anki to verify compatibility
2. **Test Croatian diacritics** preservation
3. **Verify no remnant text** in exported files
4. **Test filename consistency**: Verify exported files use original filename + "-CLEANED" suffix
5. **Test VBA functions**: Verify ImportFromAnki, ExportToAnki, CheckRequirements work correctly
6. **Test path detection**: Ensure VBA correctly finds Python script and project files
7. **Test filename handling**: Verify long filenames and special characters are handled properly

### Priority 2: Documentation Updates
1. **Update README.md** with Quick Access Toolbar setup instructions
2. **Add screenshots** of working interface
3. **Create user guide** for complete workflow

### Priority 3: Future Enhancements
1. **Revisit custom ribbon** for advanced users (if needed)
2. **Create installation package** for easy distribution
3. **Add batch processing** capabilities

### Files Ready for Use:
- ✅ `excel/complete_vba_code.txt` - Updated VBA code with all fixes
- ✅ `excel/AnkiTool.xlsm` - Excel file ready for VBA code and Quick Access Toolbar setup
- ✅ `anki_excel_tool.py` - Main Python script in project root

---

**Last Updated**: January 2025 