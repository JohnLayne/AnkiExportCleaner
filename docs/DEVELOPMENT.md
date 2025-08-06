# Development Status & Roadmap

## 📊 Current Status

### ✅ Completed
- **Core Python Functionality**: HTML cleaning, column preservation, UTF-8 encoding (tested with VBA approach)
- **Office Add-ins Infrastructure**: Yeoman generator setup, project structure created
- **Custom Ribbon Configuration**: manifest.xml configured with "Anki Tools" tab and four buttons
- **Backend API Server**: Node.js Express server code written with filename validation
- **Frontend Commands**: JavaScript functions written for ribbon button actions
- **Documentation**: Comprehensive README.md, development docs, and structure documentation
- **File Structure**: Clean git-friendly structure with proper .gitignore
- **Filename Validation**: System to warn users about problematic Anki export filenames

### 🔄 In Progress - CURRENT PHASE
- **UNTESTED**: Office Add-in ribbon buttons have not been tested yet
- **UNTESTED**: Excel integration with Office Add-ins approach not validated
- **UNTESTED**: End-to-end workflow from ribbon buttons through Python script

### 🚧 Pending - Next Steps
1. **IMMEDIATE**: Test Office Add-in deployment and ribbon button functionality
2. **IMMEDIATE**: Validate Excel opens add-in correctly and shows custom ribbon
3. **IMMEDIATE**: Test Import/Export buttons actually work with backend API
4. **User Validation**: Test with real Anki exports and Croatian diacritics
5. **Installation Package**: Create user installation guide
6. **Enhanced Features**: File selection dialogs, progress indicators, batch processing

## 🎯 Roadmap

### Phase 1: Core Stability (Current)
- [x] **Office Add-ins Infrastructure**: Custom ribbon structure with Yeoman generator
- [x] **Backend API Code**: Node.js Express server written (untested)
- [x] **Four Ribbon Buttons Code**: Import, Export, Validate, Help functions written (untested)
- [x] **File Structure**: Git-friendly structure with proper .gitignore
- [x] **Documentation**: Updated README.md with Office Add-ins approach
- [ ] **CRITICAL**: Test Office Add-in actually loads in Excel
- [ ] **CRITICAL**: Test ribbon buttons appear and are clickable
- [ ] **CRITICAL**: Test backend API server starts and responds
- [ ] **CRITICAL**: Test end-to-end workflow actually works
- [ ] Validate with real user data and Croatian diacritics
- [ ] Fix any discovered issues

### Phase 2: User Experience (Next)
- [ ] **File Selection Dialogs**: Implement proper file picker in ribbon buttons
- [ ] **Progress Indicators**: Show progress for long operations
- [ ] **Error Handling**: Improve user feedback and error messages
- [ ] **Installation Package**: Create easy installation process

### Phase 3: Advanced Features (Future)
- [ ] **Batch Processing**: Process multiple files simultaneously
- [ ] **Advanced Validation**: Custom field validation rules
- [ ] **Template System**: Different note types and formatting options
- [ ] **Configuration Options**: User preferences and settings

## 🛠️ Technical Notes

### Architecture
```
Anki Export (.txt) → Custom Ribbon → REST API → Python Script → Excel File → Ribbon Export → Anki Import (.txt)
```

### Key Components
- **anki_excel_tool.py**: Core Python script with Excel export
- **AnkiTools/anki-tools/**: Office Add-in project structure
- **manifest.xml**: Custom "Anki Tools" ribbon tab configuration
- **commands.js**: JavaScript functions for ribbon buttons
- **server.js**: Node.js backend with REST API
- **excel/**: Alternative VBA approach (legacy support)

### Dependencies
- Python 3.6+, openpyxl, chardet
- Node.js 14+, Express.js, CORS, Multer
- Microsoft Excel (desktop version)
- Office.js API, Yeoman generator ecosystem

### Known Limitations & Critical User Warnings

#### ⚠️ Anki Export Filename Issues (CRITICAL)
**Anki automatically generates very long filenames that can cause system failures:**

**Common Anki Export Patterns:**
```
❌ PROBLEMATIC: "Croatian Johns__Vocabulary__Food__Meat and Fish - meso i riba.txt"
❌ PROBLEMATIC: "Japanese N5 Vocabulary - Chapter 1-5 Complete with Audio References.txt"
❌ PROBLEMATIC: "Medical Terminology - Cardiovascular System - Terms and Definitions.txt"
```

**Issues Caused:**
- **Windows path limits**: 260-character total path length limit
- **Excel sheet names**: Limited to 31 characters
- **Special characters**: Double underscores, spaces, and special chars cause errors
- **Zip compatibility**: Long names break compression/extraction
- **Git issues**: Some Git operations fail with long paths

**MANDATORY User Action:**
```
✅ RECOMMENDED: "Croatian_Food_Meat_Fish.txt"
✅ RECOMMENDED: "Japanese_N5_Vocab_Ch1-5.txt" 
✅ RECOMMENDED: "Medical_Cardio_Terms.txt"
```

**Filename Guidelines:**
- **Length**: Keep under 50 characters total
- **Characters**: Use only letters, numbers, underscores, hyphens
- **No spaces**: Replace with underscores or hyphens
- **No special chars**: Avoid /, \, :, *, ?, ", <, >, |, double underscores

#### Technical Limitations
- **Excel sheet names**: Limited to 31 characters and cannot contain certain characters
- **Path length**: Windows has 260-character path limit (can be extended with registry changes)
- **Unicode handling**: Most Unicode characters supported via UTF-8

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