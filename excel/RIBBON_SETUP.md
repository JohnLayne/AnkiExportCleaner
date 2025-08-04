# Anki Tools Ribbon Setup

## Quick Setup (Manual)

### Option 1: Quick Access Toolbar (Easiest)
1. Open Excel with macros enabled
2. Right-click on any button in the ribbon
3. Select "Customize Quick Access Toolbar"
4. In the left panel, select "Macros"
5. Add these macros to the Quick Access Toolbar:
   - `ImportFromAnki`
   - `ExportToAnki`
   - `CheckRequirements`
   - `ShowAnkiHelp`

### Option 2: Custom Ribbon Tab (Advanced)
1. Open Excel with macros enabled
2. Go to File → Options → Customize Ribbon
3. Click "New Tab" and rename it to "Anki Tools"
4. Create groups and add the macros as buttons

## Functions Available

### Import from Anki
- Converts Anki export files to Excel format
- Cleans HTML and preserves audio links
- Creates a new Excel file with your data

### Export to Anki
- Exports Excel data back to Anki format
- Uses UTF-8 encoding for proper character support
- Maintains all required Anki fields

### Validate Format
- Checks that all required fields are present
- Validates GUID, deck name, and content fields
- Shows detailed error messages

### Check Requirements
- Verifies Python installation
- Checks if the Python script is available
- Shows system status

### Show Help
- Displays detailed usage instructions
- Explains the workflow
- Provides troubleshooting tips

## Requirements
- Python installed and in PATH
- `anki_excel_tool.py` in project root directory
- Macros enabled in Excel
- Excel file should be in the `excel/` subdirectory

## Troubleshooting
- Use "Check Requirements" to verify your setup
- Ensure the Excel file is in the correct location
- Make sure Python is accessible from the command line 