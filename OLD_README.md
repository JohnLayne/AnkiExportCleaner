Anki Export Cleaner - Modern Excel Integration
A Python utility with Office Add-ins integration for editing Anki flashcard exports directly in Excel with a custom ribbon interface.

🎯 Purpose
Streamlined workflow combining HTML cleaning with modern Excel integration. Users work in Excel with proper encoding handling and export back to Anki-compatible format using a professional custom ribbon interface.

✨ Features
HTML Content Extraction: Removes HTML tags while preserving media links
Modern Excel Integration: Office Add-ins with custom ribbon interface
Custom Excel Ribbon: Professional "Anki Tools" tab with 4 functional buttons
Automatic Encoding: UTF-8 handling throughout
Complete Column Preservation: Maintains all original Anki columns
Audio Reference Preservation: Maintains [sound:filename.mp3] references
GUID Preservation: Maintains Anki note GUIDs to ensure cards are updated rather than duplicated when re-imported
Hot-Reload Development: Instant updates during development
REST API Backend: Node.js server bridges Office Add-ins and Python processing
🚀 Quick Start
Prerequisites
Python 3.6+
Node.js 14+
Microsoft Excel (desktop version)
Dependencies: pip install -r requirements.txt
Setup
Clone and switch to branch:

git clone https://github.com/JohnLayne/AnkiExportCleaner.git
cd AnkiExportCleaner
git checkout office-addins-ribbon
Install Python dependencies:

pip install -r requirements.txt
Install Node.js dependencies:

cd AnkiTools
npm install
Start the backend server:

npm run start:backend
Start the Office Add-in development server:

npm start
Usage
Export from Anki: File → Export → "Notes in Plain Text (.txt)" → Check all boxes
Import to Excel: Click "Import Anki" button in the custom "Anki Tools" ribbon tab
Edit in Excel: Make changes in familiar Excel interface
Export back to Anki: Click "Export Anki" button in the ribbon
Import to Anki: Use exported file to re-import (existing cards will be updated, not duplicated)
📁 File Structure
AnkiExportCleaner/
├── README.md                 # This file
├── anki_excel_tool.py        # Core Python script
├── requirements.txt          # Python dependencies
├── .gitignore               # Git exclusions
├── QUICK_START_OFFICE_ADDINS.md  # Quick start guide
│
├── AnkiTools/               # Office Add-in project
│   ├── manifest.xml         # Add-in configuration
│   ├── package.json         # Node.js dependencies
│   ├── server.js            # Backend REST API server
│   └── src/                 # Add-in source code
│       └── commands/        # Ribbon button functions
│
├── docs/                    # Documentation
│   ├── DEVELOPMENT.md       # Current development status
│   └── DEVELOPMENT_OLD_HISTORY.md  # Historical VBA approach
│
└── samples/                 # Sample files
    ├── Croatian_Spices.txt  # Sample input with diacritics (exported from .apkg)
    └── Croatian Johns__Vocabulary__Food__Spices - začinima.apkg  # Original Anki deck
🔧 How It Works
Modern Architecture
Office Add-in Frontend: Custom Excel ribbon with 4 buttons
Node.js Backend: REST API server (port 3001)
Python Processing: Core Anki file processing logic
Hot-Reload Development: Instant updates during development
Workflow
Ribbon Import: "Import Anki" button triggers REST API call
Backend Processing: Node.js server calls Python script
HTML Cleaning: Python removes HTML tags, preserves media links and GUIDs
Excel Conversion: Converts to Excel format with formatting
Excel Editing: Users work in native Excel format
Ribbon Export: "Export Anki" button reads worksheet and converts back
Anki Import: Clean file ready for Anki (existing cards updated, not duplicated)
Key Components
anki_excel_tool.py
Enhanced AnkiCleaner class with Excel export
GUID preservation for card relationship maintenance
Professional Excel formatting
Comprehensive error handling
Office Add-in Integration
manifest.xml: Custom "Anki Tools" ribbon tab configuration
commands.js: JavaScript functions for ribbon buttons
server.js: Node.js backend with REST API
Hot-reload: Instant development updates
🛠️ Technical Details
Dependencies
Python: openpyxl for Excel file handling
Node.js: Express.js for REST API, Office Add-ins CLI tools
Office Add-ins: Office.js API, Yeoman generator
Development: Webpack, ESLint, Babel
Modern Integration
Office Add-ins platform (Microsoft-recommended approach)
REST API architecture for scalability
Hot-reload development environment
Professional ribbon interface
Automatic UTF-8 conversion
GUID preservation for card relationships
📊 Complete Workflow
Anki Export (.txt)
       ↓
Custom Ribbon "Import Anki" button
       ↓
REST API (Node.js backend)
       ↓
Python script (anki_excel_tool.py)
       ↓
Clean Excel file (.xlsx) ready for editing
       ↓
User edits in Excel
       ↓
Custom Ribbon "Export Anki" button
       ↓
REST API processes worksheet data
       ↓
Anki-compatible .txt file with UTF-8 encoding and preserved GUIDs
🔮 Future Enhancements
 Batch processing for multiple files
 Advanced Excel formatting options
 Custom field validation
 Audio file management
 Template system for different note types
 File selection dialogs
 Progress indicators for long operations
 Configuration options
🤝 Contributing
Fork the repository
Create feature branch: git checkout -b feature/new-feature
Make changes following PEP 8 guidelines
Test thoroughly with real Anki exports
Submit pull request
For development status and roadmap, see docs/DEVELOPMENT.md.

📞 Support
For issues:

Check GitHub Issues
Create issue with detailed information
Include sample data (with sensitive info removed)
🙏 Acknowledgments
This project builds upon the excellent work of several open-source communities:

Anki Development Team
Anki - The powerful spaced repetition flashcard system that makes learning efficient and effective
AnkiWeb - The community platform for sharing and discovering Anki decks
AnkiDroid - The Android app that brings Anki to mobile devices
Microsoft Office Add-ins Platform
Office.js - The JavaScript API that enables powerful Excel integrations
Office Add-ins Documentation - Comprehensive guides and examples
Yeoman Generator Ecosystem
Yeoman - The web app scaffolding tool that made this project possible
generator-office - The official Office Add-ins generator created by Microsoft
Office Add-ins CLI tools - Command-line tools for development and deployment
Node.js and Express.js
Node.js - The JavaScript runtime that powers our backend server
Express.js - The web framework that enables our REST API
Python Ecosystem
openpyxl - The library that enables Excel file creation and manipulation
Python Standard Library - Core functionality for file processing and encoding
Development Tools
Webpack - Module bundler for the Office Add-in
ESLint - Code quality and consistency
Babel - JavaScript transpilation
