#!/usr/bin/env python3
"""
Excel Encoding Fix Utility

This script fixes encoding issues caused by Excel when saving tab-delimited text files.
Use this after editing a cleaned Anki file in Excel to restore proper UTF-8 encoding.
"""

import os
import shutil
import sys
from pathlib import Path
from typing import Optional, Tuple
from tkinter import Tk, filedialog

# Try to import chardet, but handle gracefully if not available
try:
    import chardet
    CHARDET_AVAILABLE = True
except ImportError:
    CHARDET_AVAILABLE = False

# Constants
BACKUP_SUFFIX = '.backup'
TARGET_ENCODING = 'utf-8'
MIN_CONFIDENCE_THRESHOLD = 0.7


class ExcelEncodingFixer:
    """Utility class for fixing Excel encoding issues."""
    
    def __init__(self):
        self.backup_created = False
        self.backup_path: Optional[Path] = None
    
    def detect_encoding(self, file_path: Path) -> Tuple[str, float]:
        """Detect the encoding of a file using chardet."""
        if not CHARDET_AVAILABLE:
            print("❌ chardet not installed. Install it with: pip install chardet")
            sys.exit(1)
        
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read()
                detected = chardet.detect(raw_data)
                encoding = detected['encoding']
                confidence = detected['confidence']
            
            return encoding, confidence
            
        except Exception as e:
            print(f"❌ Error detecting encoding: {e}")
            raise
    
    def create_backup(self, file_path: Path) -> Path:
        """Create a backup of the original file."""
        try:
            backup_path = file_path.with_suffix(file_path.suffix + BACKUP_SUFFIX)
            shutil.copy2(file_path, backup_path)
            self.backup_created = True
            self.backup_path = backup_path
            return backup_path
            
        except Exception as e:
            print(f"❌ Error creating backup: {e}")
            raise
    
    def convert_encoding(self, file_path: Path, source_encoding: str) -> bool:
        """Convert file from source encoding to UTF-8."""
        try:
            # Read with detected encoding
            with open(file_path, 'r', encoding=source_encoding) as f:
                content = f.read()
            
            # Write back as UTF-8
            with open(file_path, 'w', encoding=TARGET_ENCODING, newline='') as f:
                f.write(content)
            
            return True
            
        except Exception as e:
            print(f"❌ Error converting encoding: {e}")
            return False
    
    def fix_encoding(self, file_path: Path) -> bool:
        """Main method to fix encoding issues in a file."""
        try:
            print(f"🔍 Analyzing: {file_path}")
            
            # Detect current encoding
            encoding, confidence = self.detect_encoding(file_path)
            
            if not encoding:
                print("⚠️  Could not detect encoding, assuming UTF-8")
                return True
            
            print(f"🔍 Detected encoding: {encoding} (confidence: {confidence:.2f})")
            
            # Check if conversion is needed
            if encoding.lower() == TARGET_ENCODING:
                print("✅ File is already UTF-8 encoded")
                return True
            
            # Check confidence threshold
            if confidence < MIN_CONFIDENCE_THRESHOLD:
                print(f"⚠️  Low confidence in encoding detection ({confidence:.2f} < {MIN_CONFIDENCE_THRESHOLD})")
                response = input("Continue anyway? (y/N): ").strip().lower()
                if response not in ['y', 'yes']:
                    print("Operation cancelled.")
                    return False
            
            # Create backup before modifying
            backup_path = self.create_backup(file_path)
            print(f"📁 Backup created: {backup_path}")
            
            # Convert encoding
            if self.convert_encoding(file_path, encoding):
                print(f"🔧 Successfully converted from {encoding} to {TARGET_ENCODING}")
                return True
            else:
                print("❌ Failed to convert encoding")
                return False
                
        except Exception as e:
            print(f"❌ Error fixing encoding: {e}")
            return False
    
    def select_file(self) -> Optional[Path]:
        """Open file dialog to select file to fix."""
        try:
            root = Tk()
            root.withdraw()
            
            file_path = filedialog.askopenfilename(
                title="Select the Excel-edited file to fix encoding",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            
            if not file_path:
                print("No file selected. Exiting.")
                return None
            
            return Path(file_path)
            
        except Exception as e:
            print(f"❌ Error selecting file: {e}")
            return None
    
    def run(self) -> bool:
        """Main execution method."""
        print("🔧 Excel Encoding Fix Utility")
        print("=" * 50)
        
        # Select file to fix
        file_path = self.select_file()
        if not file_path:
            return False
        
        # Check if file exists
        if not file_path.exists():
            print(f"❌ File does not exist: {file_path}")
            return False
        
        # Fix encoding
        success = self.fix_encoding(file_path)
        
        if success:
            print(f"\n✅ Success! File ready for Anki import: {file_path}")
            if self.backup_created:
                print(f"📁 Original file backed up as: {self.backup_path}")
        else:
            print(f"\n❌ Failed to fix encoding for: {file_path}")
        
        return success


def main():
    """Main entry point."""
    fixer = ExcelEncodingFixer()
    success = fixer.run()
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main() 