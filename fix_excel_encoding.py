#!/usr/bin/env python3
"""
Excel Encoding Fix Utility

This script fixes encoding issues caused by Excel when saving tab-delimited text files.
Use this after editing a cleaned Anki file in Excel to restore proper UTF-8 encoding.
"""

import os
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def fix_excel_encoding(file_path: str) -> str:
    """Fix encoding issues caused by Excel saving files."""
    import chardet
    
    # Read the file and detect encoding
    with open(file_path, 'rb') as f:
        raw_data = f.read()
        detected = chardet.detect(raw_data)
        encoding = detected['encoding']
    
    print(f"🔍 Detected encoding: {encoding} (confidence: {detected['confidence']:.2f})")
    
    # If it's not UTF-8, convert it
    if encoding and encoding.lower() != 'utf-8':
        try:
            # Decode with detected encoding and re-encode as UTF-8
            content = raw_data.decode(encoding)
            
            # Create backup
            backup_path = file_path + '.backup'
            import shutil
            shutil.copy2(file_path, backup_path)
            
            # Write back as UTF-8
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                f.write(content)
            
            print(f"🔧 Fixed encoding from {encoding} to UTF-8")
            print(f"📁 Backup saved as '{backup_path}'")
            return file_path
            
        except Exception as e:
            print(f"⚠️  Could not fix encoding: {e}")
            return file_path
    else:
        print("✅ File is already UTF-8 encoded")
    
    return file_path

def main():
    print("🔧 Excel Encoding Fix Utility")
    print("=" * 40)
    
    # Let user pick the file to fix
    root = Tk()
    root.withdraw()
    file_path = askopenfilename(
        title="Select the Excel-edited file to fix encoding",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
    
    if not file_path:
        print("No file selected. Exiting.")
        sys.exit(1)
    
    print(f"📁 Processing: {file_path}")
    
    # Fix the encoding
    fixed_path = fix_excel_encoding(file_path)
    
    print(f"\n✅ Done! File ready for Anki import: {fixed_path}")

if __name__ == "__main__":
    main() 