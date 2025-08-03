

#!/usr/bin/env python3
import csv
import os
import re
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def clean_text(txt: str) -> str:
    """Collapse whitespace, decode &nbsp;, trim."""
    return re.sub(r'\s+', ' ', txt.replace('&nbsp;', ' ')).strip()

def extract_td_content(html_content: str) -> str:
    """Extract text content from HTML, removing all HTML tags but preserving media links."""
    # Remove quotes if present
    html_content = html_content.strip('"')
    
    # First, preserve any media links (like [sound:filename.mp3])
    media_links = re.findall(r'\[sound:[^\]]+\.mp3\]', html_content)
    
    # Remove all HTML tags
    content = re.sub(r'<[^>]+>', '', html_content)
    
    # Replace HTML entities
    content = content.replace('&nbsp;', ' ')
    content = content.replace('&amp;', '&')
    content = content.replace('&lt;', '<')
    content = content.replace('&gt;', '>')
    content = content.replace('&quot;', '"')
    
    # Clean up whitespace
    content = clean_text(content)
    
    # If we have media links, add them back
    if media_links:
        # If the content is empty, just return the media link
        if not content:
            return media_links[0]
        # Otherwise, append the media link
        return content + ' ' + media_links[0]
    
    return content

def parse_anki_line(line: str) -> list:
    """Parse a single Anki data line, handling multiline HTML content."""
    fields = []
    current_field = ""
    in_quotes = False
    i = 0
    
    while i < len(line):
        char = line[i]
        
        if char == '"':
            in_quotes = not in_quotes
            current_field += char
        elif char == '\t' and not in_quotes:
            fields.append(current_field)
            current_field = ""
        else:
            current_field += char
        
        i += 1
    
    # Add the last field
    fields.append(current_field)
    
    return fields

def is_new_record(line: str) -> bool:
    """Check if a line starts a new record by looking for GUID-like patterns."""
    # First, strip any quotes from the start of the line
    line = line.lstrip('"')
    
    # Skip lines that start with HTML tags
    if line.startswith('<'):
        return False
    
    # Look for lines that start with a GUID-like pattern followed by a tab
    # GUIDs typically contain alphanumeric characters and some special chars
    guid_pattern = r'^[A-Za-z0-9,._\-+=\[\]{}|\\:;"\'<>?/~`!@#$%^&*()]+\t'
    return bool(re.match(guid_pattern, line))

def main():
    # Define required Anki headers
    REQUIRED_HEADERS = [
        '#separator:tab',
        '#html:true',
        '#guid column:1',
        '#notetype column:2',
        '#deck column:3',
        '#tags column:9'
    ]
    
    # —──────────────—
    # 1) Let user pick the input .txt via a file‐open dialog
    # —──────────────—
    root = Tk()
    root.withdraw()  # hide the main tkinter window
    in_path = askopenfilename(
        title="Select your raw Anki export (.txt)",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
    if not in_path:
        print("No file selected. Exiting.", file=sys.stderr)
        sys.exit(1)

    # Build output path by appending "-CLEANED" before the extension
    base, ext = os.path.splitext(in_path)
    out_path = f"{base}-CLEANED{ext}"
    
    # Check if output file already exists
    if os.path.exists(out_path):
        print(f"⚠️  Output file '{out_path}' already exists.")
        response = input("Do you want to overwrite it? (y/N): ").strip().lower()
        if response not in ['y', 'yes']:
            print("Operation cancelled.")
            sys.exit(0)

    # —──────────────—
    # 2) Prepare regex patterns
    # —──────────────—
    sound_re = re.compile(r'\[sound:[^\]]+\.mp3\]')

    entries = []
    headers = []
    current_record = ""

    # —──────────────—
    # 3) Parse the input file
    # —──────────────—
    with open(in_path, encoding='utf-8') as f_in:
        for line in f_in:
            line = line.rstrip('\n')
            
            # Collect header lines (lines starting with #)
            if line.startswith('#'):
                headers.append(line)
                continue
            
            # Skip empty lines
            if not line.strip():
                continue

            # Check if this line starts a new record
            if is_new_record(line):
                # This is a new record, process the previous one if it exists
                if current_record:
                    fields = parse_anki_line(current_record)
                    
                    if len(fields) >= 6:  # Ensure we have enough fields
                        # Extract and clean each field
                        guid = fields[0].strip('"')
                        note_type = fields[1].strip('"')
                        deck = fields[2].strip('"')
                        
                        # Clean the HTML content from the main content fields
                        croatian_html = fields[3] if len(fields) > 3 else ""
                        english_html = fields[4] if len(fields) > 4 else ""
                        audio_html = fields[5] if len(fields) > 5 else ""
                        
                        # Extract clean text from HTML
                        croatian = extract_td_content(croatian_html)
                        english = extract_td_content(english_html)
                        
                        # Extract audio reference
                        audio_match = sound_re.search(audio_html)
                        audio = audio_match.group(0) if audio_match else ""
                        
                        # Get remaining fields (tags, etc.)
                        remaining_fields = fields[6:] if len(fields) > 6 else []
                        
                        # Only keep complete rows with essential data
                        if guid and croatian and english and audio:
                            # Create entry with all original fields, but cleaned content
                            entry = [guid, note_type, deck, croatian, english, audio] + remaining_fields
                            entries.append(entry)
                        else:
                            print(f"Skipping incomplete: GUID={guid!r}, Croatian={croatian!r}, English={english!r}, Audio={audio!r}", file=sys.stderr)
                
                # Start new record
                current_record = line
            else:
                # This is a continuation of the current record
                current_record += "\n" + line

        # Process the last record
        if current_record:
            fields = parse_anki_line(current_record)
            
            if len(fields) >= 6:  # Ensure we have enough fields
                # Extract and clean each field
                guid = fields[0].strip('"')
                note_type = fields[1].strip('"')
                deck = fields[2].strip('"')
                
                # Clean the HTML content from the main content fields
                croatian_html = fields[3] if len(fields) > 3 else ""
                english_html = fields[4] if len(fields) > 4 else ""
                audio_html = fields[5] if len(fields) > 5 else ""
                
                # Extract clean text from HTML
                croatian = extract_td_content(croatian_html)
                english = extract_td_content(english_html)
                
                # Extract audio reference
                audio_match = sound_re.search(audio_html)
                audio = audio_match.group(0) if audio_match else ""
                
                # Get remaining fields (tags, etc.)
                remaining_fields = fields[6:] if len(fields) > 6 else []
                
                # Only keep complete rows with essential data
                if guid and croatian and english and audio:
                    # Create entry with all original fields, but cleaned content
                    entry = [guid, note_type, deck, croatian, english, audio] + remaining_fields
                    entries.append(entry)
                else:
                    print(f"Skipping incomplete: GUID={guid!r}", file=sys.stderr)

    # —──────────────—
    # 4) Write out the cleaned file with all Anki headers
    # —──────────────—
    with open(out_path, 'w', encoding='utf-8-sig', newline='') as f_out:
        # Write headers first
        for header in REQUIRED_HEADERS:
            f_out.write(header + '\n')
        
        # Write data rows manually to avoid CSV escaping issues
        for entry in entries:
            # Join fields with tabs, ensuring no special characters cause issues
            # Make sure we have enough fields (at least 9 for all columns including tags)
            while len(entry) < 9:
                entry.append('')  # Add empty fields if needed
            # Remove any newlines from fields
            entry = [str(field).replace('\n', ' ') for field in entry]
            line = '\t'.join(entry)
            f_out.write(line + '\n')

    print(f"✅ Wrote {len(entries)} entries to '{out_path}'")
    print(f"📋 Preserved {len(headers)} Anki headers")

if __name__ == "__main__":
    main()