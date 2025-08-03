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
    """Extract text content from HTML td tags, handling both single and multiple td elements."""
    # Remove quotes if present
    html_content = html_content.strip('"')
    
    # Pattern to match td content, handling multiline and nested content
    td_pattern = re.compile(r'<td[^>]*>(.*?)</td>', re.DOTALL)
    matches = td_pattern.findall(html_content)
    
    if matches:
        # Join multiple td contents with space if there are multiple
        content = ' '.join(matches)
        return clean_text(content)
    else:
        # If no td tags found, return the cleaned content as-is
        return clean_text(html_content)

def main():
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

    # —──────────────—
    # 2) Prepare regex patterns
    # —──────────────—
    sound_re = re.compile(r'\[sound:[^\]]+\.mp3\]')

    entries = []
    buffer = []
    headers = []

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
                
            buffer.append(line)

            # Process complete records when we find a sound tag
            if sound_re.search(line):
                record = "\n".join(buffer)
                buffer.clear()

                # Split the record by tabs
                fields = record.split('\t')
                
                if len(fields) >= 9:  # Ensure we have enough fields
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
        writer = csv.writer(
            f_out,
            delimiter='\t',
            quoting=csv.QUOTE_NONE,
            escapechar='\\'
        )

        # Write all the original Anki headers
        for header in headers:
            writer.writerow([header])

        # Optional header row of field names (commented out by default)
        # writer.writerow(['GUID', 'NoteType', 'Deck', 'Croatian', 'English', 'Audio', 'Remaining...'])

        # Data rows
        writer.writerows(entries)

    print(f"✅ Wrote {len(entries)} entries to '{out_path}'")
    print(f"📋 Preserved {len(headers)} Anki headers")

if __name__ == "__main__":
    main()