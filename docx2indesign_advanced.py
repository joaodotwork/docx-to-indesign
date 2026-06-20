#!/usr/bin/env python3
"""
DOCX to InDesign - Advanced Document Processor

This script:
1. Converts DOCX to HTML (requires pandoc)
2. Processes the file to:
   - Fix paragraph breaks
   - Remove unwanted line breaks
   - Preserve formatting with markup tags for InDesign
3. Outputs a clean file ready for InDesign import with GREP-friendly markup

Usage: python docx2indesign_advanced.py input.docx [output.txt]
"""

import sys
import os
import re
import glob
import subprocess
import tempfile
import argparse
from pathlib import Path
from bs4 import BeautifulSoup
from docx_notes import extract_notes
import shutil
from time import sleep
from datetime import datetime

# Terminal colors and formatting
class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    END = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def print_banner():
    """Print a simplified colorful banner."""
    banner = f"""
{Colors.CYAN}{Colors.BOLD}=== {Colors.GREEN}DOCX to InDesign{Colors.CYAN} - Advanced Document Processor ===
{Colors.BLUE}Convert DOCX files to InDesign-friendly format{Colors.END}
"""
    print(banner)

def print_progress(message, is_success=True):
    """Print a progress message with appropriate formatting."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    if is_success:
        print(f"{Colors.BLUE}[{timestamp}]{Colors.END} {Colors.GREEN}✓{Colors.END} {message}")
    else:
        print(f"{Colors.BLUE}[{timestamp}]{Colors.END} {Colors.FAIL}✗{Colors.END} {message}")

def print_processing(file_index, total_files, filename):
    """Print processing information."""
    print(f"{Colors.BLUE}[{file_index}/{total_files}]{Colors.END} Processing: {Colors.CYAN}{filename}{Colors.END}")

def convert_docx_to_html(docx_path, html_path):
    """Convert DOCX to HTML using pandoc to preserve formatting."""
    try:
        result = subprocess.run(
            ["pandoc", docx_path, "-o", html_path, "--wrap=none"],
            check=True,
            capture_output=True,
            text=True
        )
        print_progress(f"Converted {Path(docx_path).name} to HTML")
        return True
    except subprocess.CalledProcessError as e:
        print_progress(f"Error converting file: {e}", False)
        print(f"{Colors.FAIL}STDOUT: {e.stdout}{Colors.END}")
        print(f"{Colors.FAIL}STDERR: {e.stderr}{Colors.END}")
        return False
    except FileNotFoundError:
        print_progress("Error: pandoc not found. Please install pandoc first.", False)
        print(f"  {Colors.BLUE}macOS:{Colors.END} brew install pandoc")
        print(f"  {Colors.BLUE}Linux:{Colors.END} apt-get install pandoc")
        return False

def process_html(html_path, output_path, notes=None):
    """Process HTML file to clean up line breaks and preserve formatting with markup.

    Args:
        html_path: Path to the intermediate HTML produced by pandoc.
        output_path: Where to write the InDesign-friendly text.
        notes: Optional list of note dicts from docx_notes.extract_notes(),
            used to distinguish footnotes from endnotes (pandoc cannot).
    """
    # Read the HTML file
    with open(html_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    # Parse HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # Process the content
    processed_text = extract_formatted_text(soup, notes)
    
    # Write processed content to output file
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(processed_text)
    
    print_progress(f"Saved clean file to {Path(output_path).name}")

# Pandoc emits both footnote and endnote references as "[^(N)](#fnN)" with a
# single sequential numbering. This pattern matches one such marker.
NOTE_MARKER_RE = re.compile(r"\[\^\((\d+)\)\]\(#fn\d+\)")

# Per-kind marker prefix used in the rewritten output and matched by the
# InDesign converter: footnotes -> "F", endnotes -> "E".
NOTE_KIND_PREFIX = {"footnote": "F", "endnote": "E"}


def rewrite_note_markers(text, notes):
    """Rewrite pandoc's generic note markers into type-specific ones.

    "[^(2)](#fn2)" becomes "[^F(2)]" or "[^E(2)]" depending on whether note 2
    is a footnote or an endnote (per the DOCX). The number is pandoc's
    sequential note number, which equals the 1-based index in ``notes``.
    """
    by_number = {note["n"]: note for note in notes}

    def repl(match):
        number = int(match.group(1))
        note = by_number.get(number)
        if not note:
            # No mapping (shouldn't happen): leave the original marker intact.
            return match.group(0)
        prefix = NOTE_KIND_PREFIX[note["kind"]]
        return f"[^{prefix}({number})]"

    return NOTE_MARKER_RE.sub(repl, text)


def build_note_sections(notes, formatting_used=None):
    """Build the trailing FOOTNOTES / ENDNOTES blocks from extracted notes.

    Each body is emitted as "[^F<n>]: <text>" (or "[^E<n>]: ...") on its own
    line, keyed by the same number as the in-text marker so the InDesign
    converter can pair them. Returns an empty string when there are no notes.
    """
    footnotes = [n for n in notes if n["kind"] == "footnote"]
    endnotes = [n for n in notes if n["kind"] == "endnote"]

    sections = ""
    if footnotes:
        if formatting_used is not None:
            formatting_used["footnotes"] = True
        sections += "\n\n====== FOOTNOTES ======\n"
        for note in footnotes:
            sections += f"[^F{note['n']}]: {note['text']}\n"
    if endnotes:
        if formatting_used is not None:
            formatting_used["endnotes"] = True
        sections += "\n\n====== ENDNOTES ======\n"
        for note in endnotes:
            sections += f"[^E{note['n']}]: {note['text']}\n"
    return sections


def extract_formatted_text(soup, notes=None):
    """Extract text with formatting from HTML.

    Args:
        soup: Parsed pandoc HTML.
        notes: Optional list of note dicts from docx_notes.extract_notes().
            When provided, the in-text note markers are rewritten to
            type-specific markers and clean footnote/endnote sections are
            appended (pandoc itself cannot tell the two kinds apart).
    """
    result = ""

    # Track formatting elements used in the document
    formatting_used = {
        "headings": set(),
        "bold": False,
        "italic": False,
        "bold_italic": False,
        "underline": False,
        "superscript": False,
        "subscript": False,
        "links": False,
        "bullet_lists": False,
        "numbered_lists": False,
        "footnotes": False,
        "endnotes": False
    }

    # When we have authoritative note data from the DOCX, drop pandoc's
    # auto-generated notes section. It both lacks the footnote/endnote
    # distinction and (because its <li> bodies are also picked up by the
    # paragraph/list loops below) renders every note twice. We rebuild clean,
    # type-separated sections from the DOCX XML instead.
    if notes:
        note_section = soup.find('section', class_='footnotes')
        if note_section:
            note_section.decompose()

    # Process paragraphs
    for para in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
        # Determine if it's a heading
        if para.name.startswith('h'):
            heading_level = int(para.name[1])
            formatting_used["headings"].add(heading_level)
            para_text = f"{'#' * heading_level} {process_inline_elements(para, formatting_used)}\n\n"
        else:
            para_text = f"{process_inline_elements(para, formatting_used)}\n\n"
        
        result += para_text
    
    # Process lists
    for list_elem in soup.find_all(['ul', 'ol']):
        for item in list_elem.find_all('li'):
            if list_elem.name == 'ul':
                formatting_used["bullet_lists"] = True
                result += f"* {process_inline_elements(item, formatting_used)}\n"
            else:
                formatting_used["numbered_lists"] = True
                result += f"1. {process_inline_elements(item, formatting_used)}\n"
        result += "\n"
    
    # Notes: rewrite the generic in-text markers to footnote/endnote-specific
    # ones and append clean, type-separated note sections built from the DOCX.
    if notes:
        result = rewrite_note_markers(result, notes)
        result += build_note_sections(notes, formatting_used)

    # Add formatting summary at the end
    result += "\n\n" + generate_formatting_summary(formatting_used)
    
    return result

def process_inline_elements(element, formatting_used=None, depth=0):
    """Process inline elements to add markup.
    
    Args:
        element: BeautifulSoup element to process
        formatting_used: Dictionary tracking formatting used in document
        depth: Current recursion depth to prevent infinite recursion
    """
    # Prevent infinite recursion
    if depth > 20:  # Limiting recursion depth
        return "[NESTED CONTENT TOO DEEP]"
    
    text = ""
    
    # Check if element has both bold and italic without recursion
    is_bold_italic = False
    try:
        if ((element.name in ['em', 'i'] and 
             (element.find_all(['strong', 'b'], recursive=False) or 
              any(parent.name in ['strong', 'b'] for parent in list(element.parents)[:3]))) or
            (element.name in ['strong', 'b'] and 
             (element.find_all(['em', 'i'], recursive=False) or
              any(parent.name in ['em', 'i'] for parent in list(element.parents)[:3])))):
            is_bold_italic = True
            if formatting_used:
                formatting_used["bold_italic"] = True
            # Process children directly without recursing on the same element
            inner_text = ""
            for child in element.children:
                if child.name is None:  # Text node
                    inner_text += child.string if child.string else ""
                else:
                    # Process other nested elements, but skip formatting we already captured
                    inner_text += process_inline_elements(child, formatting_used, depth + 1)
            text += f"***{inner_text}***"
            return text
    except RecursionError:
        if formatting_used:
            formatting_used["bold_italic"] = True
        return "***[NESTED FORMATTING ERROR]***"
    
    try:
        for child in element.children:
            if child.name is None:  # Text node
                text += child.string if child.string else ""
            elif child.name == 'em' or child.name == 'i':
                # Check if this italic element contains or is inside a bold element
                has_bold = False
                try:
                    # Limit search depth and use non-recursive find where possible
                    has_bold = (child.find_all(['strong', 'b'], recursive=False) or 
                               any(parent.name in ['strong', 'b'] for parent in list(child.parents)[:3]))
                except RecursionError:
                    has_bold = False
                
                if has_bold:
                    if formatting_used:
                        formatting_used["bold_italic"] = True
                    text += f"***{process_inline_elements(child, formatting_used, depth + 1)}***"
                else:
                    if formatting_used:
                        formatting_used["italic"] = True
                    text += f"_{process_inline_elements(child, formatting_used, depth + 1)}_"
            elif child.name == 'strong' or child.name == 'b':
                # Check if this bold element contains or is inside an italic element
                has_italic = False
                try:
                    # Limit search depth and use non-recursive find where possible
                    has_italic = (child.find_all(['em', 'i'], recursive=False) or
                                 any(parent.name in ['em', 'i'] for parent in list(child.parents)[:3]))
                except RecursionError:
                    has_italic = False
                
                if has_italic:
                    if formatting_used:
                        formatting_used["bold_italic"] = True
                    text += f"***{process_inline_elements(child, formatting_used, depth + 1)}***"
                else:
                    if formatting_used:
                        formatting_used["bold"] = True
                    text += f"**{process_inline_elements(child, formatting_used, depth + 1)}**"
            elif child.name == 'u':  # Underline
                if formatting_used:
                    formatting_used["underline"] = True
                text += f"__{process_inline_elements(child, formatting_used, depth + 1)}__"
            elif child.name == 'sup':  # Superscript
                if formatting_used:
                    formatting_used["superscript"] = True
                text += f"^({process_inline_elements(child, formatting_used, depth + 1)})"
            elif child.name == 'sub':  # Subscript
                if formatting_used:
                    formatting_used["subscript"] = True
                text += f"~({process_inline_elements(child, formatting_used, depth + 1)})"
            elif child.name == 'a':  # Links
                if formatting_used:
                    formatting_used["links"] = True
                href = child.get('href', '')
                text += f"[{process_inline_elements(child, formatting_used, depth + 1)}]({href})"
            elif child.name == 'br':  # Line break
                text += "\\n"  # Escaped newline marker for GREP
            else:
                text += process_inline_elements(child, formatting_used, depth + 1)
    except RecursionError:
        return "[RECURSION ERROR IN DOCUMENT]"
    
    return text

def generate_formatting_summary(formatting_used):
    """Generate a summary of all formatting used in the document."""
    summary = []
    summary.append("====== FORMATTING SUMMARY ======")
    summary.append("The following formatting elements were detected in this document:")
    summary.append("")
    
    # Headers
    if formatting_used["headings"]:
        heading_levels = sorted(formatting_used["headings"])
        summary.append("HEADINGS:")
        for level in heading_levels:
            summary.append(f"  * Level {level} (#{'#' * level})")
    
    # Text styling
    text_styling = []
    if formatting_used["bold"]:
        text_styling.append("  * Bold (**text**)")
    if formatting_used["italic"]:
        text_styling.append("  * Italic (_text_)")
    if formatting_used["bold_italic"]:
        text_styling.append("  * Bold+Italic (***text***)")
    if formatting_used["underline"]:
        text_styling.append("  * Underline (__text__)")
    
    if text_styling:
        summary.append("TEXT STYLING:")
        summary.extend(text_styling)
    
    # Special formatting
    special_formatting = []
    if formatting_used["superscript"]:
        special_formatting.append("  * Superscript (^(text))")
    if formatting_used["subscript"]:
        special_formatting.append("  * Subscript (~(text))")
    if formatting_used["links"]:
        special_formatting.append("  * Links ([text](url))")
    
    if special_formatting:
        summary.append("SPECIAL FORMATTING:")
        summary.extend(special_formatting)
    
    # Lists and footnotes
    other_elements = []
    if formatting_used["bullet_lists"]:
        other_elements.append("  * Bullet Lists (* item)")
    if formatting_used["numbered_lists"]:
        other_elements.append("  * Numbered Lists (1. item)")
    if formatting_used["footnotes"]:
        other_elements.append("  * Footnotes (ref [^F(n)], body [^Fn]: text)")
    if formatting_used.get("endnotes"):
        other_elements.append("  * Endnotes (ref [^E(n)], body [^En]: text)")
    
    if other_elements:
        summary.append("OTHER ELEMENTS:")
        summary.extend(other_elements)
    
    summary.append("")
    summary.append("You can use GREP in InDesign to find and format these elements.")
    
    return "\n".join(summary)

def create_clean_version(input_path, output_dir):
    """Process a DOCX file to clean text with formatting."""
    input_path = Path(input_path)
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Define output file path
    output_path = Path(output_dir) / f"{input_path.stem}.clean.txt"
    
    # Create temp HTML file
    with tempfile.NamedTemporaryFile(suffix='.html', delete=False) as temp:
        temp_html_path = temp.name
    
    # Notes are read straight from the DOCX so footnotes and endnotes stay
    # distinguishable (pandoc merges them). Only DOCX carries this structure.
    notes = []

    try:
        # Step 1: Convert DOCX to HTML
        if input_path.suffix.lower() == '.docx':
            if not convert_docx_to_html(input_path, temp_html_path):
                return False
            try:
                notes = extract_notes(input_path)
                if notes:
                    fn = sum(1 for n in notes if n["kind"] == "footnote")
                    en = len(notes) - fn
                    print_progress(
                        f"Detected {fn} footnote(s) and {en} endnote(s)"
                    )
            except Exception as e:
                print_progress(f"Could not read notes from DOCX: {e}", False)
        elif input_path.suffix.lower() == '.html':
            # If input is already HTML, just copy it
            with open(input_path, 'rb') as src, open(temp_html_path, 'wb') as dst:
                dst.write(src.read())
        elif input_path.suffix.lower() == '.rtf':
            # For RTF, convert to HTML first
            if not convert_docx_to_html(input_path, temp_html_path):
                return False
        else:
            print_progress(f"Unsupported file format: {input_path.suffix}", False)
            return False
        
        # Step 2: Process HTML to extract formatted text
        process_html(temp_html_path, output_path, notes)
        return True
    
    finally:
        # Clean up temp file
        if os.path.exists(temp_html_path):
            os.unlink(temp_html_path)

def process_files(input_paths, output_dir):
    """Process multiple files."""
    total_files = len(input_paths)
    successful = 0
    failed = 0
    
    print(f"\n{Colors.BOLD}Found {total_files} file(s) to process{Colors.END}\n")
    
    for i, input_path in enumerate(input_paths, 1):
        file_path = Path(input_path)
        print_processing(i, total_files, file_path.name)
        
        if create_clean_version(input_path, output_dir):
            successful += 1
        else:
            failed += 1
    
    # Print summary
    print(f"\n{Colors.BOLD}Processing complete!{Colors.END}")
    print(f"{Colors.GREEN}✓ Successfully processed: {successful}{Colors.END}")
    if failed > 0:
        print(f"{Colors.FAIL}✗ Failed: {failed}{Colors.END}")
    
    if successful > 0:
        print(f"\n{Colors.BOLD}Output files are in:{Colors.END} {Colors.UNDERLINE}{os.path.abspath(output_dir)}{Colors.END}")
    
    return successful > 0

def main():
    parser = argparse.ArgumentParser(description='Convert DOCX files to InDesign-friendly format')
    parser.add_argument('input', nargs='+', help='Input file(s) or directory')
    parser.add_argument('-o', '--output', default='./output', help='Output directory')
    parser.add_argument('-r', '--recursive', action='store_true', help='Recursively process directories')
    args = parser.parse_args()
    
    print_banner()
    
    # Gather all input files
    input_files = []
    for input_path in args.input:
        path = Path(input_path)
        if path.is_file():
            if path.suffix.lower() in ['.docx', '.rtf', '.html']:
                input_files.append(str(path))
            else:
                print_progress(f"Skipping unsupported file: {path.name}", False)
        elif path.is_dir():
            if args.recursive:
                for ext in ['.docx', '.rtf', '.html']:
                    input_files.extend(glob.glob(str(path / f'**/*{ext}'), recursive=True))
            else:
                for ext in ['.docx', '.rtf', '.html']:
                    input_files.extend(glob.glob(str(path / f'*{ext}')))
        else:
            print_progress(f"Path not found: {input_path}", False)
    
    if not input_files:
        print_progress("No valid input files found.", False)
        sys.exit(1)
    
    # Process all files
    success = process_files(input_files, args.output)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{Colors.WARNING}Processing cancelled by user.{Colors.END}")
        sys.exit(130)