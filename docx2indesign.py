#!/usr/bin/env python3
"""
DOCX to InDesign - Document Processor

This script:
1. Converts DOCX to RTF (requires pandoc)
2. Processes RTF to:
   - Fix paragraph breaks
   - Remove unwanted line breaks
   - Add markup for special formatting
3. Outputs a clean file ready for InDesign import

Usage: python docx2indesign.py input.docx [output.txt]
"""

import sys
import os
import re
import subprocess
import tempfile
from pathlib import Path

def convert_docx_to_rtf(docx_path, rtf_path):
    """Convert DOCX to RTF using pandoc."""
    try:
        result = subprocess.run(
            ["pandoc", docx_path, "-o", rtf_path],
            check=True,
            capture_output=True,
            text=True
        )
        print(f"Converted {docx_path} to {rtf_path}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error converting file: {e}")
        print(f"STDOUT: {e.stdout}")
        print(f"STDERR: {e.stderr}")
        return False
    except FileNotFoundError:
        print("Error: pandoc not found. Please install pandoc first.")
        print("  macOS: brew install pandoc")
        print("  Linux: apt-get install pandoc")
        return False

def process_rtf(rtf_path, output_path):
    """Process RTF file to clean up line breaks and add markup."""
    # Read the RTF file
    with open(rtf_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
    
    # Extract pure text content from RTF (simplified approach)
    # This is a very basic approach - a proper RTF parser would be better
    text_content = extract_text_from_rtf(content)
    
    # Process the text content
    processed_text = process_text(text_content)
    
    # Write processed content to output file
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(processed_text)
    
    print(f"Processed RTF and saved to {output_path}")

def extract_text_from_rtf(rtf_content):
    """
    Very basic RTF text extraction.
    For a real implementation, use a proper RTF parser.
    """
    # Remove RTF control sequences
    text = re.sub(r'\\[a-z]+[0-9-]*', ' ', rtf_content)
    text = re.sub(r'\{|\}|\\', '', text)
    
    # Handle special characters
    text = re.sub(r"\'[0-9a-fA-F]{2}", "", text)
    
    return text

def process_text(text):
    """Process the extracted text content."""
    # Identify paragraph breaks (blank lines)
    paragraphs = re.split(r'\n\s*\n', text)
    
    processed_paragraphs = []
    for para in paragraphs:
        # Remove single line breaks within paragraphs
        clean_para = re.sub(r'(?<!\n)\n(?!\n)', ' ', para.strip())
        # Fix multiple spaces
        clean_para = re.sub(r' +', ' ', clean_para)
        processed_paragraphs.append(clean_para)
    
    # Join paragraphs with double line breaks
    processed_text = '\n\n'.join(processed_paragraphs)
    
    # Add formatting summary
    # Note: In this basic version, we don't detect actual formatting,
    # so we add a simplified summary
    formatting_summary = generate_formatting_summary()
    processed_text += "\n\n" + formatting_summary
    
    return processed_text

def generate_formatting_summary():
    """Generate a simplified formatting summary."""
    summary = []
    summary.append("====== FORMATTING SUMMARY ======")
    summary.append("Note: This basic version doesn't detect formatting.")
    summary.append("For better formatting detection, use the advanced version:")
    summary.append("  python docx2indesign_advanced.py input.docx")
    summary.append("")
    summary.append("Common formatting in documents (not detected):")
    summary.append("  * Paragraphs (separated by double newlines)")
    summary.append("  * Bold (**text**)")
    summary.append("  * Italic (_text_)")
    summary.append("  * Bold+Italic (***text***)")
    summary.append("")
    summary.append("You can use GREP in InDesign to find and format these elements.")
    
    return "\n".join(summary)

def create_clean_version(input_path, output_path=None):
    """Main function to process a DOCX file to clean text."""
    input_path = Path(input_path)
    
    # Set default output path if not provided
    if not output_path:
        output_path = input_path.with_suffix('.clean.txt')
    else:
        output_path = Path(output_path)
    
    # Create temp RTF file
    with tempfile.NamedTemporaryFile(suffix='.rtf', delete=False) as temp:
        temp_rtf_path = temp.name
    
    try:
        # Step 1: Convert DOCX to RTF
        if input_path.suffix.lower() == '.docx':
            if not convert_docx_to_rtf(input_path, temp_rtf_path):
                print("Conversion failed.")
                return False
        elif input_path.suffix.lower() == '.rtf':
            # If input is already RTF, just copy it
            with open(input_path, 'rb') as src, open(temp_rtf_path, 'wb') as dst:
                dst.write(src.read())
        else:
            print(f"Unsupported file format: {input_path.suffix}")
            return False
        
        # Step 2: Process RTF
        process_rtf(temp_rtf_path, output_path)
        return True
    
    finally:
        # Clean up temp file
        if os.path.exists(temp_rtf_path):
            os.unlink(temp_rtf_path)

def main():
    if len(sys.argv) < 2:
        print("Usage: python docx2indesign.py input.docx [output.txt]")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    if create_clean_version(input_path, output_path):
        print("Processing completed successfully!")
    else:
        print("Processing failed.")
        sys.exit(1)

if __name__ == "__main__":
    main()