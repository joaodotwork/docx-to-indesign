# DOCX to InDesign Converter

This tool helps prepare DOCX files for clean import into InDesign by:

1. Converting DOCX to an intermediate format (HTML)
2. Cleaning up paragraph breaks and line breaks
3. Adding markup for formatting that's compatible with InDesign GREP

## Repository Contents

- `docx2indesign.py` — basic converter (paragraph/line break cleanup)
- `docx2indesign_advanced.py` — advanced converter with formatting markup and a batch UI
- `endnotes_scripts/` — InDesign JSX and AppleScript scripts for converting endnote markers into real InDesign endnotes
- `indesign-template/docx2indd-template.indt` — starter InDesign template with paragraph/character styles matching the GREP workflow below
- `requirements.txt` — Python dependencies

> **Note:** This repository ships the tooling only. Source `.docx` files and the processed output they generate are not included.

## Prerequisites

- Python 3.6 or higher
- Pandoc (for document conversion)
- BeautifulSoup4 (for the advanced version)

### Installation

1. Install Pandoc:
   ```
   # macOS
   brew install pandoc
   
   # Ubuntu/Debian
   sudo apt-get install pandoc
   ```

2. Install Python dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

### Basic Version

The basic version focuses on fixing paragraph breaks and line breaks:

```
python docx2indesign.py input.docx [output.txt]
```

If no output file is specified, it will create `input.clean.txt` in the same directory.

### Advanced Version (Recommended)

The advanced version has a colorful terminal UI, supports multiple files/directories, and preserves formatting:

```
python docx2indesign_advanced.py input.docx [input2.docx ...] [-o OUTPUT_DIR] [-r]
```

Options:
- `-o, --output`: Specify output directory (default: ./output)
- `-r, --recursive`: Process directories recursively

Examples:
```bash
# Process a single file
python docx2indesign_advanced.py my_document.docx

# Process multiple files with custom output directory
python docx2indesign_advanced.py doc1.docx doc2.docx -o ./processed

# Process all docx files in a directory
python docx2indesign_advanced.py ./documents/

# Process all docx files recursively
python docx2indesign_advanced.py ./documents/ -r
```

## Formatting Markers

The advanced script adds the following markup that can be processed with InDesign GREP:

- **Bold**: `**text**`
- *Italic*: `_text_`
- ***Bold Italic***: `***text***`
- __Underline__: `__text__`
- Superscript: `^(text)`
- Subscript: `~(text)`
- Headings: `# Heading 1`, `## Heading 2`, etc.
- Footnotes: `[^1]` with footnote text at the end of the document
- Line breaks: `\n` (escaped newline)
- Links: `[text](url)`
- Lists:
  - Bulleted: `* item`
  - Numbered: `1. item`

## InDesign GREP Search/Replace

The GREP steps below assume your document already has the matching character and paragraph styles (Bold, Italic, Heading 1, etc.). The included `indesign-template/docx2indd-template.indt` provides these styles as a starting point — open it (or load its styles into your own document) before running the search/replace steps.

After importing the processed text into InDesign (using Unicode UTF-8 encoding), use GREP search/replace to convert the markup to proper formatting:

1. Bold: 
   - Search for: `\*\*([^*]+)\*\*` 
   - Change to: `$1` (with Bold character style applied)

2. Italic: 
   - Search for: `_([^_]+)_` 
   - Change to: `$1` (with Italic character style applied)

3. Bold Italic:
   - Search for: `\*\*\*([^*]+)\*\*\*`
   - Change to: `$1` (with Bold Italic character style applied)

4. Underline: 
   - Search for: `__(.+?)__` 
   - Change to: `$1` (with Underline character style applied)
   
   Alternative patterns if the above doesn't work:
   - `__([^_]*)__` (more restrictive match)
   - `__([^*]*)__` (avoids conflicts with bold asterisks)

5. Superscript:
   - Search for: `\^\(([^)]+)\)` 
   - Change to: `$1` (with Superscript character style applied)

6. Subscript:
   - Search for: `~\(([^)]+)\)` 
   - Change to: `$1` (with Subscript character style applied)

7. Headings:
   - Heading 1: Search for `^#\s+(.+)$` - Change to: `$1` (with Heading 1 paragraph style)
   - Heading 2: Search for `^##\s+(.+)$` - Change to: `$1` (with Heading 2 paragraph style)
   - Heading 3: Search for `^###\s+(.+)$` - Change to: `$1` (with Heading 3 paragraph style)
   - Heading 4: Search for `^####\s+(.+)$` - Change to: `$1` (with Heading 4 paragraph style)

8. Endnotes:
   - Standard format: Search for: `\[\^(\d+)\]` 
   - Change to: Character style for endnote references (with custom superscript formatting)
   - Pandoc format with links: Search for: `\[\^\((\d+)\)\]\(#fn\d+\)`
   - Change to: `$1` (with superscript character style applied)
   - **For actual InDesign endnotes**: Use the included script `endnotes_scripts/docx_to_indesign_endnotes.jsx`

9. Links:
   - Search for: `\[([^\]]+)\]\(([^)]+)\)` 
   - Change to: `$1` (with Hyperlink character style applied and URL set to $2)

10. Lists:
   - Bulleted: Search for `^\*\s+(.+)$` - Change to: `$1` (with Bulleted List paragraph style)
   - Numbered: Search for `^\d+\.\s+(.+)$` - Change to: `$1` (with Numbered List paragraph style)

### GREP Troubleshooting

If you have trouble with the GREP patterns:

1. Make sure "Case Sensitive" is checked if appropriate
2. Try the alternative patterns provided for underline formatting
3. For underline formatting, be careful with the pattern to avoid conflicts with other formatting
4. For patterns not working, try testing with a simpler version first
5. For multiline elements, make sure "Span Lines" is selected
6. Special characters may need more escaping in InDesign than shown here
7. Try using InDesign's GREP Editor to build and test your patterns
8. After importing, view the document in Story Editor (Edit > Edit in Story Editor) to see the raw markup

## Features

- Colorful terminal UI with progress indicators
- Process multiple files in batch
- Directory recursive scanning
- Detailed processing logs
- Custom output directory

## Limitations

- Complex document structures like tables may not be preserved
- Some special formatting might be lost in the conversion process
- The scripts provide a starting point that you may need to customize for your specific documents

## Using the Endnote Script

The `endnotes_scripts/docx_to_indesign_endnotes.jsx` script converts the endnote markers from the Python converter into proper InDesign endnotes:

1. Process your Word document with `docx2indesign_advanced.py` as usual
2. Import the resulting text file into InDesign
3. Run the EndnoteConverter script in InDesign:
   - Go to File > Scripts > Script Panel
   - Navigate to the `endnotes_scripts/docx_to_indesign_endnotes.jsx` file and double-click it
4. The script will:
   - Find all endnote references in the format `[^(1)](#fn1)`
   - Find the matching endnote text at the end of the document
   - Convert them to actual InDesign endnotes
   - Remove the original endnote text section at the end

After running the script, you can format the endnotes using InDesign's built-in endnote formatting options.

### Troubleshooting Endnotes

If the script doesn't convert your endnotes:

1. Check that your endnote references match the format: `[^(1)](#fn1)` 
2. Confirm the document has endnote content near the end, typically after a horizontal rule (---)
3. Make sure each endnote content section starts with `[#fn1]` and ends with `[↩︎](#fnref1)`
4. Try running the script with a small test document first
5. If you continue having issues, you may need to adjust the regex patterns in the script to match your specific document format

## License

Released under the [MIT License](LICENSE).