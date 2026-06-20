# DOCX to InDesign Project Guidelines

## Commands
- Run basic converter: `python docx2indesign.py input.docx [output.txt]`
- Run advanced converter: `python docx2indesign_advanced.py input.docx [input2.docx ...] [-o OUTPUT_DIR] [-r]`
- Normalize styles across files: `python normalize_docx_styles.py target.docx ... --reference ref.docx -o OUTPUT_DIR [--dry-run]`
- Install dependencies: `pip install -r requirements.txt`
- Install pandoc (macOS): `brew install pandoc`
- Install pandoc (Linux): `sudo apt-get install pandoc`

## Project Structure
- **Python Scripts**:
  - `docx2indesign.py` - Basic converter focusing on paragraph/line breaks
  - `docx2indesign_advanced.py` - Advanced version with formatting preservation, colorful UI
  - `docx_notes.py` - Reads footnote/endnote structure directly from the DOCX XML (preserves the kind distinction pandoc discards); used by the advanced converter
  - `normalize_docx_styles.py` - Copies style definitions from a reference DOCX into others to unify styling without touching text (swaps only word/styles.xml, asserts content parts byte-identical)
- **InDesign Scripts** (in `endnotes_scripts/` folder):
  - `notes_converter.jsx` - Converts both footnote (`[^F(n)]`) and endnote (`[^E(n)]`) markers into native InDesign notes
  - `batch_grep_format.jsx` - Runs every formatting GREP (bold/italic/headings/lists/links) in one pass, in dependency order; style names match the template (`H1`-`H4`, `Bold`, etc.)
  - `docx_to_indesign_endnotes.jsx` - Legacy endnote conversion script (`[^(n)](#fnn)` markers)
- **Saved Find/Change Queries** (`indesign-queries/GREP/`): one XML per conversion, `docx2indd - *`, for stepping through changes interactively. Format mirrors InDesign's own query files; `cstyle`/`pstyle` TextAttribute values must match real style names in the doc.
  - `endnote_converter_2025.jsx` - Updated converter for InDesign 2025
  - `endnote_improved_v2.jsx` - Improved version with better error handling
  - `endnotes_converter_improved.jsx` - Enhanced converter with better footnote detection
  - `endnotes_step1_find_references.jsx` - First step in multi-step conversion process
  - `endnote_capabilities.jsx` - Script to test InDesign endnote capabilities
  - `endnote_direct_test.jsx` - Test script for direct endnote conversion
  - `endnotes_converter.applescript` - AppleScript version for older InDesign versions
- **Output Directory**: Processed files are saved to the output directory by default

## Workflow
1. Process DOCX files with Python scripts to create clean text with markup
2. Import the processed text files into InDesign
3. Use InDesign GREP to convert markup to proper formatting
4. Run endnote conversion scripts in InDesign to handle endnotes

## Code Style Guidelines
- **Imports**: Group standard library imports first, followed by third-party imports, then local imports
- **Formatting**: Use 4 spaces for indentation; maintain 79-character line length
- **Naming**: Use snake_case for variables/functions, CamelCase for classes
- **Error Handling**: Use try/except blocks with specific exception classes
- **Documentation**: Include docstrings for functions and modules; follow Google docstring format
- **Type Handling**: Use Python type hints where appropriate
- **Functions**: Keep functions focused on a single responsibility
- **Logging**: Use print_progress() in advanced version for consistent messaging
- **Constants**: Define constants at module level using UPPERCASE
- **File Handling**: Use context managers (with) for file operations
- **Path Handling**: Use pathlib.Path for filesystem operations