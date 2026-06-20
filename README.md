# DOCX to InDesign Converter

This tool helps prepare DOCX files for clean import into InDesign by:

1. Converting DOCX to an intermediate format (HTML)
2. Cleaning up paragraph breaks and line breaks
3. Adding markup for formatting that's compatible with InDesign GREP

## Repository Contents

- `docx2indesign.py` — basic converter (paragraph/line break cleanup)
- `docx2indesign_advanced.py` — advanced converter with formatting markup and a batch UI
- `docx_notes.py` — reads footnote/endnote structure directly from the DOCX so the two kinds stay distinct (pandoc merges them)
- `normalize_docx_styles.py` — makes paragraph/character style *definitions* consistent across a set of DOCX files without altering any text
- `endnotes_scripts/` — InDesign JSX and AppleScript scripts: `notes_converter.jsx` (footnotes + endnotes), `batch_grep_format.jsx` (runs every formatting GREP in one pass)
- `indesign-queries/GREP/` — saved Find/Change queries you can load one at a time from the Find/Change dialog
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
- Footnotes: in-text reference `[^F(n)]`, body `[^Fn]: text` under a `====== FOOTNOTES ======` heading
- Endnotes: in-text reference `[^E(n)]`, body `[^En]: text` under a `====== ENDNOTES ======` heading
- Block quotes: `> text` (from Word's Block Quote / Quote styles)
- Bibliography: `[BIB] text` (from Word's Bibliography style, e.g. Zotero reference lists)
- Line breaks: `\n` (escaped newline)
- Links: `[text](url)`
- Lists:
  - Bulleted: `* item`
  - Numbered: `1. item`

## Applying the Markup in InDesign

You have three ways to turn the markup into formatting, from most interactive to fully automatic:

1. **Saved Find/Change queries** (`indesign-queries/GREP/`) — step through each conversion in the Find/Change dialog and watch the result before moving on. Best when you want to follow the changes after import.
2. **Batch script** (`endnotes_scripts/batch_grep_format.jsx`) — runs every formatting GREP in one pass, in the correct order. Best once you trust the conversions.
3. **Manual GREP** — type each pattern by hand (reference table below).

All three need your document to already have the matching styles. The included `indesign-template/docx2indd-template.indt` provides them (`Bold`, `Italic`, `Bold Italic`, `Underline`, `Superscript`, `Subscript`, `H1`–`H4`, `Bulleted List`, `Numbered List`, `Hyperlink`) — open it or load its styles before running any of the steps.

### Saved Find/Change Queries

Copy the query files into your InDesign Find/Change queries folder, then restart InDesign:

```bash
# macOS — adjust the version number to match your install
cp "indesign-queries/GREP/"*.xml \
  "$HOME/Library/Preferences/Adobe InDesign/Version 21.0/en_US/Find-Change Queries/GREP/"
```

They then appear in **Edit > Find/Change > GREP tab > Query** dropdown, prefixed `docx2indd -`. Run them **in this order** so the patterns don't clobber each other (triple markers before double, deepest heading first):

1. `docx2indd - Bold Italic` (`***`)
2. `docx2indd - Bold` (`**`)
3. `docx2indd - Underline` (`__`)
4. `docx2indd - Italic` (`_`)
5. `docx2indd - Superscript`
6. `docx2indd - Subscript`
7. `docx2indd - Heading 1` … `Heading 4`
8. `docx2indd - Bulleted List`, `docx2indd - Numbered List`
9. `docx2indd - Block Quote`, `docx2indd - Bibliography`
10. `docx2indd - Links`

The Block Quote and Bibliography queries (and the `docx+styles`-based markup that feeds them) require the matching `Block Quote` and `Bibliography` paragraph styles in your document. Run `endnotes_scripts/ensure_template_styles.jsx` once with the template open to create any styles the queries expect but that are missing (existing styles are left untouched), then File > Save to update the template.

Footnotes and endnotes are **not** handled by a query — their bodies have to be moved into the note itself, so use `endnotes_scripts/notes_converter.jsx` for those (see [Footnotes and Endnotes](#footnotes-and-endnotes)).

### Manual GREP Search/Replace

The GREP steps below assume your document already has the matching character and paragraph styles (Bold, Italic, H1, etc.).

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

8. Footnotes and Endnotes (DOCX inputs):
   - **For actual InDesign notes (recommended)**: Use the included script `endnotes_scripts/notes_converter.jsx`, which converts both kinds in one pass. See [Footnotes and Endnotes](#footnotes-and-endnotes) below.
   - To instead format the references as superscript text:
     - Footnote references: Search for `\[\^F\((\d+)\)\]` — Change to: `$1` (with a superscript character style applied)
     - Endnote references: Search for `\[\^E\((\d+)\)\]` — Change to: `$1` (with a superscript character style applied)
   - Legacy markup (HTML/RTF inputs): Search for `\[\^\((\d+)\)\]\(#fn\d+\)` — Change to: `$1`; for actual InDesign endnotes use `endnotes_scripts/docx_to_indesign_endnotes.jsx`

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

## Footnotes and Endnotes

For **DOCX** inputs, `docx2indesign_advanced.py` reads the note structure directly from the file (`word/footnotes.xml`, `word/endnotes.xml`, and the reference order in `word/document.xml`) rather than relying on pandoc, which collapses both kinds into a single "footnotes" stream. This means footnotes and endnotes stay distinct in the output:

- Footnote reference `[^F(n)]`, with bodies listed under `====== FOOTNOTES ======`
- Endnote reference `[^E(n)]`, with bodies listed under `====== ENDNOTES ======`

The number `n` is a stable join key matching each reference to its body; InDesign renumbers the notes itself once they are created.

To convert both kinds into native InDesign notes in one pass, run `endnotes_scripts/notes_converter.jsx`:

1. Process your Word document with `docx2indesign_advanced.py`
2. Import the resulting text file into InDesign (Unicode UTF-8)
3. Run the script via File > Scripts > Script Panel
4. It inserts native footnotes (via `InsertionPoint.footnotes.add()`) and endnotes (via `Story.endnotes.add()`, with a Convert-To-Endnote menu fallback), sets each note's text, and deletes the trailing note sections.

> Note: HTML/RTF inputs cannot carry the footnote/endnote distinction, so they fall back to the legacy `[^(n)](#fnn)` markers handled by the older endnote scripts below.

## Normalizing Styles Across Files

When a project is split across many DOCX files (e.g. one per chapter), the same
Word style often drifts in its definition from file to file — `Heading 2` might
have a different font or spacing in each. `normalize_docx_styles.py` copies the
style *definitions* from one reference file into the others so they match, **without
touching any text**:

```bash
python normalize_docx_styles.py chapters/*.docx \
  --reference "chapters/Chapter 1.docx" -o ./normalized
```

For every style id a target shares with the reference, the target's definition is
replaced with the reference's; ids the reference lacks are left alone, so no
paragraph loses the style it points at. Each output file is verified to be
byte-for-byte identical to its source in every part except `word/styles.xml`, so
content (text, footnotes, fields, hyperlinks) is provably unchanged. Use
`--dry-run` to preview, and `--include-defaults` to also copy the reference's
document defaults.

Note: for the InDesign pipeline this is largely optional — pandoc maps by
semantic role (every `Heading 2` variant becomes `## ` → InDesign `H2`
regardless of its definition). It is most useful when you also want the Word
files themselves to look consistent, or plan to merge them. It only affects
*named* styles, not direct/manual formatting.

## Using the Legacy Endnote Script

The `endnotes_scripts/docx_to_indesign_endnotes.jsx` script converts the legacy `[^(1)](#fn1)` endnote markers from the Python converter into proper InDesign endnotes:

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