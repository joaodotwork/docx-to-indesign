#!/usr/bin/env python3
"""Normalize paragraph/character style *definitions* across DOCX files.

Word stores a document's text (runs in ``word/document.xml``) separately from
its style *definitions* (``word/styles.xml``) and the style *references* that
paragraphs/runs point at (``w:pStyle`` / ``w:rStyle`` by style id). Because of
that separation, the visual definition of a style can be made identical across
many files without touching a single character of text.

This tool copies the style definitions from a chosen reference DOCX into each
target DOCX. For every style id the target shares with the reference, the
target's definition is replaced with the reference's; style ids the reference
does not have are left untouched (so nothing a paragraph points at goes
missing). The result is written to a new file — originals are never modified —
and every non-style part of the package is asserted byte-for-byte identical, so
"without altering content" is verified, not just intended.

What this does NOT change:
    * Any text, footnotes, endnotes, fields (e.g. Zotero), or hyperlinks.
    * Direct/manual formatting applied to runs without a named style.

Usage:
    python normalize_docx_styles.py TARGET.docx [TARGET2.docx ...] \
        --reference REF.docx -o OUTPUT_DIR [--include-defaults] [--dry-run]
"""

import argparse
import shutil
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STYLES_PART = "word/styles.xml"
# Parts whose bytes must never change — the actual content of the document.
CONTENT_PARTS = (
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
)

# Register the common OOXML namespaces so serialized prefixes stay readable.
# (Word resolves styles by namespace URI, not prefix, so this is cosmetic, but
# it keeps the output diff-friendly.)
for prefix, uri in {
    "w": W,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}.items():
    ET.register_namespace(prefix, uri)


def _qn(tag):
    return f"{{{W}}}{tag}"


def _style_key(style_elem):
    """Identity of a <w:style>: its (type, styleId)."""
    return (style_elem.get(_qn("type")), style_elem.get(_qn("styleId")))


def _read_styles_root(docx_path):
    """Return the parsed <w:styles> root element of a DOCX."""
    with zipfile.ZipFile(docx_path) as zf:
        if STYLES_PART not in zf.namelist():
            raise ValueError(f"{docx_path}: no {STYLES_PART}")
        return ET.fromstring(zf.read(STYLES_PART))


def normalize_styles_xml(target_root, ref_root, include_defaults=False):
    """Overlay the reference's style definitions onto the target's styles tree.

    Returns (updated_ids, added_ids) for reporting. Mutates ``target_root``.
    """
    ref_styles = {_style_key(s): s for s in ref_root.findall(_qn("style"))}
    target_styles = {_style_key(s): s for s in target_root.findall(_qn("style"))}

    updated, added = [], []

    # Replace shared styles in place; append reference-only styles at the end.
    children = list(target_root)
    for key, ref_style in ref_styles.items():
        if key in target_styles:
            old = target_styles[key]
            target_root[children.index(old)] = ref_style
            children[children.index(old)] = ref_style
            updated.append(key[1])
        else:
            target_root.append(ref_style)
            added.append(key[1])

    if include_defaults:
        ref_defaults = ref_root.find(_qn("docDefaults"))
        tgt_defaults = target_root.find(_qn("docDefaults"))
        if ref_defaults is not None and tgt_defaults is not None:
            target_root[list(target_root).index(tgt_defaults)] = ref_defaults

    return updated, added


def _rewrite_with_styles(src_docx, dst_docx, styles_bytes):
    """Copy src_docx to dst_docx, swapping in new word/styles.xml bytes."""
    with zipfile.ZipFile(src_docx) as zin:
        items = zin.infolist()
        with zipfile.ZipFile(dst_docx, "w") as zout:
            for item in items:
                data = zin.read(item.filename)
                if item.filename == STYLES_PART:
                    data = styles_bytes
                # Preserve each entry's original compression type.
                zout.writestr(item, data)


def _assert_content_unchanged(src_docx, dst_docx):
    """Verify every content part is byte-identical between src and dst."""
    with zipfile.ZipFile(src_docx) as a, zipfile.ZipFile(dst_docx) as b:
        for part in CONTENT_PARTS:
            names = a.namelist()
            if part in names:
                if a.read(part) != b.read(part):
                    raise AssertionError(f"content changed in {part}!")
        # Every part other than styles.xml must be untouched.
        for name in a.namelist():
            if name != STYLES_PART and a.read(name) != b.read(name):
                raise AssertionError(f"unexpected change in {name}!")


def normalize_file(target_path, ref_root, out_dir, include_defaults, dry_run):
    """Normalize one DOCX, writing the result into out_dir."""
    target_path = Path(target_path)
    target_root = _read_styles_root(target_path)
    updated, added = normalize_styles_xml(
        target_root, ref_root, include_defaults
    )

    print(f"{target_path.name}: {len(updated)} style(s) updated, "
          f"{len(added)} added")
    if dry_run:
        return True

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    dst = out_dir / target_path.name

    styles_bytes = ET.tostring(target_root, encoding="UTF-8", xml_declaration=True)
    _rewrite_with_styles(target_path, dst, styles_bytes)

    # Guarantee the promise: only styles changed, never content.
    _assert_content_unchanged(target_path, dst)
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Normalize style definitions across DOCX files "
                    "without altering content."
    )
    parser.add_argument("targets", nargs="+", help="DOCX files to normalize")
    parser.add_argument("--reference", required=True,
                        help="DOCX whose styles are the canonical source")
    parser.add_argument("-o", "--output", default="./normalized",
                        help="Output directory (default: ./normalized)")
    parser.add_argument("--include-defaults", action="store_true",
                        help="Also copy the reference's docDefaults block")
    parser.add_argument("--dry-run", action="store_true",
                        help="Report changes without writing files")
    args = parser.parse_args()

    ref_path = Path(args.reference)
    try:
        ref_root = _read_styles_root(ref_path)
    except Exception as e:
        print(f"error: cannot read reference styles: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Reference: {ref_path.name}\n")
    failed = 0
    for target in args.targets:
        if Path(target).resolve() == ref_path.resolve():
            print(f"{Path(target).name}: skipped (is the reference)")
            continue
        try:
            normalize_file(target, ref_root, args.output,
                           args.include_defaults, args.dry_run)
        except Exception as e:
            failed += 1
            print(f"{Path(target).name}: FAILED — {e}", file=sys.stderr)

    if not args.dry_run and failed == 0:
        print(f"\nNormalized files written to: {Path(args.output).resolve()}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
