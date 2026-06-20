#!/usr/bin/env python3
"""Direct DOCX note extraction.

Reads footnote/endnote information straight from the DOCX package XML so the
note *type* (footnote vs. endnote) is preserved. Pandoc collapses both kinds
into a single sequential "footnote" numbering, throwing the distinction away;
this module recovers it.

A DOCX is a ZIP. The relevant parts are::

    word/document.xml   in-text references, in reading order:
                        <w:footnoteReference w:id="N"/> / <w:endnoteReference .../>
    word/footnotes.xml  footnote bodies, keyed by w:id
    word/endnotes.xml   endnote bodies, keyed by w:id

Pandoc numbers notes 1..N in the order their references appear in the body,
regardless of kind. Walking ``document.xml`` in document order yields that same
sequence, so the index (1-based) of each reference here equals the ``#fnN``
number Pandoc emits. That lets the converter map each ``[^(N)](#fnN)`` marker
back to its true kind.
"""

import zipfile
from xml.etree import ElementTree as ET

# WordprocessingML main namespace.
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Note ids <= 0 and these note types are structural (the separator rules drawn
# above/within the notes area), not real notes.
_STRUCTURAL_TYPES = {"separator", "continuationSeparator", "continuationNotice"}


def _qn(tag):
    """Return a namespace-qualified tag name for the WordprocessingML ns."""
    return f"{{{W}}}{tag}"


def _local(tag):
    """Strip the namespace from an ElementTree tag, returning the local name."""
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _note_text(note_elem):
    """Concatenate the visible text of a w:footnote/w:endnote element.

    Multiple paragraphs in a note are joined with a single space so the body
    stays on one line in the line-oriented markup the converter emits.
    """
    parts = []
    for para in note_elem.iter(_qn("p")):
        text = "".join(t.text or "" for t in para.iter(_qn("t")))
        if text:
            parts.append(text)
    return " ".join(parts).strip()


def _parse_note_bodies(zf, part_name, container_tag, note_tag):
    """Build a ``{w:id -> text}`` map from footnotes.xml or endnotes.xml.

    Returns an empty dict when the part is absent (a document may have only one
    kind of note, or none).
    """
    if part_name not in zf.namelist():
        return {}

    bodies = {}
    root = ET.fromstring(zf.read(part_name))
    for note in root.findall(_qn(note_tag)):
        note_type = note.get(_qn("type"))
        if note_type in _STRUCTURAL_TYPES:
            continue
        note_id = note.get(_qn("id"))
        if note_id is None:
            continue
        try:
            if int(note_id) <= 0:
                continue
        except ValueError:
            continue
        bodies[note_id] = _note_text(note)
    return bodies


def _ordered_references(zf):
    """Yield ``(kind, w:id)`` for each note reference in document order.

    ``kind`` is ``"footnote"`` or ``"endnote"``. Order matches Pandoc's
    sequential note numbering.
    """
    root = ET.fromstring(zf.read("word/document.xml"))
    for elem in root.iter():
        name = _local(elem.tag)
        if name == "footnoteReference":
            yield "footnote", elem.get(_qn("id"))
        elif name == "endnoteReference":
            yield "endnote", elem.get(_qn("id"))


def extract_notes(docx_path):
    """Extract notes from a DOCX, preserving footnote/endnote distinction.

    Args:
        docx_path: Path to the .docx file.

    Returns:
        A list of dicts in Pandoc note order, one per in-text reference::

            {"n": 1, "kind": "footnote", "text": "..."}

        ``n`` is the 1-based number matching Pandoc's ``#fnN`` anchor. The list
        is empty when the document contains no notes.
    """
    with zipfile.ZipFile(docx_path) as zf:
        if "word/document.xml" not in zf.namelist():
            return []
        footnote_bodies = _parse_note_bodies(
            zf, "word/footnotes.xml", "footnotes", "footnote"
        )
        endnote_bodies = _parse_note_bodies(
            zf, "word/endnotes.xml", "endnotes", "endnote"
        )

        notes = []
        for index, (kind, note_id) in enumerate(_ordered_references(zf), start=1):
            bodies = footnote_bodies if kind == "footnote" else endnote_bodies
            notes.append(
                {
                    "n": index,
                    "kind": kind,
                    "text": bodies.get(note_id, ""),
                }
            )
        return notes


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 2:
        print("usage: python docx_notes.py input.docx")
        sys.exit(1)
    for note in extract_notes(sys.argv[1]):
        print(f"{note['n']:>3}  {note['kind']:<8}  {note['text']}")
