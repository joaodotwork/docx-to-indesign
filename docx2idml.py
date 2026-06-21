#!/usr/bin/env python3
"""DOCX to InDesign IDML (experimental).

Generates a finished InDesign IDML directly from a DOCX, skipping the import +
GREP step: paragraphs/runs land in real paragraph and character styles, and
footnotes become native InDesign footnotes (with inline italics/bold preserved).

It works by templating from a real IDML you export from your InDesign template
(File > Export > IDML). Style ids, document preferences, spread, and frame are
all taken from that package; only the body story is replaced. Style ids are
read verbatim from the template's Styles.xml and never reconstructed — InDesign
style names can hide non-breaking spaces and bullets that must match exactly.

Limitations (tracked separately):
    * Single text frame — long stories overset; multi-page autoflow is a
      separate piece of work (see the repo issues).
    * Endnotes are treated as footnotes.
    * Hyperlinks are styled but not made clickable.

Usage:
    python docx2idml.py input.docx --template template.idml -o output.idml
"""

import argparse
import html
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from bs4 import BeautifulSoup, NavigableString

from docx_notes import extract_notes

MIMETYPE = "mimetype"
NO_PARA = "ParagraphStyle/$ID/NormalParagraphStyle"
NO_CHAR = "CharacterStyle/$ID/[No character style]"
# Footnote auto-number marker + separator, exactly as InDesign serializes them.
FOOTNOTE_MARKER = "<?ACE 4?>\t"


# --- template introspection ------------------------------------------------

def load_self_map(template_zip):
    """Map (style_kind, normalized_name) -> exact Self id from Styles.xml.

    Names are normalized by turning non-breaking spaces into regular spaces so
    callers can look up "body • flow" without typing the hidden NBSP, while the
    returned Self id keeps the exact characters InDesign expects.
    """
    root = ET.fromstring(template_zip.read("Resources/Styles.xml"))
    local = lambda tag: tag.split("}")[-1]
    styles = {}
    for el in root.iter():
        if local(el.tag) in ("ParagraphStyle", "CharacterStyle") and el.get("Self"):
            name = (el.get("Name") or "").replace(" ", " ").strip()
            styles[(local(el.tag), name)] = el.get("Self")
    return styles


def find_body_story(template_zip):
    """Return (story_self, story_part) for the template's page text frame."""
    for name in template_zip.namelist():
        if name.startswith("Spreads/"):
            data = template_zip.read(name).decode("utf-8")
            m = re.search(r'<TextFrame\b[^>]*\bParentStory="([^"]+)"', data)
            if m:
                part = f"Stories/Story_{m.group(1)}.xml"
                if part in template_zip.namelist():
                    return m.group(1), part
    raise ValueError("No body text frame/story found in the template's spreads.")


def read_footnote_styles(template_zip):
    """Read (text_paragraph_style, marker_character_style) from FootnoteOption.

    These are the document's footnote defaults; applying them to generated
    footnotes means whatever you set in Type > Document Footnote Options is what
    the output uses. Falls back to the InDesign defaults if not found.
    """
    text_style, marker_style = NO_PARA, NO_CHAR
    for name in template_zip.namelist():
        if name.endswith(".xml"):
            data = template_zip.read(name).decode("utf-8")
            m = re.search(r"<FootnoteOption\b[^>]*>", data)
            if m:
                opt = m.group(0)
                t = re.search(r'FootnoteTextStyle="([^"]+)"', opt)
                k = re.search(r'FootnoteMarkerStyle="([^"]+)"', opt)
                if t:
                    text_style = t.group(1)
                if k:
                    marker_style = k.group(1)
                return text_style, marker_style
    return text_style, marker_style


_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def recover_title(docx_path):
    """Return the chapter title text, read straight from the DOCX.

    pandoc extracts the first Title-styled paragraph into document metadata and
    drops it from the body, so the chapter title would otherwise be lost. We
    read it back from the DOCX and re-insert it as an H1 (deduped against the
    body in case pandoc kept it). Prefers Word's Title style, then Heading1.
    """
    q = lambda t: f"{{{_W}}}{t}"
    doc = ET.fromstring(zipfile.ZipFile(docx_path).read("word/document.xml"))
    body = doc.find(q("body"))
    found = {}
    for p in body.findall(q("p")):
        ppr = p.find(q("pPr"))
        sid = None
        if ppr is not None and ppr.find(q("pStyle")) is not None:
            sid = ppr.find(q("pStyle")).get(q("val"))
        text = "".join(t.text or "" for t in p.iter(q("t"))).strip()
        if text and sid in ("Title", "Heading1") and sid not in found:
            found[sid] = text
    return found.get("Title") or found.get("Heading1")


# --- DOCX -> structured runs -----------------------------------------------

def docx_to_html(docx_path):
    """Convert DOCX to HTML via pandoc, preserving Word style names."""
    return subprocess.run(
        ["pandoc", "-f", "docx+styles", "-t", "html", "--wrap=none", str(docx_path)],
        capture_output=True, text=True, check=True,
    ).stdout


class StyleResolver:
    """Maps semantic roles (heading level, emphasis, ...) to exact style ids."""

    def __init__(self, self_map):
        self._map = self_map

    def para(self, name):
        return self._map.get(("ParagraphStyle", name), NO_PARA)

    def char(self, name):
        return self._map.get(("CharacterStyle", name), NO_CHAR)

    def paragraph_style(self, el):
        if el.name and el.name[0] == "h" and el.name[1:].isdigit():
            return self.para("H" + str(min(int(el.name[1:]), 4)))
        for anc in el.parents:
            if getattr(anc, "name", None) == "blockquote":
                return self.para("Block Quote")
            cs = anc.get("data-custom-style") if hasattr(anc, "get") else None
            if cs == "Block Quote":
                return self.para("Block Quote")
            if cs == "Bibliography":
                return self.para("Bibliography")
            # Word's Title / Heading 1 become the chapter title -> H1. (pandoc
            # keeps Title as a div here; see recover_title for the case where
            # it strips the first one into document metadata instead.)
            if cs == "Title" or (cs or "").lower() in ("heading 1", "heading1"):
                return self.para("H1")
        return self.para("body • flow")

    def character_style(self, bold, italic, sup, sub, link):
        if sup:
            return self.char("Superscript")
        if sub:
            return self.char("Subscript")
        if bold and italic:
            return self.char("Bold Italic")
        if bold:
            return self.char("Bold")
        if italic:
            return self.char("Italic")
        if link:
            return self.char("Hyperlink")
        return NO_CHAR


def _is_footnote_ref(node):
    return node.name == "a" and "footnote-ref" in (node.get("class") or [])


def paragraph_tokens(el, resolver):
    """Flatten a paragraph into ('text', cstyle, text) and ('fn', n) tokens."""
    out = []

    def walk(node, bold, italic, sup, sub, link):
        for child in node.children:
            if isinstance(child, NavigableString):
                text = str(child)
                if text:
                    out.append(("text", resolver.character_style(
                        bold, italic, sup, sub, link), text))
            elif _is_footnote_ref(child):
                out.append(("fn", int(child.get("href", "#fn0").split("fn")[-1])))
            else:
                walk(child,
                     bold or child.name in ("strong", "b"),
                     italic or child.name in ("em", "i"),
                     sup or child.name == "sup",
                     sub or child.name == "sub",
                     link or child.name == "a")

    walk(el, False, False, False, False, False)

    merged = []
    for tok in out:
        if (tok[0] == "text" and merged and merged[-1][0] == "text"
                and merged[-1][1] == tok[1]):
            merged[-1] = ("text", tok[1], merged[-1][2] + tok[2])
        else:
            merged.append(tok)
    return merged


def footnote_bodies(soup, resolver):
    """Return {n: [(cstyle, text), ...]} with footnote formatting preserved."""
    bodies = {}
    section = soup.find("section", class_="footnotes")
    if not section:
        return bodies
    for li in section.find_all("li", id=True):
        m = re.match(r"fn(\d+)$", li.get("id", ""))
        if not m:
            continue
        for back in li.find_all("a", class_="footnote-back"):
            back.decompose()
        runs = []
        for pi, para in enumerate(li.find_all("p") or [li]):
            if pi:
                runs.append((NO_CHAR, " "))  # join multi-paragraph notes
            runs += [(t[1], t[2]) for t in paragraph_tokens(para, resolver)
                     if t[0] == "text"]
        if runs:
            runs[0] = (runs[0][0], runs[0][1].lstrip())
            runs[-1] = (runs[-1][0], runs[-1][1].rstrip())
            runs = [(cs, t) for cs, t in runs if t]
        bodies[int(m.group(1))] = runs
    return bodies


# --- story serialization ---------------------------------------------------

def _esc(text):
    return html.escape(text, quote=False)


def _char_range(cstyle, text):
    return (f'<CharacterStyleRange AppliedCharacterStyle="{cstyle}">'
            f'<Content>{_esc(text)}</Content></CharacterStyleRange>')


def _footnote_range(runs, text_style, marker_style):
    if not runs:
        runs = [(NO_CHAR, "")]
    # The footnote's own leading number (ACE marker) is left unstyled, as
    # InDesign does — only the in-text reference marker gets marker_style.
    inner = (f'<CharacterStyleRange AppliedCharacterStyle="{NO_CHAR}">'
             f'<Content>{FOOTNOTE_MARKER}</Content></CharacterStyleRange>')
    inner += "".join(_char_range(cs, t) for cs, t in runs)
    return (f'<CharacterStyleRange AppliedCharacterStyle="{marker_style}" '
            f'Position="Superscript"><Footnote>'
            f'<ParagraphStyleRange AppliedParagraphStyle="{text_style}">'
            f'{inner}</ParagraphStyleRange></Footnote></CharacterStyleRange>')


def build_story_xml(html_text, resolver, story_self, title=None,
                    fn_text_style=NO_PARA, fn_marker_style=NO_CHAR):
    """Build the IDML Story part for the document body."""
    soup = BeautifulSoup(html_text, "html.parser")
    fn_runs = footnote_bodies(soup, resolver)
    section = soup.find("section", class_="footnotes")
    if section:
        section.decompose()

    paras = [p for p in soup.find_all(["p", "h1", "h2", "h3", "h4", "h5", "h6"])
             if p.get_text(strip=True)]

    # Re-insert the chapter title pandoc dropped into metadata, unless it is
    # already the first paragraph (i.e. pandoc kept it as a Title div).
    h1_style = resolver.para("H1")
    if title and not (paras and paras[0].get_text(strip=True) == title):
        out_title = [
            f'\t\t<ParagraphStyleRange AppliedParagraphStyle="{h1_style}">',
            f'\t\t\t<CharacterStyleRange AppliedCharacterStyle="{NO_CHAR}">'
            f'<Content>{_esc(title)}</Content>'
            + ("<Br />" if paras else "") + '</CharacterStyleRange>',
            '\t\t</ParagraphStyleRange>',
        ]
    else:
        out_title = []

    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<idPkg:Story xmlns:idPkg="http://ns.adobe.com/AdobeInDesign/idml/1.0/'
        'packaging" DOMVersion="21.4">',
        f'\t<Story Self="{story_self}" UserText="true" IsEndnoteStory="false" '
        'AppliedTOCStyle="n" TrackChanges="false" StoryTitle="$ID/" '
        'AppliedNamedGrid="n">',
        '\t\t<StoryPreference OpticalMarginAlignment="false" '
        'OpticalMarginSize="12" FrameType="TextFrameType" '
        'StoryOrientation="Horizontal" StoryDirection="LeftToRightDirection" />',
        '\t\t<InCopyExportOption IncludeGraphicProxies="true" '
        'IncludeAllResources="false" />',
    ]
    out += out_title
    for idx, para in enumerate(paras):
        last = idx == len(paras) - 1
        out.append('\t\t<ParagraphStyleRange AppliedParagraphStyle='
                   f'"{resolver.paragraph_style(para)}">')
        tokens = paragraph_tokens(para, resolver)
        for ti, tok in enumerate(tokens):
            br = "<Br />" if (ti == len(tokens) - 1 and not last) else ""
            if tok[0] == "text":
                out.append("\t\t\t" + _char_range(tok[1], tok[2]).replace(
                    "</Content>", f"</Content>{br}" if br else "</Content>"))
            else:
                out.append("\t\t\t" + _footnote_range(
                    fn_runs.get(tok[1], []), fn_text_style, fn_marker_style))
                if br:
                    out.append(f'\t\t\t<CharacterStyleRange '
                               f'AppliedCharacterStyle="{NO_CHAR}">{br}'
                               '</CharacterStyleRange>')
        out.append('\t\t</ParagraphStyleRange>')
    out += ['\t</Story>', '</idPkg:Story>', '']
    return "\n".join(out)


def write_idml(template_path, out_path, story_part, story_bytes):
    """Copy the template IDML, swapping in the new body story.

    mimetype must be the first entry and stored (uncompressed).
    """
    with zipfile.ZipFile(template_path) as zin:
        infos = zin.infolist()
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            zout.writestr(zipfile.ZipInfo(MIMETYPE), zin.read(MIMETYPE),
                          compress_type=zipfile.ZIP_STORED)
            for info in infos:
                if info.filename == MIMETYPE:
                    continue
                data = (story_bytes if info.filename == story_part
                        else zin.read(info.filename))
                zout.writestr(info.filename, data)


def convert(docx_path, template_path, out_path):
    """Convert one DOCX into an IDML using the given template."""
    with zipfile.ZipFile(template_path) as zt:
        resolver = StyleResolver(load_self_map(zt))
        story_self, story_part = find_body_story(zt)
        fn_text_style, fn_marker_style = read_footnote_styles(zt)

    html_text = docx_to_html(docx_path)
    story = build_story_xml(html_text, resolver, story_self,
                            title=recover_title(docx_path),
                            fn_text_style=fn_text_style,
                            fn_marker_style=fn_marker_style)
    ET.fromstring(story.encode("utf-8"))  # fail fast on malformed XML

    write_idml(template_path, out_path, story_part, story.encode("utf-8"))
    notes = extract_notes(docx_path)
    return len(notes)


def main():
    parser = argparse.ArgumentParser(
        description="Convert a DOCX into a styled InDesign IDML (experimental).")
    parser.add_argument("input", help="Input .docx file")
    parser.add_argument("--template", required=True,
                        help="IDML exported from your InDesign template")
    parser.add_argument("-o", "--output", help="Output .idml (default: alongside input)")
    args = parser.parse_args()

    docx = Path(args.input)
    out = Path(args.output) if args.output else docx.with_suffix(".idml")
    try:
        notes = convert(docx, args.template, out)
    except Exception as e:
        print(f"error: {e}", file=sys.stderr)
        sys.exit(1)
    print(f"Wrote {out} ({notes} footnote(s))")


if __name__ == "__main__":
    main()
