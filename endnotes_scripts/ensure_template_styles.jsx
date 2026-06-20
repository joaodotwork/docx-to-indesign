// Ensure Template Styles for InDesign
//
// Creates any paragraph/character styles that the docx2indd GREP queries and
// batch_grep_format.jsx expect but that are missing from the current document.
// Idempotent: existing styles are left exactly as they are; only missing ones
// are created, with modest starter formatting you can refine afterwards.
//
// Use this to bring the template (docx2indd-template.indt) up to date — open
// the template, run this script, then File > Save to update the .indt.

(function () {
    if (app.documents.length === 0) {
        alert("Please open the template (or a document) before running this script.");
        return;
    }

    var doc = app.activeDocument;
    var created = [];
    var skipped = [];

    function ensureParagraphStyle(name, props) {
        var s = doc.paragraphStyles.itemByName(name);
        if (s.isValid) { skipped.push("¶ " + name); return; }
        s = doc.paragraphStyles.add({ name: name });
        applyProps(s, props);
        created.push("¶ " + name);
    }

    function ensureCharacterStyle(name, props) {
        var s = doc.characterStyles.itemByName(name);
        if (s.isValid) { skipped.push("A " + name); return; }
        s = doc.characterStyles.add({ name: name });
        applyProps(s, props);
        created.push("A " + name);
    }

    // Apply each property defensively — a font/style not present on the system
    // shouldn't abort creation of the rest.
    function applyProps(style, props) {
        for (var key in props) {
            if (!props.hasOwnProperty(key)) continue;
            try { style[key] = props[key]; } catch (e) {}
        }
    }

    // --- Paragraph styles ---------------------------------------------------
    // Headings and lists are normally already in the template (the existing
    // queries reference them); listed here so a bare document still works.
    ensureParagraphStyle("H1", { pointSize: 24, spaceBefore: 18, spaceAfter: 9 });
    ensureParagraphStyle("H2", { pointSize: 18, spaceBefore: 14, spaceAfter: 7 });
    ensureParagraphStyle("H3", { pointSize: 14, spaceBefore: 11, spaceAfter: 6 });
    ensureParagraphStyle("H4", { pointSize: 12, spaceBefore: 9, spaceAfter: 5 });
    ensureParagraphStyle("Bulleted List", { leftIndent: 18, firstLineIndent: -9 });
    ensureParagraphStyle("Numbered List", { leftIndent: 18, firstLineIndent: -9 });

    // The genuinely new ones for this workflow:
    ensureParagraphStyle("Block Quote", {
        leftIndent: 18, rightIndent: 18, spaceBefore: 6, spaceAfter: 6,
        fontStyle: "Italic"
    });
    ensureParagraphStyle("Bibliography", {
        // hanging indent, typical for reference lists
        leftIndent: 18, firstLineIndent: -18, spaceAfter: 4
    });

    // --- Character styles ---------------------------------------------------
    ensureCharacterStyle("Bold", { fontStyle: "Bold" });
    ensureCharacterStyle("Italic", { fontStyle: "Italic" });
    ensureCharacterStyle("Bold Italic", { fontStyle: "Bold Italic" });
    ensureCharacterStyle("Underline", { underline: true });
    ensureCharacterStyle("Superscript", { position: Position.SUPERSCRIPT });
    ensureCharacterStyle("Subscript", { position: Position.SUBSCRIPT });
    ensureCharacterStyle("Hyperlink", { underline: true });

    var msg = "Style check complete.\n\n";
    msg += "Created (" + created.length + "):\n" +
           (created.length ? created.join("\n") : "  (none)") + "\n\n";
    msg += "Already present (" + skipped.length + "):\n" +
           (skipped.length ? skipped.join("\n") : "  (none)");
    if (created.length) {
        msg += "\n\nReview the new styles' formatting, then File > Save to update the template.";
    }
    alert(msg);
})();
