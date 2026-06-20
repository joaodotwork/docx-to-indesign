// Batch GREP Formatter for InDesign
//
// Runs every markup-to-formatting GREP conversion from the README in one pass,
// in the correct order, applying the matching paragraph/character styles. This
// replaces entering each pattern by hand in Find/Change.
//
// Requirements:
//   - The styles referenced below must already exist in the document. The
//     included indesign-template/docx2indd-template.indt provides them; load
//     its styles into your document first (or rename the entries here).
//   - Run AFTER importing the processed .txt, and BEFORE notes_converter.jsx.
//     Note references ([^F(n)] / [^E(n)]) are intentionally left untouched so
//     notes_converter.jsx can turn them into real footnotes/endnotes.
//
// Missing styles are reported and their rule is skipped, not fatal.

(function () {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }

    var doc = app.activeDocument;

    // ---- style lookup (returns null if the named style does not exist) ----

    function charStyle(name) {
        var s = doc.characterStyles.itemByName(name);
        return s.isValid ? s : null;
    }
    function paraStyle(name) {
        var s = doc.paragraphStyles.itemByName(name);
        return s.isValid ? s : null;
    }

    // ---- rules, in dependency order -------------------------------------
    // kind: "char" or "para" selects which applied-style property to set.
    // Order matters: triple-marker rules run before double, double before
    // single, deepest heading before shallowest, so patterns don't clobber.

    var rules = [
        { name: "Bold Italic",   find: "\\*\\*\\*([^*]+)\\*\\*\\*", style: "Bold Italic",   kind: "char" },
        { name: "Bold",          find: "\\*\\*([^*]+)\\*\\*",       style: "Bold",          kind: "char" },
        { name: "Underline",     find: "__([^_]+)__",               style: "Underline",     kind: "char" },
        { name: "Italic",        find: "_([^_]+)_",                 style: "Italic",        kind: "char" },
        { name: "Superscript",   find: "\\^\\(([^)]+)\\)",          style: "Superscript",   kind: "char" },
        { name: "Subscript",     find: "~\\(([^)]+)\\)",            style: "Subscript",     kind: "char" },
        { name: "Heading 4",     find: "^####\\s+(.+)$",            style: "H4",            kind: "para" },
        { name: "Heading 3",     find: "^###\\s+(.+)$",             style: "H3",            kind: "para" },
        { name: "Heading 2",     find: "^##\\s+(.+)$",              style: "H2",            kind: "para" },
        { name: "Heading 1",     find: "^#\\s+(.+)$",               style: "H1",            kind: "para" },
        { name: "Bulleted List", find: "^\\*\\s+(.+)$",             style: "Bulleted List", kind: "para" },
        { name: "Numbered List", find: "^\\d+\\.\\s+(.+)$",         style: "Numbered List", kind: "para" },
        // Links: strip markup and tag the visible text. The destination URL in
        // group 2 is dropped here; see createHyperlinks below for live links.
        { name: "Links",         find: "\\[([^\\]]+)\\]\\(([^)]+)\\)", style: "Hyperlink", kind: "char" }
    ];

    // ---- run ------------------------------------------------------------

    var applied = [];
    var missing = [];

    for (var i = 0; i < rules.length; i++) {
        var rule = rules[i];
        var style = rule.kind === "char" ? charStyle(rule.style) : paraStyle(rule.style);
        if (!style) {
            missing.push(rule.name + " (needs " + rule.kind + " style \"" + rule.style + "\")");
            continue;
        }

        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findGrepPreferences.findWhat = rule.find;
        app.changeGrepPreferences.changeTo = "$1";
        if (rule.kind === "char") {
            app.changeGrepPreferences.appliedCharacterStyle = style;
        } else {
            app.changeGrepPreferences.appliedParagraphStyle = style;
        }

        var changed = doc.changeGrep();
        applied.push(rule.name + ": " + changed.length);
    }

    app.findGrepPreferences = app.changeGrepPreferences = null;

    var msg = "Batch GREP formatting complete.\n\nApplied:\n" + applied.join("\n");
    if (missing.length) {
        msg += "\n\nSkipped (style not found):\n" + missing.join("\n");
    }
    msg += "\n\nNext: run notes_converter.jsx to convert [^F(n)] / [^E(n)] " +
           "markers into native footnotes and endnotes.";
    alert(msg);
})();
