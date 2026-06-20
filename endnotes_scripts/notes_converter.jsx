// Notes Converter for InDesign (2018+)
//
// Converts the type-specific note markup emitted by docx2indesign_advanced.py
// into native InDesign footnotes AND endnotes.
//
// Expected markup (produced when notes are read directly from the DOCX):
//
//   In-text reference (footnote):  [^F(3)]
//   In-text reference (endnote):   [^E(2)]
//
//   ====== FOOTNOTES ======
//   [^F1]: This is the first footnote text.
//   [^F3]: This is the second footnote text.
//
//   ====== ENDNOTES ======
//   [^E2]: This is the only endnote text.
//
// The number is a stable join key (pandoc's sequential note number). InDesign
// renumbers footnotes/endnotes automatically once they are created, so the
// original numbers do not need to be preserved.
//
// Footnotes use InsertionPoint.footnotes.add(); endnotes use
// Story.endnotes.add() with a Convert-To-Endnote menu fallback for older
// builds. Run with a document open and the imported text as the active layout.

(function () {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }

    var doc = app.activeDocument;

    // ---- helpers --------------------------------------------------------

    function resetGrep() {
        app.findGrepPreferences = app.changeGrepPreferences = null;
    }

    // Build {number: bodyText} by scanning the trailing note section.
    // prefix is "F" (footnotes) or "E" (endnotes).
    function collectBodies(prefix) {
        var bodies = {};
        resetGrep();
        // Whole body line: [^F12]: the text...
        app.findGrepPreferences.findWhat = "\\[\\^" + prefix + "(\\d+)\\]: (.+)";
        var found = doc.findGrep();
        for (var i = 0; i < found.length; i++) {
            var m = found[i].contents.match(
                new RegExp("\\[\\^" + prefix + "(\\d+)\\]: ([\\s\\S]+)")
            );
            if (m) {
                bodies[m[1]] = m[2];
            }
        }
        return bodies;
    }

    // Remove every paragraph whose text matches the given GREP. Re-finds after
    // each removal so shifting text indices never go stale.
    function removeMatchingParagraphs(grep) {
        var guard = 0;
        while (guard++ < 5000) {
            resetGrep();
            app.findGrepPreferences.findWhat = grep;
            var found = doc.findGrep();
            if (found.length === 0) break;
            try {
                found[0].paragraphs[0].remove();
            } catch (e) {
                break;
            }
        }
    }

    // Set a note's body text without clobbering its auto number/separator by
    // appending at the very end of the note story.
    function setNoteText(note, text) {
        try {
            note.insertionPoints[-1].contents = text;
        } catch (e) {
            note.texts[0].contents = text; // last-resort fallback
        }
    }

    // ---- footnotes ------------------------------------------------------

    function convertFootnotes() {
        var bodies = collectBodies("F");
        var made = 0, failed = 0;
        var guard = 0;

        // Re-find the first remaining reference each pass: inserting a footnote
        // and deleting the marker both change text positions.
        while (guard++ < 5000) {
            resetGrep();
            app.findGrepPreferences.findWhat = "\\[\\^F\\((\\d+)\\)\\]";
            var found = doc.findGrep();
            if (found.length === 0) break;

            var ref = found[0];
            var m = ref.contents.match(/\[\^F\((\d+)\)\]/);
            var num = m ? m[1] : null;

            try {
                var ip = ref.insertionPoints[0];
                var fn = ip.footnotes.add(); // marker inserted before the ref
                if (num !== null && bodies[num] !== undefined) {
                    setNoteText(fn, bodies[num]);
                }
                ref.contents = ""; // delete the [^F(n)] marker text
                made++;
            } catch (e) {
                // Avoid an infinite loop on a marker we cannot convert.
                try { ref.contents = ""; } catch (e2) {}
                failed++;
            }
        }
        return { made: made, failed: failed };
    }

    // ---- endnotes -------------------------------------------------------

    function convertEndnotes() {
        var bodies = collectBodies("E");
        var made = 0, failed = 0;
        var guard = 0;

        while (guard++ < 5000) {
            resetGrep();
            app.findGrepPreferences.findWhat = "\\[\\^E\\((\\d+)\\)\\]";
            var found = doc.findGrep();
            if (found.length === 0) break;

            var ref = found[0];
            var m = ref.contents.match(/\[\^E\((\d+)\)\]/);
            var num = m ? m[1] : null;
            var story = ref.parentStory;

            try {
                var ip = ref.insertionPoints[0];
                var endnote = null;

                // Preferred: Story.endnotes.add() at the insertion point.
                try {
                    endnote = story.endnotes.add(ip);
                } catch (e1) {
                    // Fallback: select the marker and use the menu action.
                    ref.select();
                    var action = app.menuActions.item("$ID/Convert To Endnote");
                    if (action.isValid && action.enabled) {
                        action.invoke();
                        if (doc.endnotes.length > 0) {
                            endnote = doc.endnotes[doc.endnotes.length - 1];
                        }
                    }
                }

                if (endnote && num !== null && bodies[num] !== undefined) {
                    setNoteText(endnote, bodies[num]);
                }
                try { ref.contents = ""; } catch (e3) {}
                made++;
            } catch (e) {
                try { ref.contents = ""; } catch (e4) {}
                failed++;
            }
        }
        return { made: made, failed: failed };
    }

    // ---- run ------------------------------------------------------------

    try {
        var fnResult = convertFootnotes();
        var enResult = convertEndnotes();

        // Clean up the trailing note sections now that bodies are converted.
        removeMatchingParagraphs("\\[\\^F\\d+\\]: ");
        removeMatchingParagraphs("\\[\\^E\\d+\\]: ");
        removeMatchingParagraphs("^====== FOOTNOTES ======$");
        removeMatchingParagraphs("^====== ENDNOTES ======$");
        resetGrep();

        alert(
            "Notes conversion complete.\n\n" +
            "Footnotes created: " + fnResult.made +
            (fnResult.failed ? " (failed: " + fnResult.failed + ")" : "") + "\n" +
            "Endnotes created: " + enResult.made +
            (enResult.failed ? " (failed: " + enResult.failed + ")" : "") + "\n\n" +
            "Review the document; InDesign has renumbered the notes automatically."
        );
    } catch (e) {
        resetGrep();
        alert("Error: " + e.message + (e.line ? " (line " + e.line + ")" : ""));
    }
})();
