// Endnote Converter for DOCX-to-InDesign
// This script converts endnote markers from the docx2indesign converter to InDesign endnotes
// Usage: Open your document in InDesign, run this script (File > Scripts > Run Script)

(function() {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }
    
    try {
        var doc = app.activeDocument;
        var references = [];
        var noteContent = {};
        
        // Step 1: Find and collect all references
        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findGrepPreferences.findWhat = "\\[\\^\\((\\d+)\\)\\]\\(#fn\\d+\\)";
        var foundRefs = doc.findGrep();
        
        for (var i = 0; i < foundRefs.length; i++) {
            var refText = foundRefs[i].contents;
            var match = refText.match(/\[\^\((\d+)\)\]/);
            if (match) {
                references.push({
                    ref: foundRefs[i],
                    num: match[1]
                });
            }
        }
        
        // Step 2: Find all note definitions
        // Look for text with the return link markup
        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findGrepPreferences.findWhat = "\\[↩︎\\]\\(#fnref(\\d+)\\)";
        var foundNotes = doc.findGrep();
        
        for (var i = 0; i < foundNotes.length; i++) {
            var noteItem = foundNotes[i];
            var noteText = noteItem.contents;
            
            // Extract number from the return link
            var match = noteText.match(/\[↩︎\]\(#fnref(\d+)\)/);
            if (match) {
                var num = match[1];
                
                // Get the paragraph containing this note
                var para = noteItem.paragraphs[0];
                var content = para.contents;
                
                // Clean up content by removing the return link
                if (content.indexOf("[↩︎]") > -1) {
                    content = content.substr(0, content.indexOf("[↩︎]"));
                }
                
                noteContent[num] = {
                    text: content,
                    item: para
                };
            }
        }
        
        // Count how many notes we found
        var noteCount = 0;
        for (var key in noteContent) {
            if (noteContent.hasOwnProperty(key)) noteCount++;
        }
        
        if (noteCount === 0) {
            // Try looking for the numbered items at the end of the document
            app.findGrepPreferences = app.changeGrepPreferences = null;
            app.findGrepPreferences.findWhat = "^(\\d+)\\. (.+?)$";
            var foundNumberedNotes = doc.findGrep();
            
            for (var i = 0; i < foundNumberedNotes.length; i++) {
                var noteText = foundNumberedNotes[i].contents;
                var match = noteText.match(/^(\d+)\./);
                if (match) {
                    var num = match[1];
                    var content = noteText.substr(noteText.indexOf(". ") + 2);
                    
                    // Clean up content
                    if (content.indexOf("[↩︎]") > -1) {
                        content = content.substr(0, content.indexOf("[↩︎]"));
                    }
                    
                    noteContent[num] = {
                        text: content,
                        item: foundNumberedNotes[i]
                    };
                }
            }
        }
        
        // Step 3: Process each reference to create endnotes
        var processCount = 0;
        
        // Replace all markers with just numbers first
        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findGrepPreferences.findWhat = "\\[\\^\\((\\d+)\\)\\]\\(#fn\\d+\\)";
        app.changeGrepPreferences.changeTo = "$1";
        doc.changeGrep();
        
        // Create a character style for endnote references if one doesn't exist
        var endnoteStyle;
        try {
            endnoteStyle = doc.characterStyles.itemByName("Endnote_Reference");
            // Check if the style exists
            endnoteStyle.name;
        } catch (e) {
            // Create the style
            endnoteStyle = doc.characterStyles.add({name: "Endnote_Reference"});
            endnoteStyle.baselineShift = "1.5pt";  // Superscript
            endnoteStyle.fontStyle = "Regular";
            endnoteStyle.pointSize = "9pt";
        }
        
        // Now find and select each number to apply the character style
        for (var i = 0; i < references.length; i++) {
            var num = references[i].num;
            
            // Find this reference number in the text
            app.findGrepPreferences = app.changeGrepPreferences = null;
            app.findGrepPreferences.findWhat = "\\b" + num + "\\b";
            var foundItems = doc.findGrep();
            
            for (var j = 0; j < foundItems.length; j++) {
                // Apply character style
                foundItems[j].appliedCharacterStyle = endnoteStyle;
                processCount++;
            }
        }
        
        // Ask if user wants to attempt to automatically create endnotes (InDesign 2020+ feature)
        if (processCount > 0) {
            if (confirm("References have been formatted with character style. InDesign 2020+ can convert these to actual endnotes.\n\nWould you like to attempt automatic conversion?")) {
                try {
                    // This is for InDesign 2020+ which supports endnote features
                    if (app.menuActions.item("$ID/Place Endnote").enabled) {
                        // Select all text with the endnote style
                        app.findGrepPreferences = app.changeGrepPreferences = null;
                        app.findGrepPreferences.appliedCharacterStyle = endnoteStyle;
                        var allRefs = doc.findGrep();
                        
                        for (var i = 0; i < allRefs.length; i++) {
                            var ref = allRefs[i];
                            var num = ref.contents;
                            
                            if (noteContent[num]) {
                                // Select this reference
                                ref.select();
                                
                                // Try to convert to endnote via menu action
                                app.menuActions.item("$ID/Endnotes").select();
                                app.menuActions.item("$ID/Place Endnote").invoke();
                                
                                // Get the created endnote and set its content
                                if (doc.endnotes.length > 0) {
                                    var newEndnote = doc.endnotes[-1]; // Last created endnote
                                    newEndnote.texts[0].contents = noteContent[num].text;
                                }
                            }
                        }
                    } else {
                        throw new Error("Automatic endnote creation not available");
                    }
                } catch (e) {
                    alert("Automatic endnote creation failed: " + e.message + "\n\nPlease manually create endnotes by:\n\n1. Select each reference number\n2. Use Type > Endnotes > Convert to Endnote");
                }
            }
        }
        
        // Count note definitions
        var noteCount = 0;
        for (var key in noteContent) {
            if (noteContent.hasOwnProperty(key)) {
                noteCount++;
            }
        }
        
        // Final report
        alert("Processing complete!\n" + 
              "- Found " + references.length + " references\n" + 
              "- Found " + noteCount + " note definitions\n" +
              "- Formatted " + processCount + " references\n\n" +
              "All references have been formatted with the 'Endnote_Reference' character style.\n\n" +
              "You may now want to delete the endnote list at the end of your document.");
        
    } catch (e) {
        alert("Error: " + e + " (line " + e.line + ")");
    }
})();