// Endnote Converter for InDesign 2025
// This script finds and formats endnote references ready for conversion

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
        
        // Step 2: Find all note definitions with return links
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
        
        // Report initial findings
        alert("References found: " + references.length + "\nNote definitions found: " + noteCount + "\n\nClick OK to proceed with conversion.");
        
        // Track how many references were processed
        var processCount = 0;
        
        // Since direct endnote creation isn't available, format the references
        if (references.length > 0) {
            try {
                // Create or get endnote reference style
                var endnoteStyle;
                try {
                    endnoteStyle = doc.characterStyles.itemByName("Endnote_Reference");
                    // Check if style exists
                    endnoteStyle.name;
                } catch(e) {
                    // Create the style
                    endnoteStyle = doc.characterStyles.add({name: "Endnote_Reference"});
                    endnoteStyle.baselineShift = "1.5pt";  // Superscript
                    endnoteStyle.fontStyle = "Regular";
                    endnoteStyle.pointSize = "9pt";
                }
                
                // Replace all markers with just numbers
                app.findGrepPreferences = app.changeGrepPreferences = null;
                app.findGrepPreferences.findWhat = "\\[\\^\\((\\d+)\\)\\]\\(#fn\\d+\\)";
                app.changeGrepPreferences.changeTo = "$1";
                doc.changeGrep();
                
                // Now find all numbers and apply style
                for (var i = 0; i < references.length; i++) {
                    var num = references[i].num;
                    
                    app.findGrepPreferences = app.changeGrepPreferences = null;
                    app.findGrepPreferences.findWhat = "\\b" + num + "\\b";
                    var foundNumbers = doc.findGrep();
                    
                    for (var j = 0; j < foundNumbers.length; j++) {
                        foundNumbers[j].appliedCharacterStyle = endnoteStyle;
                        processCount++;
                    }
                }
                
                // Create a text file with endnote contents for reference
                var endnoteFile = new File(Folder.desktop + "/endnote_content.txt");
                endnoteFile.encoding = "UTF-8";
                if (endnoteFile.open("w")) {
                    endnoteFile.write("ENDNOTE CONTENT:\n\n");
                    for (var num in noteContent) {
                        if (noteContent.hasOwnProperty(num)) {
                            endnoteFile.write("ENDNOTE " + num + ":\n" + noteContent[num].text + "\n\n");
                        }
                    }
                    endnoteFile.close();
                }
                
                // Create a summary document of all endnote info
                var endnoteMapDoc = null;
                try {
                    endnoteMapDoc = app.documents.add();
                    var textFrame = endnoteMapDoc.pages[0].textFrames.add();
                    textFrame.geometricBounds = [20, 20, 200, 180]; // [y1, x1, y2, x2]
                    
                    var content = "ENDNOTE CONTENT (FOR REFERENCE):\n\n";
                    for (var num in noteContent) {
                        if (noteContent.hasOwnProperty(num)) {
                            content += "ENDNOTE " + num + ":\n" + noteContent[num].text + "\n\n";
                        }
                    }
                    
                    textFrame.contents = content;
                    
                    // Save the document to desktop
                    var saveFile = new File(Folder.desktop + "/Endnote_Content.indd");
                    endnoteMapDoc.save(saveFile);
                } catch (e) {
                    alert("Could not create endnote content document: " + e.message);
                }
            } catch (e) {
                alert("Processing failed: " + e.message);
            }
        }
        
        // Final report and instructions
        if (processCount > 0) {
            alert("Processing complete!\n\n" + 
                  "- Found " + references.length + " references\n" + 
                  "- Found " + noteCount + " note definitions\n" + 
                  "- Formatted " + processCount + " references\n\n" +
                  "Since automatic endnote creation isn't available in this version of InDesign, follow these steps:\n\n" +
                  "1. All references are now formatted with the 'Endnote_Reference' character style\n" +
                  "2. Select each reference and use Type > Convert Selection to Endnote (or similar menu item)\n" +
                  "3. A text file with all endnote content has been saved to your desktop\n" +
                  "4. Copy and paste the appropriate content into each endnote\n" +
                  "5. Delete the endnote list at the end of your document when done");
        } else {
            alert("No references were processed. Please check your document and try again.");
        }
    } catch (e) {
        alert("Error: " + e.message + " (line " + e.line + ")");
    }
})();