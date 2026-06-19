// Improved Endnote Converter for InDesign 2025
// This script converts endnote markers from docx2indesign to actual InDesign endnotes

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
        
        // Step 3: Process each reference to create endnotes
        if (references.length > 0) {
            // Try modern InDesign workflow first - set content after conversion
            try {
                for (var i = 0; i < references.length; i++) {
                    var refObj = references[i];
                    var num = refObj.num;
                    var ref = refObj.ref;
                    
                    if (noteContent[num]) {
                        // Replace the marker with just the number temporarily
                        app.findGrepPreferences = app.changeGrepPreferences = null;
                        app.findGrepPreferences.findWhat = "\\[\\^\\(" + num + "\\)\\]\\(#fn" + num + "\\)";
                        app.changeGrepPreferences.changeTo = num;
                        ref.changeGrep();
                        
                        // Re-search to find our number
                        app.findGrepPreferences = app.changeGrepPreferences = null;
                        app.findGrepPreferences.findWhat = "\\b" + num + "\\b";
                        var numberRefs = doc.findGrep();
                        
                        if (numberRefs.length > 0) {
                            // Select this number to create an endnote
                            var textRef = numberRefs[0];
                            textRef.select();
                            
                            // Different methods for different InDesign versions
                            
                            // Method 1: Try direct Story.endnotes approach
                            try {
                                var story = textRef.parentStory;
                                if (story.hasOwnProperty("endnotes")) {
                                    var endnote = story.endnotes.add(textRef);
                                    endnote.texts[0].contents = noteContent[num].text;
                                    processCount++;
                                } else {
                                    throw new Error("Story doesn't have endnotes property");
                                }
                            } catch (e1) {
                                // Method 2: Try the menu action approach
                                try {
                                    var convertAction = app.menuActions.item("$ID/Convert To Endnote");
                                    if (convertAction.isValid) {
                                        convertAction.invoke();
                                        
                                        // Attempt to get the created endnote
                                        if (doc.hasOwnProperty("endnotes") && doc.endnotes.length > 0) {
                                            var latestEndnote = doc.endnotes[doc.endnotes.length - 1];
                                            latestEndnote.texts[0].contents = noteContent[num].text;
                                        }
                                        processCount++;
                                    } else {
                                        throw new Error("Convert to Endnote menu action is not valid");
                                    }
                                } catch (e2) {
                                    alert("Failed to convert reference " + num + ": " + e2.message);
                                }
                            }
                        }
                    }
                }
            } catch (e) {
                alert("Automatic conversion failed: " + e.message);
                
                // Fallback to formatting with character style
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
                    
                    // Apply style to all reference numbers
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
                } catch (e) {
                    alert("Formatting fallback also failed: " + e.message);
                }
            }
        }
        
        // Final report
        if (processCount > 0) {
            alert("Conversion complete!\n\n" + 
                  "- Found " + references.length + " references\n" + 
                  "- Found " + noteCount + " note definitions\n" + 
                  "- Processed " + processCount + " references\n\n" +
                  "If references were just formatted (not converted to actual endnotes), select each reference and use Type > Endnotes > Convert to Endnote.\n\n" +
                  "You may now want to delete the endnote list at the end of your document.");
        } else {
            alert("No references were processed. Please check your document and try again.");
        }
    } catch (e) {
        alert("Error: " + e.message + " (line " + e.line + ")");
    }
})();