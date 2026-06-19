// Endnote Converter Helper for DOCX-to-InDesign
// This script helps convert endnote markers to InDesign endnotes by providing a step-by-step workflow
// Usage: Open your document in InDesign, run this script (File > Scripts > Run Script)

(function() {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }
    
    var doc = app.activeDocument;
    var markers = {};
    var endnoteTexts = {};
    
    try {
        // Find all endnote references in the main text
        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findGrepPreferences.findWhat = "\\[\\^\\((\\d+)\\)\\]\\(#fn\\d+\\)";
        
        var endnoteRefs = doc.findGrep();
        
        // Collect reference information
        var refNums = [];
        for (var i = 0; i < endnoteRefs.length; i++) {
            var marker = endnoteRefs[i];
            var match = marker.contents.match(/\[\^\((\d+)\)\]\(#fn\d+\)/);
            if (match) {
                var num = match[1];
                refNums.push(num);
                
                // Store info about this reference for later use
                if (!markers[num]) {
                    markers[num] = [];
                }
                markers[num].push(marker);
            }
        }
        
        // Find all endnote content in numbered list items
        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findGrepPreferences.findWhat = "^(\\d+)\\. (.+?)(?:\\r|$)";
        var listItems = doc.findGrep();
        
        // Extract endnote content for each numbered item
        for (var i = 0; i < listItems.length; i++) {
            var item = listItems[i];
            var match = item.contents.match(/^(\d+)\.\s+(.*?)(?:\r|$)/);
            
            if (match && match[1] && match[2]) {
                var noteNum = match[1];
                var noteText = match[2];
                
                // Clean up the text by removing return links
                if (noteText.indexOf("[↩︎]") > -1) {
                    noteText = noteText.substring(0, noteText.indexOf("[↩︎]"));
                }
                
                // Only store if it matches one of our reference numbers
                if (markers[noteNum]) {
                    endnoteTexts[noteNum] = noteText;
                }
            }
        }
        
        // Check how many definitions we found compared to needed
        var foundCount = 0;
        for (var key in endnoteTexts) {
            foundCount++;
        }
        
        var missingNumbers = [];
        for (var num in markers) {
            if (!endnoteTexts[num]) {
                missingNumbers.push(num);
            }
        }
        
        // If we found no or very few definitions, try a different approach
        if (foundCount < refNums.length / 2) {
            var userChoice = confirm("Only found " + foundCount + " of " + refNums.length + " endnote definitions.\n\n" +
                                    "Would you like to manually create the endnotes by selecting each reference?");
            
            if (userChoice) {
                // Use a manual approach - work through each reference one by one
                for (var num in markers) {
                    if (markers[num] && markers[num].length > 0) {
                        var noteNum = num;
                        var marker = markers[num][0];
                        
                        // First, replace the marker with just the number
                        app.findGrepPreferences = app.changeGrepPreferences = null;
                        app.findGrepPreferences.findWhat = "\\[\\^\\(" + noteNum + "\\)\\]\\(#fn" + noteNum + "\\)";
                        app.changeGrepPreferences.changeTo = noteNum;
                        doc.changeGrep();
                        
                        // Ask user to create endnote manually for each reference
                        var prompt = "For endnote #" + noteNum + ":\n\n" +
                                    "1. Find and select the number '" + noteNum + "' in your document\n" +
                                    "2. Use Type > Endnotes > Convert to Endnote\n" +
                                    "3. Click OK when done, or Cancel to stop";
                                    
                        if (endnoteTexts[noteNum]) {
                            prompt += "\n\nContent: " + endnoteTexts[noteNum];
                        }
                        
                        if (!confirm(prompt)) {
                            break;
                        }
                    }
                }
                
                alert("Endnote preparation complete. Remember to delete the note definitions section at the end of your document.");
                return;
            }
        }
        
        // If we have some but not all definitions, ask how to proceed
        if (foundCount > 0 && missingNumbers.length > 0) {
            var userChoice = confirm("Found " + foundCount + " endnote definitions, but missing definitions for numbers: " + missingNumbers.join(", ") + "\n\n" +
                                    "Would you like to proceed with converting just the ones found?");
            
            if (!userChoice) {
                return;
            }
        }
        
        // Proceed with automatic conversion for the ones we have
        var totalConverted = 0;
        
        for (var num in endnoteTexts) {
            if (markers[num] && markers[num].length > 0) {
                // Replace all markers with this number
                app.findGrepPreferences = app.changeGrepPreferences = null;
                app.findGrepPreferences.findWhat = "\\[\\^\\(" + num + "\\)\\]\\(#fn" + num + "\\)";
                app.changeGrepPreferences.changeTo = num;
                doc.changeGrep();
                
                totalConverted++;
            }
        }
        
        if (totalConverted > 0) {
            var instructions = "Step 1 completed. The references have been converted to simple numbers.\n\n" +
                              "Next steps:\n" +
                              "1. Select all references in your document (you can find them using GREP: '^\\d+$')\n" +
                              "2. Go to Type > Endnotes > Convert to Endnote\n" +
                              "3. Delete the endnote list at the end of your document";
            
            alert(instructions);
        } else {
            // If we couldn't find any endnote definitions, provide assistance
            var sampleReferences = "";
            for (var i = 0; i < Math.min(3, endnoteRefs.length); i++) {
                sampleReferences += "\n- " + endnoteRefs[i].contents;
            }
            
            var sampleNotes = "";
            for (var i = 0; i < Math.min(3, listItems.length); i++) {
                sampleNotes += "\n- " + listItems[i].contents;
            }
            
            var msg = "Could not find matching endnote definitions. Please check your document format:\n\n" +
                      "Found " + endnoteRefs.length + " references:" + sampleReferences + "\n\n" +
                      "Found " + listItems.length + " numbered items:" + sampleNotes + "\n\n" +
                      "The script expects:\n" +
                      "- References in format: [^(1)](#fn1)\n" +
                      "- Endnote text in format: 1. Note text";
            
            alert(msg);
        }
    } catch (e) {
        alert("Error in script execution: " + e + "\nLine: " + e.line);
    }
})();