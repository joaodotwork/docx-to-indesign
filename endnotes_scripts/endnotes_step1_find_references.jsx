// Endnotes Step 1: Find and mark references
// This script finds all endnote references and converts them to simple numbers

(function() {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }
    
    var doc = app.activeDocument;
    var totalRefs = 0;
    
    try {
        // Find all endnote references in the text
        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findGrepPreferences.findWhat = "\\[\\^\\((\\d+)\\)\\]\\(#fn\\d+\\)";
        app.changeGrepPreferences.changeTo = "$1";
        
        // Replace all references with just their number
        var changes = doc.changeGrep();
        totalRefs = changes.length;
        
        if (totalRefs > 0) {
            alert("Found and converted " + totalRefs + " endnote references.\n\n" +
                 "Each reference has been replaced with just its number.\n\n" +
                 "Run step 2 next to identify the endnote content.");
        } else {
            alert("No endnote references found in the document.\n\n" +
                 "The script looks for text in the format: [^(1)](#fn1)");
        }
    } catch (e) {
        alert("Error: " + e);
    }
})();