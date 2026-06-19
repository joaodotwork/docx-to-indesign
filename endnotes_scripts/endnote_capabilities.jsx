// Script to check endnote capabilities in InDesign
// This helps determine which endnote methods are available in your version

(function() {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }
    
    var doc = app.activeDocument;
    var report = "InDesign Endnote Capabilities Report:\n\n";
    
    // Check basic objects
    report += "Document has endnotes collection: " + (doc.hasOwnProperty("endnotes") ? "YES" : "NO") + "\n";
    
    // Check endnote preferences
    try {
        var endnotePref = app.endnoteOptions;
        report += "app.endnoteOptions exists: YES\n";
        report += "   - Properties: ";
        for (var prop in endnotePref) {
            if (typeof endnotePref[prop] !== "function") {
                report += prop + ", ";
            }
        }
        report += "\n";
    } catch (e) {
        report += "app.endnoteOptions exists: NO (" + e.message + ")\n";
    }
    
    // Check if menu actions exist
    try {
        var endnoteMenu = app.menuActions.item("$ID/Endnotes");
        var placeEndnote = app.menuActions.item("$ID/Place Endnote");
        var convertToEndnote = app.menuActions.item("$ID/Convert To Endnote");
        
        report += "Endnote Menu Actions:\n";
        report += "   - $ID/Endnotes exists: " + (endnoteMenu.isValid ? "YES" : "NO") + "\n";
        report += "   - $ID/Place Endnote exists: " + (placeEndnote.isValid ? "YES" : "NO") + "\n";
        report += "   - $ID/Convert To Endnote exists: " + (convertToEndnote.isValid ? "YES" : "NO") + "\n";
    } catch (e) {
        report += "Error checking menu actions: " + e.message + "\n";
    }
    
    // Try to list all available menu actions related to endnotes
    report += "\nSearching for endnote-related menu actions:\n";
    var foundActions = [];
    try {
        for (var i = 0; i < app.menuActions.length; i++) {
            var action = app.menuActions[i];
            if (action.name.toLowerCase().indexOf("endnote") !== -1 || 
                action.name.toLowerCase().indexOf("note") !== -1) {
                foundActions.push(action.name);
            }
        }
        
        if (foundActions.length > 0) {
            report += "Found " + foundActions.length + " related actions:\n";
            for (var j = 0; j < foundActions.length; j++) {
                report += "   - " + foundActions[j] + "\n";
            }
        } else {
            report += "No endnote-related menu actions found.\n";
        }
    } catch (e) {
        report += "Error searching menu actions: " + e.message + "\n";
    }
    
    // Check for methods that might be directly available
    report += "\nChecking for available methods:\n";
    
    // Check selection methods
    try {
        var sel = app.selection;
        if (sel.length > 0) {
            report += "Selection methods:\n";
            var methods = ["convertToEndnote", "createEndnote", "endnoteTextFrame"];
            
            for (var m = 0; m < methods.length; m++) {
                try {
                    if (typeof sel[0][methods[m]] === "function") {
                        report += "   - " + methods[m] + ": Available\n";
                    } else {
                        report += "   - " + methods[m] + ": Not a function\n";
                    }
                } catch (e) {
                    report += "   - " + methods[m] + ": " + e.message + "\n";
                }
            }
        } else {
            report += "No selection to test methods on. Please select some text and run again.\n";
        }
    } catch (e) {
        report += "Error checking selection methods: " + e.message + "\n";
    }
    
    // Display the report
    alert(report);
})();