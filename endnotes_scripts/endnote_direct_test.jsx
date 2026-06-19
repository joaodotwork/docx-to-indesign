// Test script for direct creation of endnotes in InDesign 2025
// Run this to create a simple endnote at the cursor position

(function() {
    if (app.documents.length === 0) {
        alert("Please open a document before running this script.");
        return;
    }
    
    var doc = app.activeDocument;
    
    // Method 1: Try using the document.endnotes property
    try {
        alert("STEP 1: Testing endnotes collection\n\nPlace your cursor at the position where you want to add an endnote, then click OK.");
        
        if (doc.hasOwnProperty("endnotes")) {
            // If the cursor is placed in text
            if (app.selection.length > 0 && app.selection[0].hasOwnProperty("insertionPoints")) {
                var endnote = doc.endnotes.add();
                endnote.texts[0].contents = "This is a test endnote created using the endnotes collection.";
                alert("Success! Endnote was created. Move to step 2.");
            } else {
                alert("Please place your cursor in some text before continuing.");
            }
        } else {
            alert("The document doesn't have an 'endnotes' property. Moving to step 2.");
        }
    } catch (e) {
        alert("Method 1 failed: " + e.message + "\nMoving to step 2.");
    }
    
    // Method 2: Try using the text.endnotes property
    try {
        alert("STEP 2: Testing text.endnotes property\n\nPlace your cursor at the position where you want to add an endnote, then click OK.");
        
        if (app.selection.length > 0 && app.selection[0].hasOwnProperty("parent") && app.selection[0].parent.hasOwnProperty("endnotes")) {
            var selText = app.selection[0].parent;
            var endnote = selText.endnotes.add(app.selection[0]);
            endnote.texts[0].contents = "This is a test endnote created using the text.endnotes property.";
            alert("Success! Endnote was created. Move to step 3.");
        } else {
            alert("Selection doesn't have the expected properties. Moving to step 3.");
        }
    } catch (e) {
        alert("Method 2 failed: " + e.message + "\nMoving to step 3.");
    }
    
    // Method 3: Try using the menu action directly
    try {
        alert("STEP 3: Testing using menu actions\n\nSelect a text (don't just place the cursor, but actually select a character or word), then click OK.");
        
        if (app.selection.length > 0 && app.selection[0].hasOwnProperty("contents")) {
            var convertAction = app.menuActions.item("$ID/Convert To Endnote");
            
            if (convertAction.isValid) {
                convertAction.invoke();
                alert("Menu action was invoked. Check if an endnote was created.");
            } else {
                alert("The 'Convert To Endnote' menu action is not valid.");
            }
        } else {
            alert("Please select some text (not just place cursor) before continuing.");
        }
    } catch (e) {
        alert("Method 3 failed: " + e.message);
    }
})();