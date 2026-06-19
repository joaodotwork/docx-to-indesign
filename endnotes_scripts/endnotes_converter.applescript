-- Endnote Converter for DOCX-to-InDesign
-- This AppleScript converts endnote references to InDesign endnotes

tell application "Adobe InDesign 2025"
	if (count of documents) is 0 then
		display dialog "Please open a document before running this script." buttons {"OK"} default button "OK"
		return
	end if
	
	set theDoc to active document
	set endnoteRefs to {}
	set endnoteTexts to {}
	set refCount to 0
	set converted to 0
	
	-- Step 1: Find all endnote references
	tell theDoc
		-- Reset the find/change settings
		set find grep preferences to nothing
		set change grep preferences to nothing
		
		-- Find endnote references like [^(1)](#fn1)
		set find grep preferences's find what to "\\[\\^\\((\\d+)\\)\\]\\(#fn\\d+\\)"
		
		-- Collect all references
		set foundItems to find grep
		set refCount to count of foundItems
		
		if refCount is 0 then
			display dialog "No endnote references found in this document." buttons {"OK"} default button "OK"
			return
		end if
		
		-- Get the reference numbers
		repeat with i from 1 to refCount
			set thisRef to item i of foundItems
			set refText to contents of thisRef
			
			-- Extract the number
			set numStart to offset of "(" in refText
			set numEnd to offset of ")" in refText
			if numStart > 0 and numEnd > 0 then
				set refNum to text (numStart + 1) thru (numEnd - 1) of refText
				copy refNum to end of endnoteRefs
			end if
		end repeat
		
		-- Step 2: Find endnote content in numbered list items
		set find grep preferences to nothing
		set find grep preferences's find what to "^(\\d+)\\. (.+)$"
		
		set listItems to find grep
		set noteCount to count of listItems
		
		-- Process each list item to find endnote content
		repeat with i from 1 to noteCount
			set thisItem to item i of listItems
			set itemText to contents of thisItem
			
			-- Extract the number
			set dotPos to offset of ". " in itemText
			if dotPos > 0 then
				set noteNum to text 1 thru (dotPos - 1) of itemText
				set noteContent to text (dotPos + 2) to (count of itemText) of itemText
				
				-- Clean up content (remove return links)
				set linkPos to offset of "[↩︎]" in noteContent
				if linkPos > 0 then
					set noteContent to text 1 thru (linkPos - 1) of noteContent
				end if
				
				-- Store the content if it's a number in our list
				if noteNum is in endnoteRefs then
					set end of endnoteTexts to {num:noteNum, content:noteContent, item:thisItem}
				end if
			end if
		end repeat
		
		-- Step 3: Create actual endnotes
		display dialog "Found " & refCount & " references and " & (count of endnoteTexts) & " endnote definitions. Ready to convert?" buttons {"Cancel", "Continue"} default button "Continue"
		
		-- Replace all reference markers with just the numbers
		set find grep preferences to nothing
		set change grep preferences to nothing
		set find grep preferences's find what to "\\[\\^\\((\\d+)\\)\\]\\(#fn\\d+\\)"
		set change grep preferences's change to to "$1"
		change grep
		
		-- Let the user know what to do next
		display dialog "References have been replaced with their numbers. Now please:\n\n1. Select all numbers that should be endnotes\n2. Use Type > Convert to Endnote\n3. Delete the endnote list at the end of your document" buttons {"OK"} default button "OK"
	end tell
end tell