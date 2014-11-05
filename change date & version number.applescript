--gets current date
set {year:y, month:m, day:d} to (current date)
set m to m * 1
set thisdate to m & "/" & d & "/" & y as string

tell application "QuarkXPress"
	tell document 1
		set word 5 of story 1 of text box "Proof Sheet Info" to thisdate
		set proofSheetInfoText to story 1 of text box "Proof Sheet Info"
		set AppleScript's text item delimiters to "."
		set oldVersionNumber to last text item of proofSheetInfoText
		set newVersionNumber to oldVersionNumber + 1
		set theseTextItems to text items of proofSheetInfoText
		set last item of theseTextItems to newVersionNumber
		set proofSheetInfoText to theseTextItems as string
		set story 1 of text box "Proof Sheet Info" to proofSheetInfoText
	end tell
end tell