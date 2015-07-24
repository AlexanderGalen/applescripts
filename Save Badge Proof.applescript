tell application "QuarkXPress"
	tell project 1
		set active layout space to layout space "B"
		set layoutName to name of active layout space
	end tell
	tell document 1
		set qxpName to name
		set jobNumber to characters 1 thru 6 of qxpName as string
		set theProofInfo to story 1 of text box "proofInfo"
		set text item delimiters to "."
		set text item 1 of theProofInfo to jobNumber

	end tell
end tell

