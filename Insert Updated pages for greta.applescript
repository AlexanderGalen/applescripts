set sourceFolder to "HOM:Products:SPORTS PRODUCTS:FOOTBALL:Football2014:FB-print_files:"
set FinishedFolder to sourceFolder & "2-Sides-FB.PDFs:"
tell application "Finder"
	set FileList to entire contents of folder FinishedFolder as alias list
end tell
repeat with thisFile in FileList
	set thisFile to thisFile as string
	set productCode to characters 83 thru 94 of thisFile as string
	set sourcePDF to sourceFolder & productCode & ".pdf"
	tell application "Adobe Acrobat Pro"
		open thisFile
		open sourcePDF
		replace pages document 1 over 1 from document 2 starting with 1 number of pages 1
		close document 1 with saving
		close document 1
	end tell
end repeat