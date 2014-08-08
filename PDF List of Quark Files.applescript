--there's gotta be a better way to do this, I just didn't have time to figure it out.

tell application "Finder"
	set FinishedPath to "HOM:PRODUCTS:CALENDARS:2015 CALENDARS:2015 Calendar Pads:~PRINT FILES CORRECTIONS:"
	set CalPads to "HOM:PRODUCTS:CALENDARS:2015 CALENDARS:2015 Calendar Pads:"
	set file7 to CalPads & "7.Fresh and Healthy:2015.FreshandHealthy.qxp"
	set file8 to CalPads & "8.Cupcakes:2015.Cupcakes.qxp"
	set file9 to CalPads & "9.Scenic America:2015.Scenic_America.qxp"
	set file10 to CalPads & "10.HOME QUOTES:2015.HomeQuotes.qxp"
	
	set FileList to {file7, file8, file9, file10} as alias list
	
end tell

tell application "QuarkXPress"
	repeat with theItem in FileList
		open file theItem
		set theName to name of document 1
		set theLength to length of theName
		set theName to characters 1 thru (theLength - 4) of theName
		set FinishedPDFName to FinishedPath & theName & ".pdf"
		
		export layout space 1 of project 1 in FinishedPDFName as "PDF" PDF output style "No Compression Print"
		
		close document 1 without saving
	end repeat
end tell