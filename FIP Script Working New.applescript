--declares the log to file function and the read excel doc function

--returns file path to tab delimited file

on excelToTabDelimited(thisFile)
	
	set tabDelimitedFile to "MERGE CENTRAL:FIP AUTOMATION:Working:tab-delimited.txt"
	
	
	tell application "Microsoft Excel"
		open thisFile
		save as active sheet of active workbook filename tabDelimitedFile file format text Mac file format with overwrite
		close active workbook saving no
		return tabDelimitedFile
	end tell
	
end excelToTabDelimited

--returns a list of all data in excel file
--format is: each item of first list is a line from original document, each item of that list is a cell from that line
on parseTabDelimited(thisFile)
	
	set theData to read alias thisFile using delimiter return
	set tableData to {}
	set text item delimiters to tab
	repeat with i from 1 to count of theData
		set theLine to text items of item i of theData
		copy theLine to the end of tableData
	end repeat
	set text item delimiters to ""
	return tableData
	
end parseTabDelimited


on logToFile(logText, LogFile)
	set errorOccured to false
	open for access file LogFile with write permission
	write (logText & return) to file LogFile starting at eof
	close access file LogFile
end logToFile

--gets contents of hot folder
tell application "Finder"
	set filelist to files of folder POSIX file "/Volumes/MERGE CENTRAL/FIP AUTOMATION/Hot Folder/" as alias list
end tell



tell application "Microsoft Excel"
	open
	close every document without saving
end tell
(*
--checks to make sure all pdfs exist before moving on to actually working on them
repeat with TheItem in filelist
	tell application "Finder"
		set {name:FileName, name extension:fileExtension} to TheItem
		set ExcelDoc to TheItem
	end tell
	
	--gets contents of excel document with predefined excelToTabDelimited function
	set thisTabDelimited to excelToTabDelimited(ExcelDoc)
	set theData to parseTabDelimited(thisTabDelimited)
	
	set rowCount to count of theData
	
	--loops through all "rows" of the data
	repeat with i from 2 to rowCount - 2
		set thisQTY to item 3 of item i of theData
		
		if thisQTY is not "" then
			set YearFileName to item 4 of item i of theData
			set FileName to "FIP_" & YearFileName as string
			set pathtopdf to "MERGE CENTRAL:FIP AUTOMATION:Found Image Press Calendars:" & FileName as string
			
			tell application "Finder"
				set fileExists to exists of pathtopdf
			end tell
			
			if not fileExists then
				set ExcelName to ExcelDoc
				display dialog FileName & " of Excel Document " & ExcelName & " was not found."
			end if
			
		end if
	end repeat
end repeat*)

--actually works on the orders, using node
repeat with TheItem in filelist
	
	--gets starting time before starting, for logging purposes
	set timeString to time string of (current date) as string
	if length of timeString is 11 then
		set StartTime to characters 1 thru 8 of timeString as string
	else
		set StartTime to "0" & characters 1 thru 7 of timeString as string
	end if
	set startSeconds to time of (current date)
	
	--saves a tab delimited file for node to use
	tell application "Finder"
		set {name:FileName, name extension:fileExtension} to TheItem
		set ExcelDoc to TheItem
	end tell
	excelToTabDelimited(ExcelDoc)
	
	--timeout is to prevent the shell script from causing applescript to timeout if it takes too long to finish
	with timeout of 4000 seconds
		do shell script "/usr/local/bin/node 'Volumes/MERGE CENTRAL/FIP AUTOMATION/Working/FIP Node.js'"
	end timeout
	
	
	--gets finished time for logging purposes
	set timeString to time string of (current date) as string
	if length of timeString is 11 then
		set FinishTime to characters 1 thru 8 of timeString as string
	else
		set FinishTime to "0" & characters 1 thru 7 of timeString as string
	end if
	set endSeconds to time of (current date)
	set secondDiff to endSeconds - startSeconds
	set timeDiff to format (secondDiff / 60) into "000.00"
	
	
	--opens the logs and inputs this job's info at the end of the log
	set thisUsername to short user name of (system info)
	set logText to StartTime & tab & FinishTime & tab & timeDiff & tab & thisUsername as string
	set LogFile to "MERGE CENTRAL:FIP AUTOMATION:Logs:FIP Automation Log.txt"
	
	logToFile(logText, LogFile)
	
	--moves merged excel doc into merged folder
	
	tell application "Finder"
		duplicate ExcelDoc to "ART DEPARTMENT-NEW:For SQL:FIP:Merged:"
		delete ExcelDoc
	end tell
	
end repeat