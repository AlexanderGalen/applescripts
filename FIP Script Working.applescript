--declares the log to file function and the read excel doc function

--returns a list of all data in excel file
--format is: each item of first list is a line from original document, each item of that list is a cell from that line
on readExcelDoc(thisFile)
	
	set tempFile to (path to desktop) & ":temp.txt" as string
	
	tell application "Microsoft Excel"
		open thisFile
		save as active sheet of active workbook filename tempFile file format text Mac file format with overwrite
		close active workbook saving no
	end tell
	
	set theData to read alias tempFile using delimiter return
	set tableData to {}
	set text item delimiters to tab
	repeat with i from 1 to count of theData
		set theLine to text items of item i of theData
		copy theLine to the end of tableData
	end repeat
	set text item delimiters to ""
	do shell script "rm -f " & quoted form of POSIX path of tempFile
	return tableData
	
end readExcelDoc


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

--checks to make sure all pdfs exist before moving on to actually working on them

repeat with TheItem in filelist
	tell application "Finder"
		set ExitVar to ""
		set {name:FileName, name extension:fileExtension} to TheItem
		set ExcelDoc to TheItem
	end tell
	
	--gets contents of excel document with predefined readExcelDoc function
	set theData to readExcelDoc(ExcelDoc)
	
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
end repeat


--actually works on the orders
repeat with TheItem in filelist
	
	set timeString to time string of (current date) as string
	if length of timeString is 11 then
		set StartTime to characters 1 thru 8 of timeString as string
	else
		set StartTime to "0" & characters 1 thru 7 of timeString as string
	end if
	
	set startSeconds to time of (current date)
	
	set ExitVar to ""
	
	tell application "Finder"
		set {name:FileName, name extension:fileExtension} to TheItem
		set ExcelDoc to TheItem
	end tell
	
	set theData to readExcelDoc(ExcelDoc)
	set rowCount to count of theData
	
	set OrderNumber to item 1 of item 2 of theData
	set ClientName to item 2 of item 2 of theData
	set thisQTY to ""
	
	--compiles the finished file name from info in the data
	set finishedFilePath to "MERGE CENTRAL:FIP AUTOMATION:~~~Orders:FIP_" & OrderNumber & "." & ClientName & ".pdf"
	
	--Makes Coverpage in Quark with Order Number and Client Name
	
	tell application "QuarkXPress"
		open file "MERGE CENTRAL:FIP AUTOMATION:CoverSheet:Cover.qxp"
		tell document 1
			set text of text box 1 to "FIP_" & OrderNumber & " " & ClientName
		end tell
		export layout space 1 of project 1 in "MERGE CENTRAL:FIP AUTOMATION:CoverSheet:Cover.pdf" as "PDF"
		close project 1 without saving
	end tell
	
	
	--opens the coverpage and saves with new file name
	
	tell application "Adobe Acrobat Pro"
		close every document
		open file "MERGE CENTRAL:FIP AUTOMATION:CoverSheet:Cover.pdf"
		tell document 1
			save to file finishedFilePath
		end tell
	end tell
	
	tell application "Finder"
		delete "MERGE CENTRAL:FIP AUTOMATION:CoverSheet:Cover.pdf"
	end tell
	
	--checks for value of quantity column, row by row, until the row with "TOTAL" in the first column, which should be the final row.
	set TotalQty to item 3 of item rowCount of theData
	repeat with i from 2 to rowCount - 2
		
		set thisQTY to item 3 of item i of theData
		
		--if there is a value in the quantity column, starts process of adding pages
		
		if thisQTY is not "" then
			set trueQTY to (thisQTY / 6)
			
			--gets the name of the pdf and sets the path to it as a variable
			set YearFileName to item 4 of item i of theData
			set FileName to "FIP_" & YearFileName as string
			set pathtopdf to "MERGE CENTRAL:FIP AUTOMATION:Found Image Press Calendars:" & FileName as string
			
			--gets number of pages in each open pdf
			tell application "Adobe Acrobat Pro"
				tell document 1
					set Doc1PageQty to count pages
				end tell
				open pathtopdf
				tell document 2
					set Doc2PageQty to count pages
				end tell
				
				--adds the pages from second document to first document after last page; repeats for however many of that calendar they ordered
				with timeout of 86400 seconds
					repeat trueQTY times
						insert pages document 1 after Doc1PageQty from document 2 starting with 1 number of pages Doc2PageQty
						tell document 1
							set Doc1PageQty to count pages
						end tell
					end repeat
				end timeout
				close document 2
			end tell
			
			--if qty is 0, continue to next row.
		end if
	end repeat
	
	--saves changes made to new PDF
	with timeout of 86400 seconds
		tell application "Adobe Acrobat Pro"
			save document 1 to finishedFilePath
			close document 1
		end tell
	end timeout
	
	set timeString to time string of (current date) as string
	if length of timeString is 11 then
		set FinishTime to characters 1 thru 8 of timeString as string
	else
		set FinishTime to "0" & characters 1 thru 7 of timeString as string
	end if
	
	set endSeconds to time of (current date)
	
	set secondDiff to endSeconds - startSeconds
	set timeDiff to format (secondDiff / 60) into "000.00"
	set TotalQty to format TotalQty into "000"
	
	--opens the logs and inputs this job's info at the end of the log
	set thisUsername to short user name of (system info)
	set logText to TotalQty & tab & OrderNumber & tab & StartTime & tab & FinishTime & tab & timeDiff & tab & thisUsername as string
	set LogFile to "MERGE CENTRAL:FIP AUTOMATION:Logs:FIP Automation Log.txt"
	
	logToFile(logText, LogFile)
	
	--moves merged excel doc into merged folder
	
	tell application "Finder"
		duplicate ExcelDoc to "ART DEPARTMENT-NEW:For SQL:FIP:Merged:"
		delete ExcelDoc
	end tell
	
end repeat