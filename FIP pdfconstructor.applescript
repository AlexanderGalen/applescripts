--makes xml document for pdfconstructor

--declares the log to file function
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
	
	set RowNumber to 2
	
	tell application "Microsoft Excel"
		open ExcelDoc
	end tell
	
	repeat while ExitVar is not "Exit"
		repeat 1 times
			tell application "Microsoft Excel"
				tell row RowNumber
					if value of column 1 is "TOTAL" then
						--performs a full exit
						set ExitVar to "Exit"
						exit repeat
					end if
					
					set Column3Value to value of column 3
					
					--if there is a value in the quantity column, starts process of adding pages
					
					if Column3Value is not "" then
						set YearFileName to value of column 4
						set FileName to "FIP_" & YearFileName as string
						set pathtopdf to "MERGE CENTRAL:FIP AUTOMATION:Found Image Press Calendars:" & FileName as string
						
						tell application "Finder"
							set fileExists to exists of pathtopdf
						end tell
						
						if not fileExists then
							tell application "Microsoft Excel"
								set excelName to name of document 1
							end tell
							display dialog FileName & " of Excel Document " & excelName & " was not found."
						end if
						
					end if
					
					set RowNumber to RowNumber + 1
				end tell
			end tell
		end repeat
	end repeat
	set ExitVar to ""
	tell application "Microsoft Excel"
		close active workbook without saving
	end tell
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
	
	set ExcelDoc to TheItem as string
	
	---- STARTS TO BUILD XML FOR PDF CONSTRUCTOR ------
	
	set r to 2
	set FIPAutomation to "Merge Central:FIP AUTOMATION:"
	set pdfBasePath to FIPAutomation & "Found Image Press Calendars:"
	set theXMLFile to FIPAutomation & "pdfconstructor:FIP_Construct.pdfc"
	set finishedXML to ""
	
	tell application "Microsoft Excel"
		open file ExcelDoc
		tell row 2
			set orderNumber to string value of cell 1
			set clientName to string value of cell 2
		end tell
	end tell
	
	set coverText to "FIP_" & orderNumber & " " & clientName
	set finishedPDF to FIPAutomation & "~~~Orders:FIP_" & orderNumber & "." & clientName & ".pdf"
	
	set finishedXML to finishedXML & "<?xml version='1.0' encoding='UTF-8'?>
	<docasm linearized='true' version='1.4'>
		<resources>
			<frame id='f1' rect='0 0 432 450'/>
		</resources>
		<pages>
			<page>
				<boxes>
					<MediaBox width='432' height='450' x='0' y='0'/>
				</boxes>
				<elements>
					<text font='Helvetica' font-size='18' color='[.4 .3 .3 1]' x='Center' y='Center' style='text-align:center' frame='#f1'>" & coverText & "</text>
				</elements>
			</page>" & return
	
	repeat
		tell application "Microsoft Excel"
			tell row r
				if value of column 1 is "TOTAL" then
					set TotalQty to value of column 3
					exit repeat
				end if
				
				set thisPDF to string value of column 4
				set qty to ((value of column 3) / 6)
			end tell
		end tell
		
		set thisPDF to pdfBasePath & "FIP_" & thisPDF as string
		set thisPDF to POSIX path of thisPDF
		repeat qty times
			set finishedXML to finishedXML & "		<insert href='" & thisPDF & "' range='0-13'/>" & return
		end repeat
		set r to r + 1
	end repeat
	
	tell application "Microsoft Excel"
		close active workbook saving no
	end tell
	
	set finishedXML to finishedXML & "	</pages>
	</docasm>"
	
	set fileToEdit to open for access file theXMLFile with write permission
	set eof fileToEdit to 0
	write finishedXML to fileToEdit starting at eof
	close access file theXMLFile
	
	
	----- END OF XML BUILD ------
	(*
	------Calls PDF Constructor -----
	
	set finishedPDF to quoted form of POSIX path of finishedPDF
	set theXMLFile to quoted form of POSIX path of theXMLFile
	
	with timeout of 86400 seconds
		do shell script "pdfconstructor -f " & theXMLFile & " -o " & finishedPDF
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
	set logText to TotalQty & tab & orderNumber & tab & StartTime & tab & FinishTime & tab & timeDiff & tab & thisUsername as string
	set LogFile to "MERGE CENTRAL:FIP AUTOMATION:Logs:FIP Automation Log.txt"
	
	logToFile(logText, LogFile)
	
	--moves merged excel doc into merged folder
	
	tell application "Finder"
		duplicate ExcelDoc to "ART DEPARTMENT-NEW:For SQL:FIP:Merged:"
		delete ExcelDoc
	end tell
	*)
	
end repeat