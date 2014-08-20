set StartTime to characters 1 through 8 of time string of (current date) as text
set startSeconds to time of (current date)

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
	
	set ExitVar to ""
	
	
	tell application "Finder"
		set {name:FileName, name extension:fileExtension} to TheItem
		set ExcelDoc to TheItem
	end tell
	
	set RowNumber to 2
	
	tell application "Microsoft Excel"
		
		open ExcelDoc
		
		--gets order info from excel doc and intitializes some variables
		
		set OrderNumber to string value of column 1 of row 2
		set ClientName to string value of column 2 of row 2
		set Column3Value to ""
		
	end tell
	
	--compiles the finished file name from info in the excel doc
	set FinishedFilePath to "MERGE CENTRAL:FIP AUTOMATION:~~~Orders:FIP_" & OrderNumber & "." & ClientName & ".pdf"
	
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
			save to file FinishedFilePath
		end tell
	end tell
	
	tell application "Finder"
		delete "MERGE CENTRAL:FIP AUTOMATION:CoverSheet:Cover.pdf"
	end tell
	
	--checks for value of quantity column, row by row, until the row with "TOTAL" in the first column, which should be the final row.
	
	
	repeat while ExitVar is not "Exit"
		repeat 1 times
			tell application "Microsoft Excel"
				tell row RowNumber
					if value of column 1 is "TOTAL" then
						set TotalQty to value of column 3
						--performs a full exit
						set ExitVar to "Exit"
						exit repeat
					end if
					
					set Column3Value to value of column 3
					
					--if there is a value in the quantity column, starts process of adding pages
					
					if Column3Value is not "" then
						set Qty to (Column3Value / 6)
						
						--gets the name of the pdf and sets the path to it as a variable
						
						set YearFileName to value of column 4
						set FileName to "FIP_" & YearFileName as string
						set pathtopdf to "MERGE CENTRAL:FIP AUTOMATION:Found Image Press Calendars:" & FileName as string
						
						--checks that the pdf exists. displays dialog if it is not found.
						tell application "Finder"
							if not (exists pathtopdf) then
								tell application "System Events"
									set dialogvar to display dialog "PDF (" & pathtopdf & ") was not found. Check File Name" buttons {"Skip Row", "Try Anyway"} default button 1
								end tell
								
								
								if button returned of dialogvar is "Skip Row" then
									
									--exits just the inner loop, skipping over current row.
									set RowNumber to RowNumber + 1
									exit repeat
									
								end if
							end if
						end tell
						
						
						
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
								repeat Qty times
									insert pages document 1 after Doc1PageQty from document 2 starting with 1 number of pages Doc2PageQty
									tell document 1
										set Doc1PageQty to count pages
									end tell
								end repeat
							end timeout
							close document 2
						end tell
						set RowNumber to RowNumber + 1
					else
						set RowNumber to RowNumber + 1
					end if
				end tell
			end tell
		end repeat
	end repeat
	set ExitVar to ""
	
	--closes the excel document without saving	
	
	
	tell application "Microsoft Excel"
		close active workbook saving no
	end tell
	
	--moves merged excel doc into merged folder
	
	tell application "Finder"
		duplicate ExcelDoc to "ART DEPARTMENT-NEW:For SQL:FIP:Merged:"
		delete ExcelDoc
	end tell
	
	--saves changes made to new PDF
	
	with timeout of 86400 seconds
		tell application "Adobe Acrobat Pro"
			save document 1 to FinishedFilePath
			close document 1
		end tell
	end timeout
	
	
	set FinishTime to characters 1 through 8 of time string of (current date) as text
	set endSeconds to time of (current date)
	
	set timeDiff to round (((endSeconds - startSeconds) * 100)) / 100
	
	--opens the logs and inputs this jobs info at the end of the log
	
	tell application "Microsoft Excel"
		open "MERGE CENTRAL:FIP AUTOMATION:Logs:FIP Automation Log.xlsx"
		set i to 2
		repeat
			tell row i
				if value of column 1 is "" then
					set value of column 1 to TotalQty
					set value of column 2 to OrderNumber
					set value of column 3 to StartTime
					set value of column 4 to FinishTime
					set value of column 5 to timeDiff
					exit repeat
				end if
				set i to i + 1
			end tell
		end repeat
		close active workbook saving yes
	end tell
	
end repeat