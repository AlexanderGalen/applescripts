property firstRow : 2

--initializing some variables
set copyvar to true
set ActiveJobs to "HOM_Shortrun:~HOM Active Jobs:"
set DialogButton to ""
set ExitVariable to ""
set myDoc to "HOM_Shortrun:Databases:Supermerge.THIS.txt"
set myFolder to "HOM_Shortrun:~HOM Active Jobs:"
set sourceFolder to "HOM_Shortrun:SUPERmergeIN:CLIENT Images:"
set quarkFolder to "HOM_Shortrun:SUPERmergeOUT:Merged Quark Docs:"
tell application "Microsoft Excel" to open myDoc
set i to firstRow
repeat while ExitVariable is not "Exit"
	
	--first chunk of psuedo repeat loop
	repeat 1 times
		
		--resets dialogbutton's value for each iteration
		set DialogButton to ""
		
		tell application "Microsoft Excel"
			tell row i
				set imageNames to {value of cell 13, value of cell 14, value of cell 15, value of cell 16, value of cell 29, value of cell 30, value of cell 31, value of cell 32, value of cell 33}
				set folderName to string value of cell 1
				set jobnumber to string value of cell 1
				set quarkName to string value of cell 1 & ".HOM.qxp"
				set prevArt to string value of cell 34
				set prevInfo to string value of cell 35
			end tell
		end tell
		
		
		
		
		--does a full exit if it finds no value for job number in database: when it reaches the end of the database
		--the check for 0 is because excel was returning 0 as the value for an empty cell and I don't know why
		
		if jobnumber is "" or jobnumber is 0 then
			set ExitVariable to "Exit"
			exit repeat
		end if
		
		
		
		tell application "Finder"
			if not (exists folder folderName of folder myFolder) then
				make new folder at folder myFolder with properties {name:folderName}
			end if
			duplicate file quarkName of folder quarkFolder to folder folderName of folder myFolder
			
			
			
			--checking values of prev info/art cells; skips copying if they are different values or both empty
			--sets value of prevJob if they are the same or only one has a value
			--its kinda wonky but I know whats going on
			
			
			set prevJob to ""
			set skipvar to ""
			
			
			if prevArt is "" and prevInfo is "" then
				--both blank; skip old job folder copy
				set skipvar to "Skip"
			else
				set prevjobstatus to "Not Both Blank"
				if prevArt is "" or prevInfo is "" then
					--one blank; set whichever is not blank to prevJob
					if prevArt is "" then
						set prevJob to prevInfo
						set skipvar to "NoSkip"
					else
						set prevJob to prevArt
						set skipvar to "NoSkip"
					end if
				else
					--neither blank; check values to see if they are the same
					if prevArt is not prevInfo then
						--neither blank; different values. skip old job folder copy
						set skipvar to "Skip"
					else
						--neither blank; identical values. set value to prevJob
						set prevJob to prevArt
						set skipvar to "NoSkip"
					end if
				end if
			end if
			
			--checks if value of Prevjob is the default value, meaning the customer accidentally entered a useless value. also skips copying for group orders, except for the first one.
			if prevJob is "Enter previous order number…" then --or (character 7 of prevJob is "_" and character 8 of prevJob is not "1") then
				set skipvar to "skip"
			end if
			
			--skips over copying old job if the result of value check determines that it should be skipped
			
			if skipvar is "Skip" then
				exit repeat
			end if
			
			
			
			--this if block checks the length of prevJob from the database, strips HOM from the beginning if it is there, or allows the user
			--to input the previous Job number manually if it is an unsual length
			
			
			if length of prevJob is 9 then
				set prevJob to characters 4 thru 9 of prevJob as string
				
				--if the length of prevJob is not 6 or 9, asks user to enter previous job number manually
			else if length of prevJob is not 6 then
				set DialogVar1 to (display dialog "Previous Job Number is Likely Incorrect.\nEnter It Manually or Skip" & "\n\nOriginal: " & prevJob default answer "" buttons {"Skip", "OK"} default button 2 giving up after 5)
				
				--sets prevJob to the text returned of the dialog and saves returned button as a variable
				set prevJob to text returned of DialogVar1
				set DialogButton to button returned of DialogVar1
			end if
			
			
			--an error check to make sure preJob is a 6 digit integer
			try
				set prevJob to prevJob as integer
			on error
				set DialogVar1 to (display dialog "Previous Job Number is Likely Incorrect.\nEnter It Manually or Skip" & "\n\nOriginal: " & prevJob default answer "" buttons {"Skip", "OK"} default button 2 giving up after 5)
				
				--sets prevJob to the text returned of the dialog and saves returned button as a variable
				set prevJob to text returned of DialogVar1
				set DialogButton to button returned of DialogVar1
			end try
			
			
			
			
			--checks to see if the user selected to skip copying the old job folder for this iteration, or didnt select anything in the dialog
			--adds one to the counter variable, then exits the first inner loop to move to the next iteration, but still copies images to jobfolder
			if DialogButton is "Skip" then
				exit repeat
			end if
			
			
			
			
			
			
			
			--this whole chunk is my (alex's) copy old job to new folder script
			
			set ActiveJobFolder to "HOM_Shortrun:~HOM Active Jobs:" & folderName as string
			set posixActiveJob to POSIX path of ActiveJobFolder
			
			--whole copying process is wrapped in a try block
			--so if an error occurs a dialog is displayed
			try
				
				--this is for if the job is in the older archives
				
				if prevJob is less than 226000 then
					
					set copyvar to false
					set Chars to characters of prevJob
					set First3 to item 1 of Chars & item 2 of Chars & item 3 of Chars as string
					set OldPth to "ARCHIVES VINTAGE:HOM Archive Jobs:" & First3 & "xxx.jobs:"
					set SearchTerm to POSIX path of OldPth & prevJob
					
					-- shell script that finds and copies old job folder to new one
					
					do shell script "cp -Rp \"" & SearchTerm & "\"* " & quoted form of posixActiveJob
					
				end if
				
				--Checks Various folders where old job might be
				
				tell application "Finder"
					
					--checks for old job in HOM Calendars 2006; copies if found
					if exists "HOM_Shortrun:HOM Calendars 2006:" & prevJob then
						set FinishedOldPath to "HOM_Shortrun:HOM Calendars 2006:" & prevJob
						
						--checks if _1 is appended to end of folder name in 2006 calendars
					else if exists "HOM_Shortrun:HOM Calendars 2006:" & prevJob & "_1" then
						set FinishedOldPath to "HOM_Shortrun:HOM Calendars 2006:" & prevJob & "_1"
						
						--checks for old job in HOM Calendars PRINTED; copies if found
					else if exists "HOM_Shortrun:HOM Calendars PRINTED:" & prevJob then
						set FinishedOldPath to "HOM_Shortrun:HOM Calendars PRINTED:" & prevJob
						
						--checks if _1 is appended to end of folder name in HOM Calendars PRINTED
					else if exists "HOM_Shortrun:HOM Calendars PRINTED:" & prevJob & "_1" then
						set FinishedOldPath to "HOM_Shortrun:HOM Calendars PRINTED:" & prevJob & "_1"
						
						--checks for old job in HOM Printed Jobs; copies if found
					else if exists "HOM_Shortrun:~HOM Printed Jobs:" & prevJob then
						set FinishedOldPath to "HOM_Shortrun:~HOM Printed Jobs:" & prevJob
						
						--checks if _1 is appended to end of folder name in Printed
					else if exists "HOM_Shortrun:~HOM Printed Jobs:" & prevJob & "_1" then
						set FinishedOldPath to "HOM_Shortrun:~HOM Printed Jobs:" & prevJob & "_1"
						
						--checks for old job in Active Jobs; copies if found
					else if exists "HOM_Shortrun:~HOM Active Jobs:" & prevJob then
						set FinishedOldPath to "HOM_Shortrun:~HOM Active Jobs:" & prevJob
						
						--checks if _1 is appended to end of folder name in active
					else if exists "HOM_Shortrun:~HOM Active Jobs:" & prevJob & "_1" then
						set FinishedOldPath to "HOM_Shortrun:~HOM Active Jobs:" & prevJob & "_1"
						
						
						--if none of those exist, checks the archives for the folder
						
					else
						
						set copyvar to false
						--converts prevJob back into a string
						set prevJob to prevJob as string
						
						set Chars to characters of prevJob
						set First3 to item 1 of Chars & item 2 of Chars & item 3 of Chars as string
						set OldPth to "HOM_Shortrun:~HOM Archive Jobs:" & First3 & "xxx.jobs:"
						set SearchTerm to POSIX path of OldPth & prevJob
						
						-- shell script that finds and copies old job folder to new one
						
						do shell script "cp -Rp \"" & SearchTerm & "\"* " & quoted form of posixActiveJob
						
					end if
					
					
					--duplicates whichever path was found to exist into the Active Job Folder
					if copyvar is true then
						duplicate FinishedOldPath to ActiveJobFolder
					end if
				end tell
				
				
				--displays a dialog saying copying from old job folder didnt work
			on error
				tell application "System Events"
					set userRespose to button returned of (display dialog "Failed to Copy Previous Job Folder" buttons {"Exit Script", "Skip to Next Row"} giving up after 5)
					if userRespose is "Exit Script" then
						return "User Exited Script"
					else if userRespose is "Skip to Next Row" then
					end if
				end tell
				
			end try
			
			
		end tell
		
		
		--this is the end of the copy old folder to new folder
		
		
		
		
	end repeat
	--end of first pseudo repeat loop and beginning of second
	repeat 1 times
		tell application "Finder"
			
			
			
			--moves image files to Job Folder in Active Jobs
			repeat with anItem in imageNames
				if contents of anItem is not "" then
					try
						
						duplicate file anItem of folder sourceFolder to folder folderName of folder myFolder
						
					on error
						
						beep
						with timeout of 10000 seconds
							tell application "System Events"
								set DialogVar2 to display dialog "An Error Occured During Copying of file \"" & anItem & "\"" buttons {"Try Again", "Ignore and Continue after Error"} default button 1 giving up after 15
							end tell
						end timeout
						
						if button returned of DialogVar2 is "Try Again" then
							try
								duplicate file anItem of folder sourceFolder to folder folderName of folder myFolder
							on error
								set logText to time string of (current date) & "Job Number: " & jobnumber & "File Name: " & anItem as text
								set LogFile to "HOM_Shortrun:Merge Error Log.txt"
								open for access file LogFile with write permission
								write (logText & return) to file LogFile starting at eof
								close access file LogFile
							end try
							
							--if user selected to skip, or does not choose a button, skips the copying and logs error to file
							
						else
							set logText to time string of (current date) & "Job Number: " & jobnumber & "File Name: " & anItem as text
							set LogFile to "HOM_Shortrun:Merge Error Log.txt"
							open for access file LogFile with write permission
							write (logText & return) to file LogFile starting at eof
							close access file LogFile
						end if
						
					end try
				end if
			end repeat
		end tell
		
		set i to i + 1
		
	end repeat
end repeat



--loops through again, and deletes client images from client images folder if
--they were sucessfully copied into the job folder
set ExitVariable to ""
set i to firstRow
repeat while ExitVariable is not "Exit"
	tell application "Microsoft Excel"
		tell row i
			set imageNames to {value of cell 13, value of cell 14, value of cell 15, value of cell 16, value of cell 29, value of cell 30, value of cell 31, value of cell 32, value of cell 33}
			set folderName to string value of cell 1
			set jobnumber to string value of cell 1
			set quarkName to string value of cell 1 & ".HOM.qxp"
		end tell
	end tell
	if jobnumber is "" then exit repeat
	tell application "Finder"
		repeat with anItem in imageNames
			if contents of anItem is not "" then
				if (exists file anItem of folder folderName of folder myFolder) then
					set fileToDelete to sourceFolder & anItem
					set POSIXFileToDelete to quoted form of POSIX path of fileToDelete
					do shell script "rm -d -f " & POSIXFileToDelete
					--delete file anItem of folder sourceFolder
				end if
			end if
		end repeat
	end tell
	set i to i + 1
end repeat



tell application "Microsoft Excel"
	-- Close the worksheet that we've just created
	close active workbook saving no
end tell


--moves unprocessed images from their folders into original client images
tell application "Finder"
	with timeout of 10000 seconds
		set destinationFolder to "HOM_Shortrun:SUPERmergeOUT:Original Client Images:"
		set POSIXDestination to quoted form of POSIX path of destinationFolder
		
		set moveSourceFolder to "HOM_Shortrun:PDFs to process:"
		set POSIXMoveSourceFolder to quoted form of POSIX path of moveSourceFolder
		
		--move entire contents of folder "HOM_Shortrun:PDFs to process" to folder "HOM_Shortrun:SUPERmergeOUT:Original Client Images:" with replacing
		
		do shell script "cp -Rp -f " & POSIXMoveSourceFolder & " " & POSIXDestination
		do shell script "rm -d -f " & POSIXMoveSourceFolder & "*"
		
		--move entire contents of folder "HOM_Shortrun:Process Client Images" to folder "HOM_Shortrun:SUPERmergeOUT:Original Client Images:" with replacing
		
		set moveSourceFolder to "HOM_Shortrun:Process Client Images:"
		set POSIXMoveSourceFolder to quoted form of POSIX path of moveSourceFolder
		
		do shell script "cp -Rp -f " & POSIXMoveSourceFolder & " " & POSIXDestination
		do shell script "rm -d -f " & POSIXMoveSourceFolder & "*"
		
		
	end timeout
end tell