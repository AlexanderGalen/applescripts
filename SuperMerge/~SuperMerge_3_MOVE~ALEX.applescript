property firstRow : 2

--these subroutines take a mac path as parameters. they fail if there is nothing in the folder so it checks first and does nothing if folder is empty
--copy functions will replace existing files

on cp_all(source, destination)
	tell application "Finder"
		set theFiles to entire contents of folder source as alias list
	end tell
	if theFiles is not {} then
		set source to quoted form of POSIX path of source
		set destination to quoted form of POSIX path of destination
		do shell script "cp -Rpf " & source & "* " & destination
	end if
end cp_all

on cp(source, destination)
	set source to quoted form of POSIX path of source
	set destination to quoted form of POSIX path of destination
	do shell script "cp -Rpf " & source & " " & destination
end cp

on rm_all(source)
	tell application "Finder"
		set theFiles to entire contents of folder source as alias list
	end tell
	if theFiles is not {} then
		set source to quoted form of POSIX path of source
		do shell script "rm -d -f " & source & "*"
	end if
end rm_all

on rm(source)
	set source to quoted form of POSIX path of source
	do shell script "rm -d -f " & source
end rm

--initializing some variables
set copyvar to true
set ExitVariable to ""
set superMergeDB to "HOM_Shortrun:Databases:Supermerge.THIS.txt"
set activeJobsFolder to "HOM_Shortrun:~HOM Active Jobs:"
set clientImagesFolder to "HOM_Shortrun:SUPERmergeIN:CLIENT Images:"
set mergedQuarkDocsFolder to "HOM_Shortrun:SUPERmergeOUT:Merged Quark Docs:"
tell application "Microsoft Excel" to open superMergeDB
set i to firstRow
repeat while ExitVariable is not "Exit"
	
	--first chunk of psuedo repeat loop
	repeat 1 times
		
		tell application "Microsoft Excel"
			tell row i
				set imageNames to {value of cell 13, value of cell 14, value of cell 15, value of cell 16, value of cell 29, value of cell 30, value of cell 31, value of cell 32, value of cell 33}
				set jobnumber to string value of cell 1
				set quarkName to string value of cell 1 & ".HOM.qxp"
				set prevArt to string value of cell 34
				set prevInfo to string value of cell 35
			end tell
		end tell

		--sets variables specific to this row
		set thisQuarkDoc to mergedQuarkDocsFolder & quarkName as string
		set thisJobFolder to activeJobsFolder & jobnumber as string
		
		--does a full exit if it finds no value for job number in database: when it reaches the end of the database
		--the check for 0 is because excel was returning 0 as the value for an empty cell and I don't know why
		
		if jobnumber is "" or jobnumber is 0 then
			set ExitVariable to "Exit"
			exit repeat
		end if
		
		
		tell application "Finder"
			if not (exists folder jobnumber of folder activeJobsFolder) then
				make new folder at folder activeJobsFolder with properties {name:jobnumber}
			end if
		end tell

			cp(thisQuarkDoc,thisJobFolder)
			
			
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
			if prevJob is "Enter previous order number" then
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
				--previous job number is not 6 or 9 characters long. Script does not yet have a way to deduce a usable job number from this yet. Skips copying this job.
				set skipvar to "Skip"
			end if
			
			--an error check to make sure prevJob is an integer
			try
				set prevJob to prevJob as integer
			on error
				--prevJob cannot be coerced into an integer. Likely contains alpha characters. Script does not yet have a way to deduc a usable job number from this. Skips copying this job.
				set skipvar to "Skip"
			end try
			
			--if script was unable to get a usable previous job number, skips copying the previous job.
			if skipvar is "Skip" then
				exit repeat
			end if
			
			--this whole chunk is my (alex's) copy old job to new folder script
			--whole copying process is wrapped in a try block. If it fails, it just continues on to the next job.
			try
				
				--this is for if the job is in the older archives
				
				if prevJob is less than 226000 then
					
					set copyvar to false
					set Chars to characters of prevJob
					set First3 to item 1 of Chars & item 2 of Chars & item 3 of Chars as string
					set OldPth to "ARCHIVES VINTAGE:HOM Archive Jobs:" & First3 & "xxx.jobs:"
					set thisSource to OldPth & prevJob
					
					-- shell script that finds and copies old job folder to new one
					cp_all(thisSource,thisJobFolder)
					
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
						set thisSource to OldPth & prevJob
						
						-- shell script that finds and copies old job folder to new one
						
						cp(thisSource,thisJobFolder)
						
					end if
				end tell
									
					--duplicates whichever path was found to exist into the Active Job Folder
					if copyvar is true then
						cp(FinishedOldPath,thisJobFolder)
					end if
			on error				
			end try
		
		--end of the copy old folder to new folder
		
		
	end repeat
	--end of first pseudo repeat loop and beginning of second
	repeat 1 times
			
		--moves image files to Job Folder in Active Jobs
		repeat with thisImage in imageNames

			--sets variables specific to this row
			set thisImagePath to clientImagesFolder & thisImage

			if contents of thisImage is not "" then
				try
					cp(thisImagePath,thisJobFolder)					
				on error
					
					set logText to time string of (current date) & "Job Number: " & jobnumber & "File Name: " & thisImage as text
					set LogFile to "HOM_Shortrun:Merge Error Log.txt"
					open for access file LogFile with write permission
					write (logText & return) to file LogFile starting at eof
					close access file LogFile
					
				end try
			end if
		end repeat
		
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
			set jobnumber to string value of cell 1
			set quarkName to string value of cell 1 & ".HOM.qxp"
		end tell
	end tell
	if jobnumber is "" then exit repeat
	
		repeat with thisImage in imageNames

		--sets variable specific to this row
			tell application "Finder"
				if contents of thisImage is not "" then
					if (exists file thisImage of folder jobnumber of folder activeJobsFolder) then
						set fileToDelete to clientImagesFolder & thisImage as string
						rm(fileToDelete)
					end if
				end if
			end tell
		end repeat
	set i to i + 1
end repeat

tell application "Microsoft Excel"
	close active workbook saving no
end tell

--moves unprocessed images from their folders into original client images
set originalImagesFolder to "HOM_Shortrun:SUPERmergeOUT:Original Client Images:"
set pdfsToProcessFolder to "HOM_Shortrun:PDFs to process:"
set imagesToProcessFolder to "HOM_Shortrun:Process Client Images:"
cp_all(pdfsToProcessFolder,originalImagesFolder)
rm_all(pdfsToProcessFolder)
cp_all(imagesToProcessFolder,originalImagesFolder)
rm_all(imagesToProcessFolder)