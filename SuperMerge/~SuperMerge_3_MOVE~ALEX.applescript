property firstRow : 2

--these subroutines take a mac path as parameters. they fail if there is nothing in the folder so it checks first and does nothing if folder is empty
--copy functions will replace existing files

on cp_all(source, destination)
	set source to quoted form of POSIX path of source
	set destination to quoted form of POSIX path of destination
	do shell script "cp -Rpf " & source & "* " & destination
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

on logToFile(logText, LogFile)
	set openedFile to open for access file LogFile with write permission
	write (logText & return) to openedFile starting at eof
	close access openedFile
end logToFile

--initializing some variables
set copyvar to true
set LogFile to "HOM_Shortrun:Merge Error Log.txt"
set errorOccured to false
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
		
		--resets value of prevArt, prevInfo, and prevJob to avoid jobs copying previous folder from previous row in DB
		set prevArt to ""
		set prevInfo to ""
		set prevJob to ""
		
		
		tell application "Microsoft Excel"
			tell row i
				set imageNames to {value of cell 13, value of cell 14, value of cell 15, value of cell 16, value of cell 29, value of cell 30, value of cell 31, value of cell 32, value of cell 33}
				set jobNumber to string value of cell 1
				set quarkName to string value of cell 1 & ".HOM.qxp"
				set prevArt to string value of cell 34
				set prevInfo to string value of cell 35
			end tell
		end tell
		
		--sets variables specific to this row
		set thisQuarkDoc to mergedQuarkDocsFolder & quarkName as string
		set thisJobFolder to activeJobsFolder & jobNumber as string
		
		--does a full exit if it finds no value for job number in database: when it reaches the end of the database
		--the check for 0 is because excel was returning 0 as the value for an empty cell and I don't know why
		
		if jobNumber is "" or jobNumber is 0 then
			set ExitVariable to "Exit"
			exit repeat
		end if
		
		
		tell application "Finder"
			if not (exists folder jobNumber of folder activeJobsFolder) then
				make new folder at folder activeJobsFolder with properties {name:jobNumber}
			end if
		end tell
		
		cp(thisQuarkDoc, thisJobFolder)
		
		if prevArt is not "" then
			try
				set prevJob to find text "[0-9]{6}" in prevArt with regexp and string result
				set skipvar to false
				set textToLog to false
			on error errStr number errorNumber
				set textToLog to true
				set logText to (current date) & tab & jobNumber & tab & tab & tab & errStr & tab & errorNumber as text
				set skipvar to true
				set prevJob to ""
			end try
		else if prevInfo is not "" then
			try
				set prevJob to find text "[0-9]{6}" in prevInfo with regexp and string result
				set skipvar to false
				set textToLog to false
			on error errStr number errorNumber
				set textToLog to true
				set logText to (current date) & tab & jobNumber & tab & tab & tab & errStr & tab & errorNumber as text
				set skipvar to true
				set prevJob to ""
			end try
		else
			set skipvar to true
			set textToLog to false
		end if
		
		--skips over copying old job if the result of value check determines that it should be skipped
		
		if textToLog then
			set errorOccured to true
			logToFile(logText, LogFile)
		end if
		
		if skipvar then
			set skipvar to false
			exit repeat
		end if
		
		set testLog to "Macintosh HD:Users:maggie:desktop:testlog.txt"
		set testLogText to "row: " & i & "\tJob Number: " & jobNumber & "\tPrevious Job Number: " & prevJob & "\n"
		logToFile(testLogText, testLog)
		
		--this whole chunk is my (Alex) copy old job to new folder script
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
				log thisSource
				log thisJobFolder
				cp_all(thisSource, thisJobFolder)
				
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
					
					set OldPth to "HOM_Shortrun:~HOM Archive Jobs:" & characters 1 thru 3 of prevJob & "xxx.jobs:" as string
					set FinishedArchivePath to OldPth & prevJob
					
				end if
			end tell
			
			--duplicates whichever path was found to exist into the Active Job Folder
			if copyvar then
				log FinishedOldPath
				log thisJobFolder
				cp(FinishedOldPath, thisJobFolder)
			else
				log FinishedArchivePath
				log thisJobFolder
				cp_all(FinishedArchivePath, thisJobFolder)
			end if
			
		on error errStr number errorNumber
			set errorOccured to true
			set logText to (current date) & tab & jobNumber & tab & tab & prevJob & tab & errStr & tab & errorNumber as text
			logToFile(logText, LogFile)
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
					cp(thisImagePath, thisJobFolder)
				on error errStr number errorNumber
					set errorOccured to true
					set logText to (current date) & tab & jobNumber & tab & thisImage & tab & tab & errStr & tab & errorNumber as text
					logToFile(logText, LogFile)
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
			set jobNumber to string value of cell 1
			set quarkName to string value of cell 1 & ".HOM.qxp"
		end tell
	end tell
	if jobNumber is "" then exit repeat
	
	repeat with thisImage in imageNames
		
		--sets variable specific to this row
		if contents of thisImage is not "" then
			tell application "Finder"
				if (exists file thisImage of folder jobNumber of folder activeJobsFolder) then
					set fileToDelete to clientImagesFolder & thisImage as string
				else
					set fileToDelete to ""
				end if
			end tell
			if fileToDelete is not "" then
				rm(fileToDelete)
			end if
		end if
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
try
	cp_all(pdfsToProcessFolder, originalImagesFolder)
	set clearFail to false
on error
	set clearFail to true
end try
try
	rm_all(pdfsToProcessFolder)
	set clearFail to false
on error
	set clearFail to true
end try
try
	cp_all(imagesToProcessFolder, originalImagesFolder)
	set clearFail to false
on error
	set clearFail to true
end try
try
	rm_all(imagesToProcessFolder)
	set clearFail to false
on error
	set clearFail to true
end try



--displays a dialog to alert user if there were any errors, or exit cleanly if not

if errorOccured then
	tell application "Microsoft Excel"
		open file "HOM_Shortrun:Merge Error Log.txt"
	end tell
	display dialog "Some jobs did not process correctly"
end if

if clearFail then
	display dialog "Process image folders may not have been cleared out successfully"
end if
