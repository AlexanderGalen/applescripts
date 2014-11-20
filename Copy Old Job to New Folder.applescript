--gets variable values from Keyboard Maestro

tell application "Keyboard Maestro Engine"
	set myVar2 to make variable with properties {name:"Previous Job Number"}
	set OldJN to value of myVar2
end tell

--for determining if finder should copy the folder, or if it has already been copied with the shell script.
set copyvar to true

--sets the path to the current selection as a variable
--also checks for _1 after job number and corrects accordingly

tell application "Finder"
	set userSelection to selection as string
end tell
set ActiveJobFolder to characters 1 through 36 of userSelection & ":" as string
tell application "Finder"
	if exists ActiveJobFolder then
	else
		set ActiveJobFolder to characters 1 through 36 of userSelection & "_1:" as string
	end if
end tell

set posixActiveJob to POSIX path of ActiveJobFolder

--whole copying process is wrapped in a try block
--so if an error occurs a dialog is displayed

try
	
	--this is for if the job is in the older archives
	
	if OldJN is less than 226000 then
		mount volume "smb://arc/ARCHIVES VINTAGE"
		set copyvar to false
		set Chars to characters of OldJN
		set First3 to item 1 of Chars & item 2 of Chars & item 3 of Chars as string
		set OldPth to "ARCHIVES VINTAGE:HOM Archive Jobs:" & First3 & "xxx.jobs:"
		set SearchTerm to POSIX path of OldPth & OldJN
		
		-- shell script that finds and copies old job folder to new one
		do shell script "cp -Rp \"" & SearchTerm & "\"* " & quoted form of posixActiveJob

		
	else
	
	--Checks Various folders where old job might be
	
		tell application "Finder"
			
			--checks for old job in HOM Calendars 2006; copies if found
			if exists "HOM_Shortrun:HOM Calendars 2006:" & OldJN then
				set FinishedOldPath to "HOM_Shortrun:HOM Calendars 2006:" & OldJN
				
				--checks if _1 is appended to end of folder name in 2006 calendars
			else if exists "HOM_Shortrun:HOM Calendars 2006:" & OldJN & "_1" then
				set FinishedOldPath to "HOM_Shortrun:HOM Calendars 2006:" & OldJN & "_1"
				
				
				--checks for old job in HOM Calendars PRINTED; copies if found
			else if exists "HOM_Shortrun:HOM Calendars PRINTED:" & OldJN then
				set FinishedOldPath to "HOM_Shortrun:HOM Calendars PRINTED:" & OldJN
				
				--checks if _1 is appended to end of folder name in HOM Calendars PRINTED
			else if exists "HOM_Shortrun:HOM Calendars PRINTED:" & OldJN & "_1" then
				set FinishedOldPath to "HOM_Shortrun:HOM Calendars PRINTED:" & OldJN & "_1"
				
				--checks for old job in HOM Printed Jobs; copies if found
			else if exists "HOM_Shortrun:~HOM Printed Jobs:" & OldJN then
				set FinishedOldPath to "HOM_Shortrun:~HOM Printed Jobs:" & OldJN
				
				--checks if _1 is appended to end of folder name in Printed
			else if exists "HOM_Shortrun:~HOM Printed Jobs:" & OldJN & "_1" then
				set FinishedOldPath to "HOM_Shortrun:~HOM Printed Jobs:" & OldJN & "_1"
				
				--checks for old job in Active Jobs; copies if found
			else if exists "HOM_Shortrun:~HOM Active Jobs:" & OldJN then
				set FinishedOldPath to "HOM_Shortrun:~HOM Active Jobs:" & OldJN
				
				--checks if _1 is appended to end of folder name in active
			else if exists "HOM_Shortrun:~HOM Active Jobs:" & OldJN & "_1" then
				set FinishedOldPath to "HOM_Shortrun:~HOM Active Jobs:" & OldJN & "_1"
				
				--copies whichever file is found to the current job folder
				
				
				--if none of those exist, checks the archives for the folder
				
			else
				set copyvar to false
				set Chars to characters of OldJN
				set First3 to item 1 of Chars & item 2 of Chars & item 3 of Chars as string
				set OldPth to "HOM_Shortrun:~HOM Archive Jobs:" & First3 & "xxx.jobs:"
				set SearchTerm to POSIX path of OldPth & OldJN
				
				-- shell script that finds and copies old job folder to new one
				
				do shell script "cp -Rp \"" & SearchTerm & "\"* " & quoted form of posixActiveJob
				
			end if
			
			if copyvar is true then
				duplicate FinishedOldPath to ActiveJobFolder
			end if
			
		end tell
		
	end if
	
	
	
	--displays a dialog saying the thing didnt work
on error errStr number errorNumber

	tell application "System Events"
		display dialog "Could Not Copy Folder\n" & errStr & " " & errorNumber
		activate
		return "script exited because it encountered an error"
	end tell
	
end try