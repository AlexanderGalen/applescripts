--annotations for a this script(well, one that is very similar at least) can be found in the script "Resource:Scripting:AAASpectacularAlex Scripts:Copy Old Job to New Folder"

tell application "Keyboard Maestro Engine"
	set myVar1 to make variable with properties {name:"Job Number"}
	set NewJN to value of myVar1
	set myVar2 to make variable with properties {name:"Previous Job Number"}
	set OldJN to value of myVar2
end tell
set NewPth to "HOM_Shortrun:~HOM Active Jobs:" & NewJN as string


--Makes a new folder in active jobs for the current job, and opens a new finder window displaying it
tell application "Finder"
	if not (exists "HOM_Shortrun:~HOM Active Jobs:" & NewJN & ":") then
		try
			make new folder at folder "HOM_Shortrun:~HOM Active Jobs:" with properties {name:NewJN}
		on error
			tell application "System Events"
				display dialog "Could not make new folder. Script is exiting"
			end tell
			return "script exited because it encountered an error"
		end try
	end if
	set JobFolder to "HOM_Shortrun:~HOM Active Jobs:" & NewJN & ":"
	set NewWindow to make new Finder window
	set properties of NewWindow to {target:JobFolder, position:{-635, 78}, bounds:{-635, 78, 0, 568}, current view:column view, sidebar width:205}
end tell


--whole copying process is wrapped in a try block
--so if an error occurs a dialog is displayed
try
	
	--this is for if the job is in the older archives
	
	if OldJN is less than 226000 then
		mount volume "smb://arc/ARCHIVES VINTAGE"
		set Chars to characters of OldJN
		set First3 to item 1 of Chars & item 2 of Chars & item 3 of Chars as string
		set OldPth to "ARCHIVES VINTAGE:HOM Archive Jobs:" & First3 & "xxx.jobs:"
		set SearchTerm to POSIX path of OldPth & OldJN
		
		-- shell script that finds and copies old job folder to new one
		
		do shell script " cp -Rp \"" & SearchTerm & "\"* \"/volumes/HOM_Shortrun/~HOM Active Jobs/" & NewJN & "/\""
		
	end if
	
	--Checks Various folders where old job might be
	
	tell application "Finder"
		
		--checks for old job in HOM Calendars 2006; copies if found
		if exists "HOM_Shortrun:HOM Calendars 2006:" & OldJN then
			duplicate "HOM_Shortrun:HOM Calendars 2006:" & OldJN to NewPth
			
			--checks if _1 is appended to end of folder name in 2006 calendars
		else if exists "HOM_Shortrun:HOM Calendars 2006:" & OldJN & "_1" then
			duplicate "HOM_Shortrun:HOM Calendars 2006:" & OldJN & "_1" to NewPth
			
			
			--checks for old job in HOM Calendars PRINTED; copies if found
		else if exists "HOM_Shortrun:HOM Calendars PRINTED:" & OldJN then
			duplicate "HOM_Shortrun:HOM Calendars PRINTED:" & OldJN to NewPth
			
			--checks if _1 is appended to end of folder name in HOM Calendars PRINTED
		else if exists "HOM_Shortrun:HOM Calendars PRINTED:" & OldJN & "_1" then
			duplicate "HOM_Shortrun:HOM Calendars PRINTED:" & OldJN & "_1" to NewPth
			
			--checks for old job in HOM Printed Jobs; copies if found
		else if exists "HOM_Shortrun:~HOM Printed Jobs:" & OldJN then
			duplicate "HOM_Shortrun:~HOM Printed Jobs:" & OldJN to NewPth
			
			--checks if _1 is appended to end of folder name in Printed
		else if exists "HOM_Shortrun:~HOM Printed Jobs:" & OldJN & "_1" then
			duplicate "HOM_Shortrun:~HOM Printed Jobs:" & OldJN & "_1" to NewPth
			
			--checks for old job in Active Jobs; copies if found
		else if exists "HOM_Shortrun:~HOM Active Jobs:" & OldJN then
			duplicate "HOM_Shortrun:~HOM Active Jobs:" & OldJN to NewPth
			
			--checks if _1 is appended to end of folder name in active
		else if exists "HOM_Shortrun:~HOM Active Jobs:" & OldJN & "_1" then
			duplicate "HOM_Shortrun:~HOM Active Jobs:" & OldJN & "_1" to NewPth
			
			
			
			
			--if none of those work, checks the archives for the folder
		else
			
			set Chars to characters of OldJN
			set First3 to item 1 of Chars & item 2 of Chars & item 3 of Chars as string
			set OldPth to "HOM_Shortrun:~HOM Archive Jobs:" & First3 & "xxx.jobs:"
			set SearchTerm to POSIX path of OldPth & OldJN
			
			-- shell script that finds and copies old job folder to new one
			
			do shell script " cp -Rp \"" & SearchTerm & "\"* \"/volumes/HOM_Shortrun/~HOM Active Jobs/" & NewJN & "/\""
			
		end if
	end tell
	
	--gets the path to the old job copied to new folder and saves it to Keyboard Maestro
	set shellScriptString to quoted form of "/Volumes/HOM_shortrun/~Hom Active Jobs/" & NewJN & "/" & OldJN & "*/" & OldJN & "*HOM.qxp"
	set copiedPath to do shell script "find " & shellScriptString
	set copiedPath to POSIX file copiedPath
	
	tell application "Keyboard Maestro Engine"
		make new variable with properties {name:"copied path", value:copiedPath}
	end tell
	
	--if even that fails, Displays a Dialog stating that the script failed
on error
	
	tell application "System Events"
		display dialog "Could Not Copy Folder"
		activate
		return
	end tell
	
end try