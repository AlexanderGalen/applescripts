tell application "Finder"
	set theTarget to target of Finder window 1 as string
	if theTarget does not contain "HOM_Shortrun:~HOM Active Jobs:" then
		display dialog "Selection must be inside of a job Folder in Active Jobs. Script will now exit."
		return
	end if
	set theFiles to selection as alias list
end tell

set thisFile to item 1 of theFiles as string
set jobNumber to characters 31 thru 36 of thisFile as string
set jobFolder to "HOM_Shortrun:~HOM Active Jobs:" & jobNumber & ":"
set infoFile to (path to home folder) & ".jobInfo" as string

set fileCount to count theFiles

set JSONText to "{ \"jobNumber\":" & jobNumber & ", \"files\":["
repeat with thisFile in theFiles
	set thisFile to POSIX path of thisFile as string
	set JSONText to JSONText & "\"" & thisFile & "\","
end repeat

--removes extra trailing comma from JSON Text
set JSONText to characters 1 thru ((count of JSONText) - 1) of JSONText as string

set JSONText to JSONText & "]}"

set thisInfoFile to open for access file infoFile with write permission
set eof of thisInfoFile to 0
write JSONText to thisInfoFile starting at eof
close access thisInfoFile

do shell script "/usr/local/bin/node ${HOME}/test.js"