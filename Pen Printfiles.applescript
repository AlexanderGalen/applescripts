set activeJobs to "/Volumes/HOM_Shortrun/~HOM Active Jobs/"
set printFilesFolder to "/Volumes/ART DEPARTMENT-NEW/For HoM PrintFiles/"

tell application "Finder"
	set theFile to selection as string
	set jobnumber to characters 31 thru 36 of theFile as string
end tell

set originalPOSIXFile to quoted form of POSIX path of theFile

set jobFolderPrintfile to quoted form of activeJobs & jobnumber & "/" & jobnumber & ".pen.printfile.eps"
set copiedPrintfile to quoted form of printFilesFolder & jobnumber & ".pen.printfile.eps"
--return originalPOSIXFile & "\n" & jobFolderPrintfile & "\n" & copiedPrintfile
do shell script "cp -Rp " & originalPOSIXFile & " " & jobFolderPrintfile
do shell script "cp -Rp " & originalPOSIXFile & " " & copiedPrintfile
