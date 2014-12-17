--function to replace "oldItem" with "newItem" in "someText" and return the new string
on replaceText(someText, oldItem, newItem)
	
	set {tempTID, AppleScript's text item delimiters} to {AppleScript's text item delimiters, oldItem}
	try
		set {itemList, AppleScript's text item delimiters} to {text items of someText, newItem}
		set {someText, AppleScript's text item delimiters} to {itemList as text, tempTID}
	on error errorMessage number errorNumber -- oops
		set AppleScript's text item delimiters to tempTID
		error errorMessage number errorNumber -- pass it on
	end try
	return someText
end replaceText

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

set sourceFolder to "HOM_shortrun:Databases:SuperMergeDB New"

tell application "Finder"
	sort (get files of folder sourceFolder) by creation date
	-- This raises an error if the folder doesn't contain any files
	set theFile to (last item of result) as alias
	set theFilename to the name of theFile as string
	set dbdate to text 11 thru 14 of theFilename
	set mergeFolder to "MERGE CENTRAL:SUPERmerge 2:Merges:"
	set mergeFOlderName to dbdate & "2014_merge"
	make new folder at folder mergeFolder with properties {name:mergeFOlderName}
	set theTemplate to "MERGE CENTRAL:SUPERmerge 2:Merges:x_HOM_MergeTHIS.qxp"
	set thisMergeFolder to mergeFolder & mergeFOlderName
end tell

--copies template to merge folder
cp(theTemplate, thisMergeFolder)

tell application "Finder"
	get files of folder mergeFOlderName of folder mergeFolder
	set newTemplate to (item 1 of result) as alias
end tell

tell application "Microsoft Excel"
	activate
	close workbooks
	open theFile
	tell active sheet
		tell used range
			set rc to count of rows
		end tell
		set fileNames to value of range ("A2:A" & rc)
		--loops through and replaces new lines with spaces for every value in the art instructions column.
		set r to 2
		repeat while r is less than rc
			tell row r
				set thisText to value of cell 36
				set newText to my replaceText(thisText, "\r", " ")
				set value of cell 36 to newText
			end tell
			set r to r + 1
		end repeat
		
	end tell
end tell

--builds file names text to write to file
set fileNamesText to ""
repeat with thisItem in fileNames
	set thisItem to thisItem as string
	set fileNamesText to fileNamesText & thisItem & "\n"
end repeat

set fileNamesFile to "MERGE CENTRAL:SUPERmerge 2:Databases:File Names.txt"
set openedFile to open for access file fileNamesFile with write permission
set eof of openedFile to 0
write fileNamesText to openedFile starting at eof
close access openedFile

tell application "Microsoft Excel"
	-- Set the current worksheet to our loop position
	set theWorksheet to active sheet of active workbook
	--activate object theWorksheet
	
	-- Save the worksheet as a CSV file
	set theSheetsPath to "HOM_Shortrun:Databases:SuperMerge.THIS.txt" as string
	save as theWorksheet filename theSheetsPath file format text Mac file format with overwrite
	
	-- Close the worksheet that we've just created
	close active workbook saving no
	
	delay 4
	
	-- Clean up and close files
	close workbooks
	
	open "HOM_Shortrun:Databases:SuperMerge.THIS.txt"
	select range ("AJ2:AJ" & rc)
end tell

tell application "Finder"
	--set backupFolder to folder "HOM_Shortrun:SUPERmergeOUT:Original Client Images"
	set sourceFolder1 to "switchdata:Uploader2:Photo:"
	set sourceFolder2 to "switchdata:Uploader2:Logo:"
	set sourceFolder3 to "switchdata:Uploader2:Misc:"
	set sourceFolder4 to "HOM_Shortrun:CLIENT LOGOS:"
	set sourceFolder5 to "HOM_Shortrun:CLIENT MISC:"
	set sourceFolder6 to "HOM_Shortrun:CLIENT PHOTOS:"
	set targetFolder to "HOM_Shortrun:Process Client Images:"
	set PDFTargetFolder to "HOM_Shortrun:PDFs to process:"
	--move entire contents of targetFolder to folder backupFolder
	
end tell

--moves all image files to the images to process folder
--doesn't work if nothing is in the folders, so It checks first

cp_all(sourceFolder1, targetFolder)
cp_all(sourceFolder2, targetFolder)
cp_all(sourceFolder3, targetFolder)
cp_all(sourceFolder4, targetFolder)
cp_all(sourceFolder5, targetFolder)
cp_all(sourceFolder6, targetFolder)

set theFile to theFile as string
set cpDest to "HOM_Shortrun:Databases:SuperMerge Old DBs:"
cp(theFile, cpDest)


rm_all(sourceFolder1)
rm_all(sourceFolder2)
rm_all(sourceFolder3)
rm_all(sourceFolder4)
rm_all(sourceFolder5)
rm_all(sourceFolder6)

rm(theFile)




tell application "Finder"
	
	--converts targetfolder to an alias, then sets a filelist for looping through.
	set targetFolder to targetFolder as alias
	set FileList2 to (files of entire contents of targetFolder) as alias list
	
	
	--moves pdfs from target folder to pdfs to process folder
end tell
repeat with TheItem in FileList2
	tell application "Finder"
		set {name:FileName, name extension:fileExtension} to TheItem
	end tell
	ignoring case
		if fileExtension is "pdf" then
			cp(TheItem, PDFTargetFolder)
			rm(TheItem)
			--move TheItem to PDFTargetFolder
		end if
	end ignoring
end repeat


tell application "QuarkXPress"
	open newTemplate
end tell






--Processes PDFS that were separated from images; puts them into Client Images

set ProcessedImagesFolder to "HOM_Shortrun:SUPERMergeIN:CLIENT Images:"
tell application "Finder"
	set fileList to files of folder POSIX file "/Volumes/HOM_Shortrun/PDFs to process/" as alias list
	repeat with TheItem in fileList
		set {name:FileName, name extension:fileExtension} to TheItem
		if name extension of TheItem is "pdf" then
			try
				set docNameToSaveTo to text 1 thru -5 of FileName & ".eps"
				set docToSavePath to ProcessedImagesFolder & docNameToSaveTo
				tell application "Adobe Acrobat Pro"
					open file (TheItem as text)
					activate
					save document 1 to file docToSavePath using EPS Conversion --with embedded fonts, halftones and TrueType without binary, annotation, images and preview
					close document 1 saving no
				end tell
			on error errMsg number errNum
				display dialog "An error occurred: " & errNum & " - " & errMsg buttons {"Cancel", "OK"} default button "OK"
			end try
		end if
	end repeat
end tell


ignoring application responses
	tell application "Adobe Photoshop CS6"
		set batchScript to "Macintosh HD:Applications:Adobe Photoshop CS6:Presets:Scripts:SM Batch Resize.jsx"
		do javascript file batchScript
	end tell
end ignoring



