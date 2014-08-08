--these subroutines take a mac path as parameters. they fail if there is nothing in the folder so it checks first and does nothing if folder is empty

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
	set dbdate to text 12 thru 15 of theFilename
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
	-- Get Excel to activate
	activate
	
	-- Close any workbooks that we have open
	close workbooks
	
	-- Ask Excel to open the theFile spreadsheet
	open theFile
	tell active sheet
		tell used range
			set rc to count of rows
		end tell
		copy range range ("A2:A" & rc)
	end tell
end tell

set fileNamesFile to "MERGE CENTRAL:SUPERmerge 2:Databases:File Names.txt"
tell application "TextWrangler"
	open file fileNamesFile
	select every text of document 1
	activate
	paste
	save text document 1 to file fileNamesFile without saving as stationery
	quit
end tell

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
	--switched to using shell scripts to copy/move/delete above. these are still here just in case i need to look at them
	
	--move entire contents of sourceFolder1 to folder targetFolder with replacing
	--move entire contents of sourceFolder2 to folder targetFolder with replacing
	--move entire contents of sourceFolder3 to folder targetFolder with replacing
	--move entire contents of sourceFolder4 to folder targetFolder with replacing
	--move entire contents of sourceFolder5 to folder targetFolder with replacing
	--move entire contents of sourceFolder6 to folder targetFolder with replacing
	--delete entire contents of sourceFolder1
	--delete entire contents of sourceFolder2
	--delete entire contents of sourceFolder3
	--move file theFile to folder "HOM_Shortrun:Databases:SuperMerge Old DBs:"
	
	
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
	set filelist to files of folder POSIX file "/Volumes/HOM_Shortrun/PDFs to process/" as alias list
	repeat with TheItem in filelist
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



