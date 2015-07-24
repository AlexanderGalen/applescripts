-- ######################################################
-- DesignMerge Automation Process
-- 
-- This process will be implemented via an AppleScript that is initially triggered using
-- the normal Macintosh "folder action" method.
-- 
-- The file that is submitted to the watched folder will be a tab-separated
-- text file designed to be processed by the script and also by DesignMerge
-- (a/k/a the database file). The database file will consist of a header row
-- (with field names) followed by one or more rows of tab-separated data values.
-- 
-- In additional to the variable items to be merged into a template, the
-- database will also hold several values that will be required by the
-- script for processing. In this respect, the database file serves both as
-- a means for passing through variable data, and also for passing through
-- script parameters.
-- 
-- The uniquely named fields in the data file will be as follows:
-- 
-- Template: Holds the full path to the QuarkXPress document to be merged.
-- 
-- DDF Name: Holds the name of the DDF to use, or if empty use the DDF name
-- assigned to the template document.
-- 
-- Output Path: Full path (including filename (based on the Product Job ID)) to the
-- final PDF file to be produced.
-- 
-- Log File Path: Full path (including filename) to a log file that will be
-- created by the script documenting success or failure of each record.
-- 
-- Output Style: Name of the PDF Output Style to use. This must have been
-- previously defined in QuarkXPress.
-- 
-- The DesignMerge Database Definition must have a set of Variable Link
-- names defined to match the field names specified above. The script will
-- query DesignMerge to obtain the value for each of the above items.
-- 
-- When the folder action is triggered, it will launch a script developed by
-- Meadows. The script will open the template document, perform a merge
-- using the database file, and then export the template document to the
-- output path specified in the first record of the data file. This will be repeated for
-- each record of the data file, resulting in (potentially) a unique PDF being
-- produced for each data record.
-- 
-- If a pdf already exists with the same name it will be over-written.
-- 
-- Once all records have merged a log will be created documenting success and failure
-- of each record.
-- 
-- DesignMerge will process each DB file in the order that they are created and delete
-- when completed.
-- 
-- ######################################################
-- Copyright 2013 Meadows Information Systems, LLC
-- All rights reserved worldwide.
-- Provided to Document Business Solutions under non-exclusive license.
-- ######################################################

property kDebugMode : true
property kShowDialogs : true

global gDebugMode
global gShowDialogs
global gLocalLogFilePath
global gLogFilePath
global gDDFName
global gMasterDDFName
global gDatabasePath
global gDatabasePathAlias
global gDatabaseName
global gTemplatePath
global gTemplateName
global gOutputPath
global gOutputPathAlias
global gPDFOutputStyle
global gErrorMessage

property kMasterDDFName : "Graphic Business Solutions"
property kDefaultPDFPresetName : "Graphic Business Solutions"

-- Predefined Variable Link Names
property kTemplateLinkName : "Template"
property kDDFLinkName : "DDF Name"
property kOutputPathLinkName : "Output Path"
property kLogFilePathLinkName : "Log File Path"
property kOutputStyleLinkName : "Output Style"
property kTab : (ASCII character 9)

set gShowDialogs to kShowDialogs -- Will show all error messages for 1 second
set gDebugMode to kDebugMode -- Will prompt to choose folder

if gDebugMode is equal to true then
	tell application "Finder"
		set this_folder to choose folder with prompt "Please select directory."
		my MPSStartFolderAction(this_folder)
	end tell
end if

-- // --------------------------------------------------------------------
on adding folder items to this_folder after receiving these_items
	my MPSStartFolderAction(this_folder)
end adding folder items to

-- // --------------------------------------------------------------------
on MPSStartFolderAction(this_folder)
	set is_completed to false
	
	-- Initialize globals
	set gLocalLogFilePath to ""
	set gLogFilePath to ""
	set gMasterDDFName to kMasterDDFName
	set gDDFName to ""
	set gTemplatePath to ""
	set gPDFOutputStyle to ""
	set gDatabasePath to ""
	set gDatabaseName to ""
	set gTemplatePath to ""
	set gTemplateName to ""
	set gOutputPath to ""
	set gShowDialogs to kShowDialogs -- Will show all error messages for 1 second
	set gDebugMode to kDebugMode -- Will prompt to choose folder
	
	-- Set up local log file.
	set gLocalLogFilePath to (this_folder as string)
	set gLocalLogFilePath to gLocalLogFilePath & "local.log"
	set theResult to my MPSFilePathExists(gLocalLogFilePath)
	if theResult is false then
		my MPSLogFileStartNewLog(gLocalLogFilePath)
	end if
	repeat while is_completed is equal to false
		set theInputFile to my MPSGetNextInputFile(this_folder)
		if (theInputFile is not equal to null) then
			tell application "Finder"
				set gDatabasePath to theInputFile
				set gDatabasePathAlias to (gDatabasePath as alias)
			end tell
			my MPSStartMerge(gDatabasePath)
			try
				tell application "Finder"
					try
						delete file gDatabasePathAlias
					on error
						-- If some error deleting the file, we've got to rename it so we stop processing it.
						set oldDelims to AppleScript's text item delimiters
						set AppleScript's text item delimiters to {"."}
						try
							set fileName to name of gDatabasePathAlias
							set nameWithoutExtension to first text item of fileName
							set newName to nameWithoutExtension & ".done"
							set name of gDatabasePathAlias to newName
							set AppleScript's text item delimiters to oldDelims
						on error
							set AppleScript's text item delimiters to oldDelims
						end try
					end try
				end tell
			end try
		else
			set is_completed to true
		end if
	end repeat
end MPSStartFolderAction

-- // --------------------------------------------------------------------
on MPSGetNextInputFile(item_list)
	local item_path
	local item_list
	local the_item
	local item_info
	local the_items
	local item_count
	local basePath
	local start_offset
	local end_offset
	local the_ext
	local i
	
	set item_path to null
	set the the_items to list folder item_list without invisibles
	set basePath to item_list as string
	set item_count to (get count of items in the_items)
	
	repeat with i from 1 to item_count
		set the_item to item i of the the_items as text
		set start_offset to ((offset of "." in the_item) + 1)
		if (start_offset > 0) then
			set end_offset to count of items in the_item
			if (end_offset > start_offset) then
				set the_ext to text start_offset thru end_offset of the_item
				if the_ext is in {"csv", "txt"} then
					set the item_path to basePath & the_item as string
					return item_path
				end if
			end if
		end if
	end repeat
	return item_path
end MPSGetNextInputFile

-- // --------------------------------------------------------------------
on mainProcess(theDatabaseFile)
	display dialog theDatabaseFile
	tell application "Finder"
		delete file theDatabaseFile
		empty trash
	end tell
end mainProcess

-- // --------------------------------------------------------------------
on MPSSetUpDatabase(theDatabasePath, theDDFName)
	local theResult
	local recordCount
	
	set theResult to null
	
	tell application "QuarkXPress"
		try
			«event MdwsDMii»
			«event MdwsDMDF» gMasterDDFName -- Select this one first in case the other one is not available.
			«event MdwsDMDF» given «class ddfN»:theDDFName
		end try
		set theDatabaseAlias to theDatabasePath as alias
		set gDatabasePath to theDatabasePath
		set theResult to «event MdwsDMDB» given «class dbNm»:theDatabaseAlias
		if theResult is equal to null or item 1 of theResult is less than 0 then
			set gErrorMessage to theDatabasePath & kTab & "N/A" & kTab & "Error code " & (item 1 of theResult) & " returned while selecting database file. Unable to continue." & kTab & ((current date) as string)
			return -1
		end if
	end tell
	
	set recordCount to item 2 of theResult
	if (recordCount ≤ 1) then
		set gErrorMessage to theDatabasePath & kTab & "N/A" & kTab & "Error processing database file. The database appears to be empty. Unable to continue." & kTab & ((current date) as string)
		return -1
	end if
	
	return recordCount
	
end MPSSetUpDatabase

-- // --------------------------------------------------------------------
on MPSSetUpGlobals(curRecord)
	local theList
	local theLinkName
	local theResult
	tell application "QuarkXPress"
		activate
		
		-- Get the template to process
		set theLinkName to kTemplateLinkName
		set theList to («event MdwsDM05» given «class dmLI»:curRecord, «class dmFN»:theLinkName)
		if item 1 of theList is not equal to 0 then
			set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "Error code " & (item 1 of theList) & " was returned for Variable Link " & theLinkName & ". Unable to process record." & kTab & ((current date) as string)
			return -1
		else
			set gTemplatePath to item 2 of theList
		end if
		-- Check if template exists.
		set theResult to my MPSFilePathExists(gTemplatePath)
		if theResult is false then
			set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "The specified template path " & gTemplatePath & " does not exist. Unable to process record." & kTab & ((current date) as string)
			return -1
		end if
		
		-- Output path.
		set theLinkName to kOutputPathLinkName
		set theList to («event MdwsDM05» given «class dmLI»:curRecord, «class dmFN»:theLinkName)
		if item 1 of theList is not equal to 0 then
			set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "Error code " & (item 1 of theList) & " was returned for Variable Link " & theLinkName & ". Unable to process record." & kTab & ((current date) as string)
			return -1
		else
			set gOutputPath to item 2 of theList
		end if
		
		-- Check if path is valid.
		if (gOutputPath is equal to null) or ((count of items in gOutputPath) is equal to 0) then
			set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "The specified output path is contained an empty field. Unable to process record." & kTab & ((current date) as string)
			return -1
		end if
		
		-- Attempt to create a temp version of the file to verify the path is valid
		try
			set theResult to my MPSLogFileCreate(gOutputPath)
			if theResult is equal to 0 then
				set gOutputPathAlias to (gOutputPath as alias)
				tell application "Finder"
					delete file gOutputPath
				end tell
			end if
		on error
			set theResult to -1
		end try
		if theResult is less than 0 then
			set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "The specified output path \"" & gOutputPath & "\" does not appear to be valid. Cannot create alias using this path. Unable to process record." & kTab & ((current date) as string)
			return -1
		end if
		
		-- Log file path
		set theLinkName to kLogFilePathLinkName
		set theList to («event MdwsDM05» given «class dmLI»:curRecord, «class dmFN»:theLinkName)
		if item 1 of theList is not equal to 0 then
			set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "Error code " & (item 1 of theList) & " was returned for Variable Link " & theLinkName & ". Unable to process record." & kTab & ((current date) as string)
			return -1
		else
			set gLogFilePath to item 2 of theList
		end if
		
		-- Check if path is valid.
		if (gLogFilePath is equal to null) or ((count of items in gLogFilePath) is equal to 0) then
			set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "The specified log file path contained an empty field. Default log file path will be used instead." & kTab & ((current date) as string)
			set gLogFilePath to gLocalLogFilePath
		else
			set theResult to my MPSFilePathExists(gLogFilePath)
			if theResult is false then
				set theResult to my MPSLogFileStartNewLog(gLogFilePath)
				if theResult is less than 0 then
					set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "The specified log file path " & gLogFilePath & " does not appear to be valid. Default log file path will be used instead." & kTab & ((current date) as string)
					set gLogFilePath to gLocalLogFilePath
				end if
			end if
		end if
		
		-- PDF Preset Name
		set theLinkName to kOutputStyleLinkName
		set theList to («event MdwsDM05» given «class dmLI»:curRecord, «class dmFN»:theLinkName)
		if item 1 of theList is not equal to 0 then
			set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "Error code " & (item 1 of theList) & " was returned for Variable Link " & theLinkName & ". Unable to process record." & kTab & ((current date) as string)
			return -1
		else
			set gPDFOutputStyle to item 2 of theList
		end if
		
		-- Check if preset is valid.
		if (gPDFOutputStyle is equal to null) or ((count of items in gPDFOutputStyle) is equal to 0) then
			set gPDFOutputStyle to kDefaultPDFPresetName
		end if
		
		-- DDF Name
		set theLinkName to kDDFLinkName
		set theList to («event MdwsDM05» given «class dmLI»:curRecord, «class dmFN»:theLinkName)
		if item 1 of theList is not equal to 0 then
			set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "Error code " & (item 1 of theList) & " was returned for Variable Link " & theLinkName & ". Unable to process record." & kTab & ((current date) as string)
			return -1
		else
			set gDDFName to item 2 of theList
		end if
		
		-- Check if DDF is valid.
		if (gDDFName is equal to null) or ((count of items in gDDFName) is equal to 0) then
			set gDDFName to kMasterDDFName
		end if
	end tell
	return 0
	
end MPSSetUpGlobals

-- // --------------------------------------------------------------------
on MPSStartMerge(theDatabasePath)
	tell application "QuarkXPress"
		activate
		
		local theResult
		local theRecordCount
		local curRecord
		
		-- Get count of records in data file.
		set theRecordCount to my MPSSetUpDatabase(theDatabasePath, gMasterDDFName)
		if (theRecordCount < 0) then
			my MPSLogFileWrite(gLocalLogFilePath, gErrorMessage, 1)
			if gShowDialogs is equal to true then
				display dialog gErrorMessage giving up after 1
			end if
			return -1
		end if
		
		set gErrorMessage to gDatabasePath & kTab & "---" & kTab & "Start processing, total of " & theRecordCount & " records." & kTab & ((current date) as string)
		my MPSLogFileWrite(gLocalLogFilePath, gErrorMessage, 1)
		
		-- Process one record at a time.
		set curRecord to 2 -- Account for skip first record.
		repeat while curRecord is less than or equal to theRecordCount
			repeat 1 times -- Simulated "continue" loop, so we can use "exit repeat"
				
				-- Set up globals, including document path, log file, and PDF Preset
				set gErrorMessage to ""
				set theResult to my MPSSetUpGlobals(curRecord)
				if (gErrorMessage is not equal to "") then
					my MPSLogFileWrite(gLogFilePath, gErrorMessage, 1)
					
					if (gLogFilePath is not equal to gLocalLogFilePath) then
						my MPSLogFileWrite(gLocalLogFilePath, gErrorMessage, 1)
					end if
					
					if gShowDialogs is equal to true then
						display dialog gErrorMessage giving up after 1
					end if
					
					
				end if
				if theResult < 0 then
					exit repeat
				end if
				
				-- Open the document
				set oldDocCount to count documents
				set gTemplateAlias to (gTemplatePath as alias)
				try
					open gTemplateAlias use doc prefs "yes" with Suppress All Warnings
				end try
				set newDocCount to count documents
				if oldDocCount is equal to newDocCount then
					set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "Error opening template " & gTemplatePath & ". Unable to process record." & kTab & ((current date) as string)
					my MPSLogFileWrite(gLogFilePath, gErrorMessage, 1)
					if (gLogFilePath is not equal to gLocalLogFilePath) then
						my MPSLogFileWrite(gLocalLogFilePath, gErrorMessage, 1)
					end if
					if gShowDialogs is equal to true then
						display dialog gErrorMessage giving up after 1
					end if
					
					
					exit repeat
				end if
				
				-- Document is open. Start processing.
				set curDoc to front document
				set curDocName to name of curDoc
				set curDocPath to file path of curDoc
				set curDocAlias to (curDocPath as alias)
				
				-- Re-select the data file if the DDF is different from the master DDF
				if gDDFName is not equal to gMasterDDFName then
					my MPSSetUpDatabase(theDatabasePath, gDDFName)
				end if
				
				set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "Starting DesignMerge merge process." & kTab & ((current date) as string)
				my MPSLogFileWrite(gLocalLogFilePath, gErrorMessage, 1)
				
				-- Initialize DesignMerge and index selected database.
				try
					«event MdwsDMii»
					«event MdwsDM04»
					«event MdwsDMsr» without «class dmBV» given «class dmST»:curRecord, «class dmEN»:curRecord
					«event MdwsDMss» given «class dmSV»:1
				end try
				«event MdwsDMmd»
				delay 1
				set stillBusy to («event MdwsDMbz»)
				repeat while stillBusy is true
					delay 1
					set stillBusy to («event MdwsDMbz»)
				end repeat
				
				set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "Merge complete. Starting PDF Export process." & kTab & ((current date) as string)
				my MPSLogFileWrite(gLocalLogFilePath, gErrorMessage, 1)
				
				-- Now export as EPS, PostScript, etc.
				export curDoc in gOutputPath as "PDF" PDF output style gPDFOutputStyle
				
				set gErrorMessage to gDatabasePath & kTab & curRecord - 1 & kTab & "PDF Output complete. Created file \"" & gOutputPath & "\"." & kTab & ((current date) as string)
				my MPSLogFileWrite(gLocalLogFilePath, gErrorMessage, 1)
				
				close curDoc without saving
				display dialog "Processed record " & (curRecord - 1) giving up after 1
			end repeat
			set curRecord to curRecord + 1
		end repeat
	end tell
end MPSStartMerge

-- // --------------------------------------------------------------------
on MPSFolderPathFromFullPath(thePath)
	set fullPath to (thePath as text)
	if fullPath contains ":" then
		set pathDelim to ":"
	else
		set pathDelim to "/"
	end if
	set saveDelim to AppleScript's text item delimiters
	set AppleScript's text item delimiters to {pathDelim}
	set pathAsList to text items of fullPath
	if the last character of fullPath is pathDelim then
		set idx to (the number of text items in fullPath) - 2
	else
		set idx to -2
	end if
	set folderPath to ((text items 1 through idx of pathAsList) as text) & pathDelim
	set AppleScript's text item delimiters to saveDelim
	return folderPath
end MPSFolderPathFromFullPath

-- // --------------------------------------------------------------------
on MPSFileNameFromFullPath(thePath)
	set fullPath to (thePath as text)
	if fullPath contains ":" then
		set pathDelim to ":"
	else
		set pathDelim to "/"
	end if
	set saveDelim to AppleScript's text item delimiters
	set AppleScript's text item delimiters to {pathDelim}
	set pathAsList to text items of fullPath
	if the last character of fullPath is pathDelim then
		set idx to (the number of text items in fullPath) - 2
	else
		set idx to -2
	end if
	set folderName to item idx of pathAsList
	set AppleScript's text item delimiters to saveDelim
	return folderName
end MPSFileNameFromFullPath

-- // --------------------------------------------------------------------
on MPSCreateSubfolder(baseFolderPathPlusColon, subFolderNameNoColon)
	tell application "Finder"
		set newFolder to baseFolderPathPlusColon & subFolderNameNoColon & ":"
		if exists folder newFolder then
			return 0
		else
			try
				set theLocation to baseFolderPathPlusColon as alias
				tell application "Finder"
					set archiveFolder to make new folder at theLocation with properties {name:subFolderNameNoColon}
				end tell
			on error
				return -1
			end try
		end if
		return 0
	end tell
end MPSCreateSubfolder

-- // -----------------------------------------------
on MPSLogFileCreate(theLogFile)
	-- Got to have a POSIX path
	--	set theLogFile to POSIX path of theFilePath
	
	tell application "Finder"
		set theResult to my MPSFilePathExists(theLogFile)
		if theResult is false then
			do shell script "touch \"" & theLogFile & "\""
			try
				set theFileRef to open for access file theLogFile
			on error
				return -1
			end try
			close access theFileRef
		end if
	end tell
	return 0
end MPSLogFileCreate

-- // -----------------------------------------------
on MPSLogFileWrite(theLogFile, themessage, theReturnCount)
	local outMessage
	set lineEndChar to ((ASCII character 13) & (ASCII character 10)) -- Windows
	--	set lineEndChar to ((ASCII character 13)) -- Mac
	--	set lineEndChar to ((ASCII character 10)) -- Unix
	
	tell application "Finder"
		try
			set outMessage to themessage as string
			set theFileReference to open for access file theLogFile with write permission
			
			--	set the openFile to open for access file theLogFile with write permission
			if theReturnCount is greater than 0 then
				repeat theReturnCount times
					set outMessage to outMessage & lineEndChar
				end repeat
			end if
			write outMessage to theFileReference starting at eof
			--	write outMessage to the openFile starting at eof
			close access theFileReference
			--	close access the openFile
			return 0
		on error
			try
				close access theFileReference
			end try
			return -1
		end try
	end tell
end MPSLogFileWrite

-- // -----------------------------------------------
on MPSLogFileClear(theLogFile)
	tell application "Finder"
		try
			set the openFile to open for access file theLogFile with write permission
			set eof of the openFile to 0
			close access the openFile
			return 0
		on error
			try
				close access file openFile
			end try
			return -1
		end try
	end tell
end MPSLogFileClear

-- // -----------------------------------------------
on MPSLogFileStartNewLog(theLogFile)
	local returnVal
	tell application "Finder"
		set theResult to my MPSLogFileCreate(theLogFile)
		if (theResult is equal to 0) then
			set returnVal to 0
			set themessage to "--------------------------------------------------------------------"
			my MPSLogFileWrite(theLogFile, themessage, 1)
			set themessage to "Start Log on " & (current date) as string
			my MPSLogFileWrite(theLogFile, themessage, 1)
			set themessage to "--------------------------------------------------------------------"
			my MPSLogFileWrite(theLogFile, themessage, 1)
			set themessage to "Database" & (ASCII character 9) & "Record" & (ASCII character 9) & "Message" & (ASCII character 9) & "Date/Time"
			my MPSLogFileWrite(theLogFile, themessage, 1)
		else
			set returnVal to -1
		end if
	end tell
	return returnVal
end MPSLogFileStartNewLog

-- // -----------------------------------------------
on MPSFilePathExists(theFilePath)
	set fileFound to false
	tell application "Finder"
		try
			set fileSearchPath to theFilePath as alias
		on error
			return fileFound
		end try
		
		if exists fileSearchPath then
			set fileFound to true
		end if
	end tell
	return fileFound
end MPSFilePathExists

-- // -----------------------------------------------
on getFullPathOfActiveDocument()
	tell application "Adobe InDesign CS6"
		set theFilePath to file path of active document as string
		set theFilePath to theFilePath & (name of active document as string)
		return theFilePath
	end tell
end getFullPathOfActiveDocument



