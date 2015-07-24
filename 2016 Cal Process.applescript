--removes characters from beginning or end of string
on trim_line(this_text, trim_chars, trim_indicator)
	-- 0 = beginning, 1 = end, 2 = both
	set x to the length of the trim_chars
	-- TRIM BEGINNING
	if the trim_indicator is in {0, 2} then
		repeat while this_text begins with the trim_chars
			try
				set this_text to characters (x + 1) thru -1 of this_text as string
			on error
				-- the text contains nothing but the trim characters
				return ""
			end try
		end repeat
	end if
	-- TRIM ENDING
	if the trim_indicator is in {1, 2} then
		repeat while this_text ends with the trim_chars
			try
				set this_text to characters 1 thru -(x + 1) of this_text as string
			on error
				-- the text contains nothing but the trim characters
				return ""
			end try
		end repeat
	end if
	return this_text
end trim_line

tell application "Finder"
	set sel to selection as alias
	set filePath to sel as string
	set fileName to name of sel
	set noExt to my trim_line(fileName, ".psd", 1)
	set parentFolder to my trim_line(filePath, fileName, 1)
	set qxp to parentFolder & noExt & ".qxp" as string

	set selection to parentFolder
end tell

tell application "Adobe Photoshop CS6"
	do javascript file "Macintosh HD:Applications:Adobe Photoshop CS6:Presets:Scripts:save&exportEPS.jsx"
end tell

(*tell application "QuarkXPress"

	open file qxp

	tell document 1

		set allBoxes to every generic box

		repeat with thisBox in allBoxes
			-- name die line
			if name of color of frame of thisBox is "Die Cut" then
				set name of thisBox to "Die Cut"
			end if

			if class of thisBox is picture box then

				--update BG image to new image

				set filePath to file path of image 1 of thisBox
				if filePath is not no disk file then
					tell application "Finder"
						set filePath to filePath as alias
						set imgName to name of filePath
						set imgParent to my trim_line(filePath as string, imgName, 1)
						set newImgName to characters 1 thru 4 of imgName & "16" & characters 7 thru 14 of imgName as string
						set newImg to imgParent & newImgName as string as alias
					end tell

					--change path to 2016 path
					set image 1 of thisBox to newImg

				end if
			end if

		end repeat

		save
		set fileName to name
		set fileName to fileName as string
		set gpDest to ("HOM_Shortrun:SUPERmergeIN:Custom Magnet Backgrounds:CA:" & characters 1 thru 10 of fileName as string) & ".gp"

		--group entire document
		set selected of every generic box to true
		set grouped of group box 1 to true

		save in gpDest



		close

	end tell
end tell*)


