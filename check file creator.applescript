tell application "Finder"
	--set myFile to "Macintosh HD:Users:maggie:Desktop:vector-test.eps" as alias
	set myFile to "Macintosh HD:Users:maggie:Desktop:raster-test.eps" as alias
	open for access myFile
	set FileContents to read myFile
	close access myFile
end tell


if FileContents contains "%%Creator: Adobe Illustrator" then
	return "Vector"
else if FileContents contains "%%Creator: Adobe Photoshop" then
	return "Raster"
else
	return "neither"
end if