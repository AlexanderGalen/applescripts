#!/bin/bash
# get all pdfs in directory to impose from
find "/Volumes/PRINT/Hot Folders/Impose-House" -type f -name "*.pdf" -print0 | while read -d $'\0' file
do
	filename=$(basename "$file")
	noExt="${filename%.*}"
	imposedPath="/Volumes/PRINT/CACH/${noExt}.imposed.pdf"
	# call the imposition script with all it's necessary parameters
	/usr/local/bin/node "/Volumes/RESOURCE/Scripting/AAASpectacularAlex Scripts/Impose Calendar Pads.js" "$file" "$imposedPath" "CACH"
done
