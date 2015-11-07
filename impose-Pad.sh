#!/bin/bash
# make sure necessary drives are mounted
if [ ! -d /Volumes/PRINT/ ]; then
        open "smb://svc_designmerge@arc/PRINT"
fi
if [ ! -d "/Volumes/MERGE CENTRAL" ]; then
        open "smb://svc_designmerge@arc/MERGE CENTRAL"
fi

# get all pdfs in directory to impose from
find "/Volumes/PRINT/Hot Folders/Impose-Pad" -type f -name "*.pdf" -print0 | while read -d $'\0' file
do
	filename=$(basename "$file")
	noExt="${filename%.*}"
	imposedPath="/Volumes/PRINT/CACC/${noExt}.imposed.pdf"
	# call the imposition script with all it's necessary parameters
	/usr/local/bin/node "/Volumes/RESOURCE/Scripting/AAASpectacularAlex Scripts/Impose Calendar Pads.js" "$file" "$imposedPath" "CACC"
done
