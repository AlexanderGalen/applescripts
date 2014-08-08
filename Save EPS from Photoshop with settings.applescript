tell application "Adobe Photoshop CS6"
	activate
	set PSDfilePath to file path of document 1 as string
	set CurrentDoc to document 1
	set myOptions to {class:EPS save options, preview type:eight bit TIFF, encoding:binary, halftone screen:false, transfer function:false, PostScript color management:false}
	save CurrentDoc in file PSDfilePath as Photoshop EPS with options myOptions with replacing and copying
	close CurrentDoc
	
end tell
