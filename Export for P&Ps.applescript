tell application "Keyboard Maestro Engine"
	set Temp to make variable with properties {name:"Job Number"}
	set JobNumber to value of Temp
end tell
tell application "QuarkXPress"
	print document 1 print output style "Proof"
	export layout space 1 of project 1 in "Macintosh HD:Users:Maggie:Documents:Temp PDFs:" & JobNumber & ".v1.1.pdf" as "PDF" PDF output style "PDF Proof"
	save project 1 in "HOM_Shortrun:~HOM Active Jobs:" & JobNumber & ":" & JobNumber & ".HOM.qxp"
	close every project without saving
end tell

