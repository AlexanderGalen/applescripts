tell application "Adobe Photoshop CS6"
	set scriptToRun to (path to applications folder as string) & "Adobe Photoshop CS6:Presets:Scripts:Hot Folder Image Resize.jsx"
	do javascript file scriptToRun
end tell