tell application "QuarkXPress"
	set theBox to selection
	set {y1, x1, y2, x2} to bounds of theBox as list
	set x1 to (coerce x1 to real)
	set y1 to (coerce y1 to real)
	set x2 to (coerce x2 to real)
	set y2 to (coerce y2 to real)
	set theWidth to (round (x2 - x1) * 100) / 100
	set theHeight to (round (y2 - y1) * 100) / 100
end tell
{x1:x1, x2:x2, y1:y1, y2:y2, theWidth:theWidth, theHeight:theHeight}