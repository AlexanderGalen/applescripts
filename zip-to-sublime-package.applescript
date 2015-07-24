tell application "Finder"
	set thisSelection to selection as alias
	set thisSelectionString to thisSelection as string
	set thisFilePath to characters 1 thru ((length of thisSelectionString) - 1) of thisSelectionString as string
	set thisDestinationPath to quoted form of POSIX path of (thisFilePath & ".sublime-package")
	set thisFilePath to quoted form of POSIX path of thisFilePath
end tell

do shell script "zip -r " & thisDestinationPath & " " & thisFilePath