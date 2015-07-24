tell application "Spotify"
	set t to name of current track
	set a to artist of current track
end tell
tell application "Keyboard Maestro Engine"
	make variable with properties {name:"TrackName", value:t}
	make variable with properties {name:"ArtistName", value:a}
end tell

