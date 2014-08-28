tell application "Microsoft Excel"
	set r to 23
	repeat while r < 399
		tell row r
			set StartTime to string value of cell 3
			set endTime to string value of cell 4
			set timeDiff to value of cell 5
			if timeDiff is "" then
				
				if length of StartTime is 7 then
					set StartTime to "0" & StartTime
				end if
				if length of endTime is 7 then
					set endTime to "0" & endTime
				end if
				
				
				set shr to characters 1 thru 2 of StartTime as string as integer
				set smin to characters 4 thru 5 of StartTime as string as integer
				set ssec to characters 7 thru 8 of StartTime as string as integer
				set ehr to characters 1 thru 2 of endTime as string as integer
				set emin to characters 4 thru 5 of endTime as string as integer
				set esec to characters 7 thru 8 of endTime as string as integer
				
				set startSeconds to ((shr * 60) * 60) + (smin * 60) + ssec
				set endSeconds to ((ehr * 60) * 60) + (emin * 60) + esec
				
				set secondDiff to endSeconds - startSeconds
				set timeDiff to format (secondDiff / 60) into "000.00"
				set value of cell 5 to timeDiff
				
			end if
			
			set r to r + 1
			
		end tell
		
	end repeat
end tell