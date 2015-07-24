--gets product type from keyboard maestro

tell application "Keyboard Maestro Engine"
	set tempvar to make variable with properties {name:"productType"}
	set productType to value of tempvar
end tell

to savePrintFile(product)
	
	tell application "QuarkXPress"
		
		if product is "QX" then
			set theProperties to {page width:"3.75\"", page height:"9.25\""}
		else if product is "QS" then
			set theProperties to {page width:"3.75\"", page height:"7.25\""}
		else if product is "BC" then
			set theProperties to {page width:"3.75\"", page height:"2.25\""}
		end if
		
		set printfileDoc to make new document with properties theProperties
		
	end tell
	
end savePrintFile


savePrintFile(productType)