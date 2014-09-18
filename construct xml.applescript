set r to 2
set FIPAutomation to "Merge Central:FIP AUTOMATION:"
set pdfBasePath to FIPAutomation & "Found Image Press Calendars:"
set excelDoc to FIPAutomation & "pdfconstructor:test.xlsx"
set theXML to FIPAutomation & "pdfconstructor:FIP_Construct.pdfc"
set finishedXML to ""


tell application "Microsoft Excel"
	open file excelDoc
	tell row 2
		set orderNumber to string value of cell 1
		set clientName to string value of cell 2
	end tell
end tell

set coverText to "FIP_" & orderNumber & " " & clientName


set finishedXML to finishedXML & "<?xml version='1.0' encoding='UTF-8'?>
<docasm linearized='true' version='1.4'>
	<resources>
		<frame id='f1' rect='0 0 432 450'/>
	</resources>
	<pages>
		<page>
			<boxes>
				<MediaBox width='432' height='450' x='0' y='0'/>
			</boxes>
			<elements>
				<text font='Helvetica' font-size='18' color='[.4 .3 .3 1]' x='Center' y='Center' style='text-align:center' frame='#f1'>" & coverText & "</text>
			</elements>
		</page>" & return
repeat while r < 7
	tell application "Microsoft Excel"
		tell row r
			set thisPDF to string value of column 4
			set qty to (value of column 3/6)
		end tell
	end tell

	set thisPDF to pdfBasePath & "FIP_" & thisPDF as string
	set thisPDF to POSIX path of thisPDF
	repeat qty times
		set finishedXML to finishedXML & "		<insert href='" & thisPDF & "' range='0-13'/>" & return
	end repeat
	set r to r + 1

end repeat

tell application "Microsoft Excel"
	close active workbook saving no
end tell

set finishedXML to finishedXML & "	</pages>
</docasm>"

set fileToEdit to open for access file theXML with write permission
	set eof fileToEdit to 0
	write finishedXML to fileToEdit starting at eof
close access file theXML
