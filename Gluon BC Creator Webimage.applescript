
--subroutines for encoding text so it makes it through http okay.
on encode_char(this_char)
	set the ASCII_num to (the ASCII number this_char)
	set the hex_list to {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F"}
	set x to item ((ASCII_num div 16) + 1) of the hex_list
	set y to item ((ASCII_num mod 16) + 1) of the hex_list
	return ("%" & x & y) as string
end encode_char

on encode_text(this_text, encode_URL_A, encode_URL_B)
	set the standard_characters to "abcdefghijklmnopqrstuvwxyz0123456789-._"
	set the URL_A_chars to "$+!'/?;&@=#%><{}[]\"~`^\\|*"
	set the URL_B_chars to ":"
	set the acceptable_characters to the standard_characters
	if encode_URL_A is false then set the acceptable_characters to the acceptable_characters & the URL_A_chars
	if encode_URL_B is false then set the acceptable_characters to the acceptable_characters & the URL_B_chars
	set the encoded_text to ""
	repeat with this_char in this_text
		if this_char is in the acceptable_characters then
			set the encoded_text to (the encoded_text & this_char)
		else
			set the encoded_text to (the encoded_text & encode_char(this_char)) as string
		end if
	end repeat
	return the encoded_text
end encode_text

set staticData to "__EVENTTARGET=&__EVENTARGUMENT=&__VIEWSTATE=%2FwEPDwULLTIwNjMyODQxMDkPZBYCAgMPZBYGAkEPDxYEHgtOYXZpZ2F0ZVVybAVGaHR0cDovL2dsdW9uLmhvdXNlb2ZtYWduZXRzLmNvbS9JblByb2R1Y3Rpb24vVEVTVC9CUEVSSDEtMDAxLTExOC4xLnBkZh4EVGV4dAVGaHR0cDovL2dsdW9uLmhvdXNlb2ZtYWduZXRzLmNvbS9JblByb2R1Y3Rpb24vVEVTVC9CUEVSSDEtMDAxLTExOC4xLnBkZmRkAkMPFgIeCWlubmVyaHRtbAVzUmVxdWVzdCB0byBRdWFyayBTdGFydGVkOiAxMToyNzo0MC41NjM5Mjc8YnIgLz5SZXNwb25zZSBSZWNlaXZlZDogMTE6Mjc6NDIuNDM1OTMxPGJyIC8%2BRWxsYXBzZWQgVGltZTogMS44NzIgc2Vjb25kc2QCRQ8WAh8CZWRky0gFkWYylkUeBhJf4nnyJ5bSBmp0uW%2F%2BK7eyomONJbs%3D&__VIEWSTATEGENERATOR=468D649C&__EVENTVALIDATION=%2FwEdADAHuCvsQMrPTFQQT70yE94ATIFlxYLM4NyOJyecU%2FDt9lMm09rAyIFpl%2FLBskfxW8Ey7dqkZFyjA5g%2FWwnIMzpHSV7I7WzXIYBRn%2BbDPqPIJDA24ELhULpYb5LoE0jbP5GOnmG0zEVsxBvIlOW3XtWbKGsOsCBXPVObRTukox8CwR28MYWtckzKTVo2xhdNqvuUjgR4nOy9v6RGmJb07REc9S1VN%2FyQkn%2B20XYxxZ%2B8y64%2FFH0z4IPkryOh687ZVcwU5xjsXsellgnrDe55v9rrHz9qf%2BXXMoQNdGXunxxgBqk8hGWJvSXj5DM3HDGBnBwuwGLHLh8TMeGki%2FAkvoSxIMwn7266qHpSHOX%2FI45YrK%2Bm2ziB2vS5rO3rPSuoHgbizXWVt%2B49EPqkTxJkIaDQsisuaWKdHMMUfnEa%2BWZlo9BNTger49GWffmccw7DKYSdZDdoz9nbS9gI4Imiq231S81wMjCF56E5iqkRoY6d4Z5kxShrUnTfxJqeT8qdpI3iG7iaUT5f%2B1MPfpxDVmM4TBie%2BAghxYZZGdboIV12mwAMPlefdvxv6qgQ9cUwwWd%2BBNF7b6mZ9TYh1Lbd1J%2BHiqm57TozZUFslwvwk5PNBsadRlhCTtNIWQXFhK3Rzfg%2FVmT2kI8nH3b1MbQKjm%2B7qjNUOsfENxn9oxT%2B44G4MQb3k4%2Firf12UzXPwtWjZhSzhi%2FgTAGaZhKvSgrbvJDCmbFsghzo4bU%2FzjaVKlPCL3eShLfyo%2BrQhKLq8LV4kU8PQRpm5yXLT8s5Y5IlVZITgFKriLzcdn4jmBVyldnRg%2F2uFx0VNV%2FnX1lgQY810XCqyrUUNYmuTpRoA3cM2KtudYz0%2BIKEY7wnfnBzXynheWS8YVcyyd7l1%2FOD9WYNDa%2Bg8YsVGK0nbvphC%2BSmpYu7uwpMOlDmvmyk4%2BoSWbh458QYyT%2BLihQdtX18E2c8EQmkb2zgNyGtb8E%2FB8Ks%2BPmxdtuFzNPyNRoNxMB%2BqIwzRhKekJuqqgYBm%2FAZmfRxRA6IYYpQYyo1oAKK1G1u8kdbFB4wW%2F07s2Gv1dKXrbtMAw%3D%3D&ddlQuarkLayout=JPG&ddlOutputType=PDF&"

tell application "Microsoft Excel"
	tell active sheet
		tell used range
			set rc to count of rows
		end tell
		set r to 2
		repeat while r is less than rc + 1
			tell row r
				set thisTemplateDir to my encode_text(string value of cell 1, true, true)
				set thisTemplateName to my encode_text(string value of cell 2, true, true)
				set thisPhoto to my encode_text(string value of cell 3, true, true)
				
				
				
				--for the Gallery image, use empty info lines to let the text in the original document be used, and use the 3 default symbols
				set galleryOutputFile to thisTemplateName & ".G.G"
				
				set dataString to staticData & "ddlTemplateDir=" & thisTemplateDir & "&txtTemplate=" & thisTemplateName & "&txtOutputFilename=" & galleryOutputFile & "&txtCompression=10&txtScale=300&txtPhoto1=" & thisPhoto & "&txtLogo1=GBS-Xserve%3ALibrary%3AWebServer%3ADocuments%3Aprep-webimages%3A&txtbackground=GBS-Xserve%3ALibrary%3AWebServer%3ADocuments%3Aprep-webimages%3A&txtInfoLine1=&txtB_Font_InfoLine1=&txtInfoLine2=&txtInfoLine3=&txtInfoLine4=&txtInfoLine5=&txtInfoLine6=&txtInfoLine7=&txtInfoLine8=&txtInfoLine9=&txtInfoLine10=&txtInfoLine11=&txtInfoLine12=&txtInfoLine13=&txtInfoLine14=&txtSymbol1=Quark%3AHPS+Assets%3ARealtor+Symbols+Stroked%3AA.Realtor.R.Stroke.eps&txtSymbol2=Quark%3AHPS+Assets%3ARealtor+Symbols+Stroked%3AC.Realtor.MLS.Stroke.eps&txtSymbol3=Quark%3AHPS+Assets%3ARealtor+Symbols+Stroked%3AB.EqualHousing.Stroke.eps&btnSend=Send%20to%20Quark"
				
				set pdfDestination to quoted form of ("/Users/maggie/documents/WEB MERGE/Full PDFS/" & galleryOutputFile & ".pdf")
				set galleryPdfUrl to "http://gluon.houseofmagnets.com/InProduction/TEST/" & galleryOutputFile & ".pdf"
				--sends an http request with all that data
				do shell script "curl --data " & quoted form of dataString & " http://dev.houseofmagnets.com/utilities/bcardcreator/"
				do shell script "curl -o " & pdfDestination & " " & galleryPdfUrl
				
				
				
				--build preview PDFs
				
				set previewOutputFile to thisTemplateName & ".P.P"
				set placeholderPhoto to "GBS-Xserve%3ALibrary%3AWebServer%3ADocuments%3Aprep-webimages%3APlaceHolder_Headshot.eps"
				
				set dataString to staticData & "ddlTemplateDir=" & thisTemplateDir & "&txtTemplate=" & thisTemplateName & "&txtOutputFilename=" & previewOutputFile & "&txtCompression=10&txtScale=300&txtPhoto1=" & placeholderPhoto & "&txtLogo1=GBS-Xserve%3ALibrary%3AWebServer%3ADocuments%3Aprep-webimages%3A&txtbackground=GBS-Xserve%3ALibrary%3AWebServer%3ADocuments%3Aprep-webimages%3A&txtInfoLine1=Agent%20Name&txtB_Font_InfoLine1=&txtInfoLine2=Title&txtInfoLine3=Info%20Line%201&txtInfoLine4=Info%20Line%202&txtInfoLine5=Info%20Line%203&txtInfoLine6=Info%20Line%204&txtInfoLine7=Info%20Line%205&txtInfoLine8=Info%20Line%206&txtInfoLine9=Info%20Line%207&txtInfoLine10=Info%20Line%208&txtInfoLine11=Info%20Line%209&txtInfoLine12=Info%20Line%2010&txtInfoLine13=Info%20Line%2011&txtInfoLine14=Info%20Line%2012&txtSymbol1=Quark%3AHPS%20Assets%3ARealtor%20Symbols%20Stroked%3Asymbol1-placeholder.jpg&txtSymbol2=Quark%3AHPS%20Assets%3ARealtor%20Symbols%20Stroked%3Asymbol2-placeholder.jpg&txtSymbol3=Quark%3AHPS%20Assets%3ARealtor%20Symbols%20Stroked%3Asymbol3-placeholder.jpg&btnSend=Send%20to%20Quark"
				
				set pdfDestination to quoted form of ("/Users/maggie/documents/WEB MERGE/Full PDFS/" & previewOutputFile & ".pdf")
				set previewPdfUrl to "http://gluon.houseofmagnets.com/InProduction/TEST/" & previewOutputFile & ".pdf"
				--sends an http request with all that data
				do shell script "curl --data " & quoted form of dataString & " http://dev.houseofmagnets.com/utilities/bcardcreator/"
				do shell script "curl -o " & pdfDestination & " " & previewPdfUrl
				
				
				set r to r + 1
				
			end tell
		end repeat
	end tell
end tell
