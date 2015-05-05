
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

set staticData to "__EVENTTARGET=&__EVENTARGUMENT=&__VIEWSTATE=%2FwEPDwULLTIwNjMyODQxMDkPZBYCAgMPZBYGAkEPDxYEHgtOYXZpZ2F0ZVVybAVEaHR0cDovL2dsdW9uLmhvdXNlb2ZtYWduZXRzLmNvbS9JblByb2R1Y3Rpb24vVEVTVC9PVVRQVVRGSUxFTkFNRS5wZGYeBFRleHQFRGh0dHA6Ly9nbHVvbi5ob3VzZW9mbWFnbmV0cy5jb20vSW5Qcm9kdWN0aW9uL1RFU1QvT1VUUFVURklMRU5BTUUucGRmZGQCQw8WAh4JaW5uZXJodG1sBXNSZXF1ZXN0IHRvIFF1YXJrIFN0YXJ0ZWQ6IDAyOjQwOjUwLjU3MDI0MDxiciAvPlJlc3BvbnNlIFJlY2VpdmVkOiAwMjo0MDo1MS4wODUwNDE8YnIgLz5FbGxhcHNlZCBUaW1lOiAwLjUxNCBzZWNvbmRzZAJFDxYCHwIFvAQ8IURPQ1RZUEUgSFRNTCBQVUJMSUMgJy0vL1czQy8vRFREIEhUTUwgMy4yLy9FTic%2BPEhUTUw%2BPEhFQUQ%2BDTxNRVRBIEhUVFAtRVFVSVY9J0NvbnRlbnQtVHlwZScgQ09OVEVOVD0ndGV4dC9odG1sO0NIQVJTRVQ9aXNvLTg4NTktMScvPg08TUVUQSBOQU1FPSdHRU5FUkFUT1InIENvbnRlbnQ9J0dMVU9OJy8%2BPFRJVExFPkdMVU9OIE5vdGlmaWNhdGlvbjwvVElUTEU%2BPC9IRUFEPjxib2R5Pg08ZGl2IGFsaWduPSdjZW50ZXInPg08dGFibGUgd2lkdGg9JzEwMCUnIGJvcmRlcj0nMCcgY2VsbHNwYWNpbmc9JzInIGNlbGxwYWRkaW5nPScwJyBoZWlnaHQ9JzEwMCUnPjx0cj48dGQ%2BDTxkaXYgYWxpZ249J2NlbnRlcic%2BPHRhYmxlIHdpZHRoPScyNzAnIGJvcmRlcj0nMScgY2VsbHNwYWNpbmc9JzInIGNlbGxwYWRkaW5nPScxMicgYmdjb2xvcj0neWVsbG93Jz4NPHRyPjx0ZD48ZGl2IGFsaWduPSdjZW50ZXInPg1UaGUgcmVxdWVzdGVkIHRlbXBsYXRlIGNvdWxkIG5vdCBiZSBvcGVuZWQuDTwvZGl2PjwvdGQ%2BPC90cj48L3RhYmxlPjwvZGl2PjwvdGQ%2BPC90cj48L3RhYmxlPjwvZGl2PjwvYm9keT48L2h0bWw%2BDWRkkE2qcpfmaZXcaC7uG0Ga8HN9wXua6M4tyiRbqT3xx5Y%3D&__VIEWSTATEGENERATOR=468D649C&__EVENTVALIDATION=%2FwEdADCgc1aAJvvaqit07Jy6hAGGTIFlxYLM4NyOJyecU%2FDt9lMm09rAyIFpl%2FLBskfxW8Ey7dqkZFyjA5g%2FWwnIMzpHSV7I7WzXIYBRn%2BbDPqPIJDA24ELhULpYb5LoE0jbP5GOnmG0zEVsxBvIlOW3XtWbKGsOsCBXPVObRTukox8CwR28MYWtckzKTVo2xhdNqvuUjgR4nOy9v6RGmJb07REc9S1VN%2FyQkn%2B20XYxxZ%2B8y64%2FFH0z4IPkryOh687ZVcwU5xjsXsellgnrDe55v9rrHz9qf%2BXXMoQNdGXunxxgBqk8hGWJvSXj5DM3HDGBnBwuwGLHLh8TMeGki%2FAkvoSxIMwn7266qHpSHOX%2FI45YrK%2Bm2ziB2vS5rO3rPSuoHgbizXWVt%2B49EPqkTxJkIaDQsisuaWKdHMMUfnEa%2BWZlo9BNTger49GWffmccw7DKYSdZDdoz9nbS9gI4Imiq231S81wMjCF56E5iqkRoY6d4Z5kxShrUnTfxJqeT8qdpI3iG7iaUT5f%2B1MPfpxDVmM4TBie%2BAghxYZZGdboIV12mwAMPlefdvxv6qgQ9cUwwWd%2BBNF7b6mZ9TYh1Lbd1J%2BHiqm57TozZUFslwvwk5PNBsadRlhCTtNIWQXFhK3Rzfg%2FVmT2kI8nH3b1MbQKjm%2B7qjNUOsfENxn9oxT%2B44G4MQb3k4%2Firf12UzXPwtWjZhSzhi%2FgTAGaZhKvSgrbvJDCmbFsghzo4bU%2FzjaVKlPCL3eShLfyo%2BrQhKLq8LV4kU8PQRpm5yXLT8s5Y5IlVZITgFKriLzcdn4jmBVyldnRg%2F2uFx0VNV%2FnX1lgQY810XCqyrUUNYmuTpRoA3cM2KtudYz0%2BIKEY7wnfnBzXynheWS8YVcyyd7l1%2FOD9WYNDa%2Bg8YsVGK0nbvphC%2BSmpYu7uwpMOlDmvmyk4%2BoSWbh458QYyT%2BLihQdtX18E2c8EQmkb2zgNyGtb8E%2FB8Ks%2BPmxdtuFzNPyNRoNxMB%2BqIwzRhKekJuqqgYBm%2FAZmfRxRA77U2Yey1speTmuhR7R3uzW2uoCNWH4znLhUTNaRooVig%3D%3D&ddlQuarkLayout=JPG&ddlOutputType=PDF&"

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
				
				set line1 to "Agent%20Name"
				set line2 to "Title"
				set line3 to "Info%20Line%201"
				set line4 to "Info%20Line%202"
				set line5 to "Info%20Line%203"
				set line6 to "Info%20Line%204"
				set line7 to "Info%20Line%205"
				set line8 to "Info%20Line%206"
				set line9 to "Info%20Line%207"
				set line10 to "Info%20Line%208"
				set line11 to "Info%20Line%209"
				set line12 to "Info%20Line%2010"
				set line13 to "Info%20Line%2011"
				set line14 to "Info%20Line%2012"
				
				set dataString to staticData & "ddlTemplateDir=" & thisTemplateDir & "&txtTemplate=" & thisTemplateName & "&txtOutputFilename=" & previewOutputFile & "&txtCompression=10&txtScale=300&txtPhoto1=" & placeholderPhoto & "&txtLogo1=GBS-Xserve%3ALibrary%3AWebServer%3ADocuments%3Aprep-webimages%3A&txtbackground=GBS-Xserve%3ALibrary%3AWebServer%3ADocuments%3Aprep-webimages%3A&txtInfoLine1=" & line1 & "&txtB_Font_InfoLine1=&txtInfoLine2=" & line2 & "&txtInfoLine3=" & line3 & "&txtInfoLine4=" & line4 & "&txtInfoLine5=" & line5 & "&txtInfoLine6=" & line6 & "&txtInfoLine7=" & line7 & "&txtInfoLine8=" & line8 & "&txtInfoLine9=" & line9 & "&txtInfoLine10=" & line10 & "&txtInfoLine11=" & line11 & "&txtInfoLine12=" & line12 & "&txtInfoLine13=" & line13 & "&txtInfoLine14=" & line14 & "&txtSymbol1=Quark%3AHPS%20Assets%3ARealtor%20Symbols%20Stroked%3Asymbol1-placeholder.jpg&txtSymbol2=Quark%3AHPS%20Assets%3ARealtor%20Symbols%20Stroked%3Asymbol2-placeholder.jpg&txtSymbol3=Quark%3AHPS%20Assets%3ARealtor%20Symbols%20Stroked%3Asymbol3-placeholder.jpg&btnSend=Send%20to%20Quark"
				
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
