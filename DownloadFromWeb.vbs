
Sub DownloadFromWeb(myURL, myPath)
	dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
	dim bStrm: Set bStrm = createobject("Adodb.Stream")


	xHttp.Open "GET", myURL, False
	xHttp.Send

	with bStrm
		.type = 1 '//binary
		.open
		.write xHttp.responseBody
		.savetofile myPath, 2 '//overwrite
	end with
end sub

If WScript.Arguments.Count < 2 Then
	WScript.Echo "Wrong number of arguments."
Else
	DownloadFromWeb WScript.Arguments(0), WScript.Arguments(1)
End if