Sub getTextRegex(inText, inPattern, inOcurrence)
	Set re = New RegExp
	re.Pattern = inPattern
	re.Global = True
	re.Multiline = true
	Set matches = re.Execute(inText)
	
	If CInt(inOcurrence) => 0 And CInt(inOcurrence) <= matches.count - 1 Then
		WScript.StdOut.Write matches(CInt(inOcurrence)).value
	Else
		WScript.StdErr.Write "Ocurrence " & inocurrence & " not found."
	End if
End Sub

If WScript.Arguments.Count < 2 Then 
	WScript.StdErr.Write "Wrong number of arguments! Please provide <text> <pattern> <ocurrence>"
End If

getTextRegex WScript.Arguments(0), WScript.Arguments(1), WScript.Arguments(2)