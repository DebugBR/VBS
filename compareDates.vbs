On Error Resume next

Function compareDates(date1, format1, date2, format2)
	Set xl = WScript.CreateObject("Excel.Application")
	compareDates = DateDiff("d", xl.WorksheetFunction.Text(date1, format1), xl.WorksheetFunction.Text(date2, format2))
	xl.Quit
	Set xl = nothing
End function

If WScript.Arguments(0) = "-format" Then
	If WScript.Arguments.Count < 5 Then
		WScript.StdErr.Write "Wrong number of arguments! <date1> <format1> <date2> <format2>."
	Else
		WScript.StdOut.Write compareDates(WScript.Arguments(1), WScript.Arguments(2), WScript.Arguments(3), WScript.Arguments(4))
	End If
Else
	If WScript.Arguments.Count < 2 Then
		WScript.StdErr.Write "Wrong number of arguments! <date1> <date2>."
	Else
		WScript.StdOut.Write compareDates(WScript.Arguments(0), "mm/dd/yyyy", WScript.Arguments(1), "mm/dd/yyyy")
	End If
End if