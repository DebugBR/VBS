Sub formatDate(inDate, inFormat)
   dim xl
   set xl = CreateObject("Excel.Application")
   WScript.StdOut.Write xl.WorksheetFunction.Text(inDate, inFormat)  
   xl.Quit
   Set xl = nothing
End Sub

If WScript.Arguments.Count < 2 Then
    WScript.echo "Wrong number of arguments. Please provide a date string and a date format string."
Else
    formatDate WScript.Arguments(0), WScript.Arguments(1)
End If