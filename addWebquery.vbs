If WScript.Arguments.Count < 2 Then
    WScript.Echo "Wrong number of arguments."
Else
    QueryDataFromDiscoverer WScript.Arguments(0), WScript.Arguments(1)
End If

Sub QueryDataFromDiscoverer(strFullNameIQY, strFilePathToSave)       
    Dim strText
    dim wb 	
    
    'Create Excel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    set wb = objExcel.Workbooks.Add
     
    'Add query file
    objExcel.ActiveSheet.QueryTables.Add "FINDER;" & strFullNameIQY, objExcel.ActiveSheet.Range("$A$1")
    objExcel.ActiveSheet.QueryTables(1).Refresh False 'change the background to false to force waiting
    'Saving
    objExcel.DisplayAlerts = False	
    wb.SaveAs strFilePathToSave, 50,,,,,2 '50 xlsb 2 overwrite
    wb.close false
    objExcel.Quit
End Sub