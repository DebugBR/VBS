'importFolderPath = "C:\Users\candif\OneDrive - Alcoa Corporation\Desktop\test\MRN-12-11-2018.txt" 
'semicolonTxtToExcel importFolderPath

Sub semicolonTxtToExcel(importFolderPath)
	
	Dim objExcel, objSheet, objFSO, objFile, aline, l, irow, icol, wb, fullFilePath, lNum

	Const ForReading = 1

	Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Check if file exists
    If Not objFSO.FileExists(importFolderPath) Then
        wscript.echo "File " & importFolderPath & " does not exist!"
        Exit Sub
    End If
		
	Set objFile = objFSO.OpenTextFile(importFolderPath, ForReading)
	
	Set objExcel = CreateObject("Excel.Application")

	If (Err.Number <> 0) Then
		Wscript.Echo "Excel application not found."
		Wscript.Quit
	End If
		
	 objExcel.Visible = True
	set wb = objExcel.Workbooks.Add
	Set objSheet = wb.Worksheets(1)
	objSheet.Name = "Data"

	irow= 1
	icol= 1
	While Not objFile.AtEndOfStream
		l = objFile.ReadLine
		objSheet.Cells(irow, icol) = l
		irow= irow+ 1
	Wend
	
	'Format table
	objSheet.Columns(1).TextToColumns , 1, 1, False, False, True ', False, False, False
	objSheet.ListObjects.Add(1, objSheet.cells(1,1).CurrentRegion, , 1).Name = "tbData"
	
	lNum = objSheet.ListObjects(1).ListColumns.Count
	objSheet.ListObjects(1).DataBodyRange.RemoveDuplicates BuildColArray(lNum), 2
	'objSheet.ListObjects(1).unlist

	objExcel.DisplayAlerts = False
	fullFilePath = Left(importFolderPath,len(importFolderPath)-3) & "xlsb"
	wb.SaveAs fullFilePath, 50', , , , , 2, 2
	objExcel.DisplayAlerts = True

	wb.Close (True)
	objExcel.quit
		
	Wscript.Quit

end sub

Function BuildColArray(lNum)
  Dim vMyArray
  Dim idx
  
  ReDim vMyArray(lNum - 1)
  
  For idx = 1 To lNum
    vMyArray(idx - 1) = idx
  Next 
  
  BuildColArray = vMyArray
End Function

	If WScript.Arguments.Count < 1 Then
		WScript.Echo "Wrong number of arguments."
	Else
		semicolonTxtToExcel WScript.Arguments(0)
	End if
