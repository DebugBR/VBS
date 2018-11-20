Sub createFolderStructure(inPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    If fso.FolderExists(inPath) Then
    	Exit sub
    End if
    
    If InStr(1,fso.GetFileName(inPath),".", vbTextCompare) > 0 Then
       inPath = fso.GetParentFolderName(inPath)
    End If
    
    If Not fso.FolderExists(fso.GetParentFolderName(inPath)) Then
        createFolderStructure fso.GetParentFolderName(inPath)
    End If
    
    If fso.FolderExists(fso.GetParentFolderName(inPath)) Then
        fso.CreateFolder inPath
    End If
    
    On Error GoTo 0
End Sub

If WScript.Arguments.Count < 1 Then
	WScript.Echo "Wrong number of arguments."
Else
	createFolderStructure WScript.Arguments(0)
End if