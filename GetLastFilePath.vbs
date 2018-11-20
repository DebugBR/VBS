Sub getLastChangedFilePath(inFolder, ByRef outFilePath)
    Dim Fso
    Dim objFile, objFolder, lstCreatedFile
    
    On Error Resume Next

    Set Fso = CreateObject("Scripting.FileSystemObject") 'Creates "FileSystemObject" Object.
    
    ' Check if folder exists
    If Not Fso.FolderExists(inFolder) Then
        wscript.echo "Folder " & inFolder & " does not exist!"
        Exit Sub
    End If
    
    Set objFolder = Fso.GetFolder(inFolder)
    
    ' Check if folder has files
    If objFolder.Files.Count = 0 Then
        wscript.echo "Folder is empty!"
	Exit Sub
    End If
    
    Set lstCreatedFile = objFolder.Files(1)
    
    ' Loop files and get last modified file path
    For Each objFile In objFolder.Files
        If objFile.DateCreated > lstCreatedFile.DateCreated Then
            Set lstCreatedFile = objFile
        End If
    Next
    
    WScript.StdOut.Write lstCreatedFile.Path
   
    If err.Number <> 0 Then
       wscript.echo "Unhandled error on GetLastFilePath::getLastChangedFilePath(): " & Err.Description & ". Code:" & Err.Number _
	& " Line: " & err.Line
    End If

    On Error Goto 0
End Sub

If wscript.Arguments.Count <> 2 Then
    wscript.echo "Wrong number of parameters."
Else
    getLastChangedFilePath wscript.Arguments(0), wscript.Arguments(1)
End If