
If WScript.Arguments.Count < 2 Then
	WScript.Echo "Wrong number of arguments."
Else
	SyncFolders WScript.Arguments(0), WScript.Arguments(1)
End if

''''You must only change the absolute paths of the two folders here
''''Call SyncFolders("C:\Users\Orig","C:\Users\Dest") 
'**********************Don't Change nothing below this line *****************************
Sub SyncFolders(strFolder1, strFolder2)
  Dim objFileSys
  Dim objFolder1
  Dim objFolder2
  Dim objFile1
  Dim objFile2
  Dim objSubFolder
  Dim arrFolders
  Dim i
  Set objFileSys = CreateObject("Scripting.FileSystemObject")
  arrFolders = Array(strFolder1, strFolder2)
  For i = 0 To 1 ' Make sure that missing folders are created first:
    If objFileSys.FolderExists(arrFolders(i)) = False Then
      'wscript.echo("Creating folder " & arrFolders(i))
      objFileSys.CreateFolder(arrFolders(i))
    End If
  Next
  Set objFolder1 = objFileSys.GetFolder(strFolder1)
  Set objFolder2 = objFileSys.GetFolder(strFolder2)
  For i = 0 To 0 'Change to 1 to sync back
    If i = 1 Then ' Reverse direction of file compare in second run
      Set objFolder1 = objFileSys.GetFolder(strFolder2)
      Set objFolder2 = objFileSys.GetFolder(strFolder1)
    End If
    For Each objFile1 in objFolder1.files
      If Not objFileSys.FileExists(objFolder2 & "\" & objFile1.name) Then
        'Wscript.Echo("Copying " & objFolder1 & "\" & objFile1.name & _
        '  " to " & objFolder2 & "\" & objFile1.name)
        objFileSys.CopyFile objFolder1 & "\" & objFile1.name, _
          objFolder2 & "\" & objFile1.name
      Else
        Set objFile2 = objFileSys.GetFile(objFolder2 & "\" & objFile1.name)
        If objFile1.DateLastModified > objFile2.DateLastModified Then
          'Wscript.Echo("Overwriting " & objFolder2 & "\" & objFile1.name & _
          '  " with " & objFolder1 & "\" & objFile1.name)
          objFileSys.CopyFile objFolder1 & "\" & objFile1.name, _
            objFolder2 & "\" & objFile1.name    
        End If
      End If
    Next
  Next
  For Each objSubFolder in objFolder1.subFolders
    Call SyncFolders(strFolder1 & "\" & objSubFolder.name, strFolder2 & _
      "\" & objSubFolder.name)
  Next
  Set objFileSys = Nothing
End Sub