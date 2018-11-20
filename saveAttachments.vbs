Sub  saveAttachments(vEmailAcc,vEmailFoder,vDownloadFolder)
	
	Set outlook = createobject("outlook.application")
	Set session = outlook.getnamespace("mapi")
	
	'session.logon

    vEmailFoder = Split(vEmailFoder, "/")
    nFolders = UBound(vEmailFoder) + 1
    
    Select Case nFolders
    
        Case 1
            Set inbox = session.Folders(vEmailAcc).Folders(vEmailFoder(0))
        Case 2
            Set inbox = session.Folders(vEmailAcc).Folders(vEmailFoder(0)).Folders(vEmailFoder(1))
        Case 3
            Set inbox = session.Folders(vEmailAcc).Folders(vEmailFoder(0)).Folders(vEmailFoder(1)).Folders(vEmailFoder(2))
        Case Else
    
    End Select
	
	For Each m In inbox.items
		intCount = m.Attachments.Count
		If intCount > 0 Then
			For i = 1 To intCount
				m.Attachments.Item(i).SaveAsFile vDownloadFolder &"\" & CleanName(m.Attachments.Item(i).FileName)
			Next 
		End If
		m.Unread = False
	Next
	
	'session.logoff
	
	Set outlook = Nothing
	Set session = Nothing
		
End sub

Function CleanName(strName)
    Dim strPattern
    Dim strReplace

    strReplace = "" 'The replacement for the special characters
    strPattern = "/[^a-zA-Z ]/" 'The regex pattern to find special characters
    
    Set regEx = CreateObject("vbscript.regexp") 'Initialize the regex object
       
    ' Configure the regex object
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    
    ' Perform the regex replacement
    CleanName = regEx.Replace(strName, strReplace)

End Function

If WScript.Arguments.Count < 3 Then
	WScript.Echo "Wrong number of arguments."
Else
	saveAttachments WScript.Arguments(0), WScript.Arguments(1), WScript.Arguments(2)
End if