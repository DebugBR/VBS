Dim oShell
Dim Service
Dim Process
dim vFound

Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.run "taskkill /IM " & WScript.arguments(0) & " /F", 0, True

Set service = GetObject ("winmgmts:")

Do
	vFound = False
	For each Process in Service.InstancesOf ("Win32_Process")
		If Process.Name = WScript.arguments(0) then
			oShell.run "taskkill /IM " & WScript.arguments(0) & " /F", 0, True
			vFound = True
		End if
	Next
Loop while vFound = True