On Error Resume next

Class Logger

	private CONFIG_PATH
    private CONFIG_TEMPLATE 
    private LOGPATH_DATA_PLACEHOLDER
    private TASKNAME_DATA_PLACEHOLDER
    private SNAPSHOT_PATH_DATA_PLACEHOLDER
    private HEADER_TEMPLATE
    private mFso
    private mLogPath
    Private mSnapshotPath
    private mTaskName

    private sub class_Initialize
        ' Called automatically when class is created
        set mFso = CreateObject("Scripting.FileSystemObject")
        CONFIG_TEMPLATE = "<config>" & vbCrLf _
        	& "<logpath><!--LOGPATH_DATA--></logpath>" & vbCrLf _
        	& "<taskname><!--TASKNAME--></taskname>" & vbCrLf _
        	& "<snapshotpath><!--SNAPSHOTPATH_DATA--></snapshotpath>" & vbCrLf _
        	& "</config>"
        LOGPATH_DATA_PLACEHOLDER = "<!--LOGPATH_DATA-->"
        SNAPSHOT_PATH_DATA_PLACEHOLDER = "<!--SNAPSHOTPATH_DATA-->"
        TASKNAME_DATA_PLACEHOLDER = "<!--TASKNAME-->"
        Set oShell = CreateObject( "WScript.Shell" )
        CONFIG_PATH = oShell.ExpandEnvironmentStrings("%TEMP%") & "/AALogSystemConfig.xml"
        Set oShell = Nothing
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
        set mFso = nothing
    end sub

    public sub init()
        set ts = mFso.CreateTextFile(CONFIG_PATH, true)
        strConfigData = Replace(CONFIG_TEMPLATE, LOGPATH_DATA_PLACEHOLDER, mLogPath)
        strConfigData = Replace(strConfigData, TASKNAME_DATA_PLACEHOLDER, mTaskName)
        strConfigData = Replace(strConfigData, SNAPSHOT_PATH_DATA_PLACEHOLDER, mSnapshotPath)
        ts.writeLine(strConfigData)
        ts.close
        set ts = nothing
    end sub

    public sub readConfig()
        if not mFso.fileexists(CONFIG_PATH) then 
            wscript.StdErr.write "Could not find a configuration! Please initialize first with -init <log_file_path>"
            wscript.quit
        end if
        set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.Async = "False"
        xmlDoc.Load(CONFIG_PATH)
        set colNodes = xmlDoc.selectNodes("//logpath")
        mLogPath = colNodes(0).Text
        set colNodes = xmlDoc.selectNodes("//snapshotpath")
        mSnapshotPath = colNodes(0).Text
        set colNodes = xmlDoc.selectNodes("//taskname")
        mTaskName = colNodes(0).Text
    end sub

    public sub logHeader()
        if not mFso.fileExists(mLogPath) then mfso.CreateTextFile(mLogPath)
        set ts = mFso.OpenTextFile(mLogPath, 8)
        Set oShell = CreateObject( "WScript.Shell" )
        HEADER_TEMPLATE = "===============================================================" & vbCrLf _
            & "   Task Name: " & mTaskName & vbCrLf _
            & "   Computer Name: " & oShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & vbCrLf _
            & "   User Name: " & oShell.ExpandEnvironmentStrings("%USERNAME%") & vbCrLf _
            & "   Start Time: " & now & vbCrLf _
            & "===============================================================" & vbcrlf

        ts.write(HEADER_TEMPLATE)
        ts.close
        set ts = nothing
    end sub

    public Sub logMessage(strMsg, strLevel)
        if not mFso.fileExists(mLogPath) then mfso.CreateTextFile(mLogPath)

        set ts = mFso.OpenTextFile(mLogPath, 8)
        Select case lcase(strLevel)
            Case "info"
                ts.WriteLine(now & " - (INFO): " & strMsg)
            Case "warning"
                ts.WriteLine(now & " - (WARNING): " & strMsg)
            Case "error"
                ts.WriteLine(now & " - (ERROR): " & strMsg)
            Case Else
            	wscript.StdErr.Write "Invalid log level: " & strlevel
        End select
        ts.close
        set ts = nothing
    End Sub
    
    Public Function getSnapshotFilePath()
    	Set xl = CreateObject("Excel.Application")
    	getSnapshotFilePath = mFso.BuildPath(mSnapshotPath, "snapshot_" & xl.WorksheetFunction.Text(Now, "mm.dd.yy-hh.mm.ss") & ".png" )
    	Set xl = nothing
    End Function
    
    Public Sub finalize()
    	mFso.DeleteFile(CONFIG_PATH)
    End sub

    public property get LogPath
        LogPath = mLogPath
    end Property
    public property let LogPath(Value)
        if left(value, 1) = "'" Or left(value, 1) = """" then value = right(value, len(value) - 1)
        if right(value, 1) = "'" Or right(value, 1) = """"  then value = left(value, len(value) - 1)
        mLogPath = value
    end Property
    public property get SnapshotPath
        set SnapshotPath = mSnapshotPath
    end property
    public property let SnapshotPath(Value)
    	if left(value, 1) = "'" Or left(value, 1) = """" then value = right(value, len(value) - 1)
        if right(value, 1) = "'" Or right(value, 1) = """" then value = left(value, len(value) - 1)
        mSnapshotPath = value
    end Property
    public property get TaskName
        set TaskName = mTaskName
    end property
    public property let TaskName(Value)
        mTaskName = value
    end Property
End Class

if wscript.arguments.count = 0 Then
    wscript.StdOut.Write "Wrong number of arguments! Please use -help to see all the features."
    wscript.quit
end if

set lg = new Logger
if lcase(wscript.Arguments(0)) = "-init" then
    if wscript.arguments.count < 4 then
        wscript.StdErr.Write "Wrong number of arguments! Please use: -init <log_file_path> <snapshot_folder_path> <task_name>"
        wscript.quit
    end If
    lg.LogPath = wscript.Arguments(1)
    lg.SnapshotPath = wscript.Arguments(2)
    lg.TaskName = wscript.Arguments(3)
    lg.init
elseif lcase(wscript.Arguments(0)) = "-header" then
    if wscript.arguments.count < 1 then
        wscript.StdErr.Write "Wrong number of arguments! Please use: -header"
        wscript.quit
    end if
    lg.readConfig
    lg.logHeader
ElseIf lcase(wscript.Arguments(0)) = "-snapshotpath" Then
	if wscript.arguments.count < 1 then
        wscript.StdErr.Write "Wrong number of arguments! Please use: -snapshotpath"
        wscript.quit
    end if
    lg.readConfig
    WScript.StdOut.Write lg.getSnapshotFilePath
ElseIf LCase(wscript.Arguments(0)) = "-help" Then
	WScript.stdout.Write "--- Log System Help ---" & vbCrLf _
		& "cscript aux_log_system.vbs -init <log_file_path> <task_name>: Initializes a new logging configuration." & vbCrLf _
		& "cscript aux_log_system.vbs -header: Create a header in the log file with useful information about the current execution." & vbCrLf _
		& "cscript aux_log_system.vbs <log_message> <log_level(info, warning, error)>: Logs a message in the log file with timestamp and a level label." & vbCrLf _
		& "cscript aux_log_system.vbs -help: Display script help."
ElseIf LCase(wscript.Arguments(0)) = "-finalize" Then
	lg.finalize
else
    if wscript.arguments.count < 2 then
        wscript.StdErr.Write "Wrong number of arguments! Please use: <log_message> <log_level(info, warning or error)>"
        wscript.quit
    end if
    lg.readConfig
    lg.logMessage wscript.Arguments(0), wscript.Arguments(1)
end if
