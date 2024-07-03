'-------------------------------------------------------------------------------------------------------------------------'
' NOTES                                                                                                                   '
'-------------------------------------------------------------------------------------------------------------------------'
' Requires .NET Framework 3.5
' "Cobra Backup Script.txt" requires a trailing blank line
' As a security precaution, using any ProcessID besides "BackupProject" will cause the script to exit before taking action.
' The TextOutput flag determines if this script will generate a log during execution.
' Cobra username/password can be passed to script as arguments or defined below.
' Having a space in the Cobra password may break the script. This is untested but I'd recommend avoiding it.
' The account used to run the Windows Task Scheduler task can impact access to network drives!
'-------------------------------------------------------------------------------------------------------------------------'
' CONFIGURATION                                                                                                           '
'-------------------------------------------------------------------------------------------------------------------------'
sBatchFolder = "C:\Users\michaelstricklin\Desktop\Backup Script"
sBatchFile = sBatchFolder & "\Backup Script Input.txt"
sCobraUser = "BACKUP_USER"
sCobraPass = "reallygoodpassword"
sInstallFolder = "C:\Program Files (x86)\Deltek\Cobra"
TextOutput = True
'-------------------------------------------------------------------------------------------------------------------------'
' SCRIPT                                                                                                                  '
'-------------------------------------------------------------------------------------------------------------------------'

'Object Creation/Constant Definition/String Initialization
sStampTime = Now()
t1 = Timer()
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oRegExProc = New RegExp
Set oRegExProj = New RegExp
Set oRegExDest = New RegExp
Set oStringBuilder = CreateObject("System.Text.StringBuilder")
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const ShowWindow = 1
Const DontShowWindow = 0
Const WaitUntilFinished = True
Const DontWaitUntilFinished = False
sHost = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
sUser = oShell.ExpandEnvironmentStrings("%USERNAME%")
sStampFormat = " yyyyMMddHHmmss"
sPath = oFSO.GetParentFolderName(oFSO.GetFile(WScript.ScriptFullName)) & "\"
sTextFile = "VBScript Output.txt"
sCredentials = "user:" & sCobraUser & "/" & sCobraPass
sExecutable = """" & sInstallFolder & "\Cobra.Api.exe" & """"

'Log environment information if enabled
If TextOutput Then
	sTextFileFullPath = sPath & sTextFile
	If oFSO.FileExists(sTextFileFullPath) Then
		Set oTXT = oFSO.OpenTextFile(sTextFileFullPath,ForAppending)
	Else
		Set oTXT = oFSO.CreateTextFile(sTextFileFullPath)
	End If
	oTXT.WriteLine vbNewLine & String(60,"-") & vbNewLine & vbNewLine & "Time: " & DateStamp(sStampTime, "yyyy-MM-dd HH:mm:ss") & vbNewLine & "Host: " & sHost & vbNewLine & "User: " & sUser & vbNewLine
End If

'Determine if Cobra credentials are being passed to script or used from manual entry
If WScript.Arguments.Count > 1 Then
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Too many arguments passed to WScript! Exiting script."
	WScript.Quit
ElseIf WScript.Arguments.Count = 1 Then
	sCredentials = WScript.Arguments(0)
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Credentials being passed from WScript.Arguments."
Else 
	'sCredentials = vbNullString
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Credentials being passed from manual string entry."
	'If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Credentials not given! Exiting script."
	'WScript.Quit
End If

'Check for Cobra.Api.exe
If oFSO.FileExists(Replace(sExecutable,"""","")) Then
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Cobra API sucessfully found."
Else
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Cobra API not found! Double check installation location in script configuration. Exiting script."
	WScript.Quit
End If

'Check for API batch file
If oFSO.FileExists(sBatchFile) Then
	Set oBatchText = oFSO.OpenTextFile(sBatchFile, ForReading)
	sBatchText = oBatchText.ReadAll
	oBatchText.Close
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Cobra API Script sucessfully found."
Else
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Cobra API Script not found! Confirm that file exists and is accessible. Exiting script."
	WScript.Quit
End If

'Define RegEx items used for batch file safety/structure checks
With oRegExProc
	.IgnoreCase = True
	.Global = True
	.Multiline = True
	.Pattern = "ProcessID=(?!BackupProject).*"
End With
With oRegExProj
	.IgnoreCase = True
	.Global = True
	.Multiline = True
	.Pattern = "Project=(.+)\r\n"
End With
With oRegExDest
	.IgnoreCase = True
	.Global = True
	.Multiline = True
	.Pattern = "Destination=(.+)\r\n"
End With

Set cProc = oRegExProc.Execute(sBatchText)
Set cProj = oRegExProj.Execute(sBatchText)
Set cDest = oRegExDest.Execute(sBatchText)

'Ensure BackupProject is the only ProcessID used
If cProc.Count > 0 Then
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Batch file contains processes other than BackupProject! Exiting script."
	WScript.Quit
End If

'Ensure a destination for every project and vice versa
If cProj.Count = cDest.Count Then
	iCount = cProj.Count
Else
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Batch file has inequal numbers of projects and destinations! Exiting script."
	WScript.Quit
End If

If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Preparing to backup " & iCount & " Cobra project(s)."
Dim aPaths
ReDim aPaths(iCount - 1, 1)

'Create array of path/filenames
For i = 0 to UBound(aPaths)
	aPaths(i, 0) = cDest(i).SubMatches(0)
	aPaths(i, 1) = cProj(i).SubMatches(0)
Next

'Prevent file collision by renaming any files that already exist by appending their DateLastModified
For i = 0 to UBound(aPaths)
	If Not oFSO.FolderExists(aPaths(i,0)) Then oFSO.CreateFolder(aPaths(i,0))
	If oFSO.FileExists(aPaths(i,0) & aPaths(i,1) & ".bkcp") Then
		Set oTMP = oFSO.GetFile(aPaths(i,0) & aPaths(i,1) & ".bkcp")
		sName = NameOnly(oTMP.Name)
		sExt = ExtOnly(oTMP.Name)
		oTMP.Name = sName & DateStamp(oTMP.DateLastModified, sStampFormat) & LCase(sExt)
		Set oTMP = Nothing
	End If
Next

'See if Cobra.Api.exe is running; if so, kill it
If IsRunning("Cobra.Api.exe") Then
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Hung process found! Killing Cobra.Api.exe..."
	KillProc("Cobra.Api.exe")
	WScript.Sleep 5000
Else
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "No hung process found."
End If

'Build command string
sCMD = sExecutable & " script:""" & sBatchFile & """ " & sCredentials

If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Running Cobra.Api.exe..."
t2 = Timer()
If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "sCMD:" & sCMD
oShell.Run sCMD, DontShowWindow, WaitUntilFinished
t3 = Timer()
If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "...complete! Process ran for " & FormatNumber(t3-t2,2) & " sec."

'Rename fresh backup files by appending the script execution start time
For i = 0 to UBound(aPaths)
	If oFSO.FileExists(aPaths(i,0) & aPaths(i,1) & ".bkcp") Then
		Set oFile = oFSO.GetFile(aPaths(i,0) & aPaths(i,1) & ".bkcp")
		sName = NameOnly(oFile.Name)
		sExt = ExtOnly(oFile.Name)
		oFile.Name = sName & DateStamp(sStampTime, sStampFormat) & LCase(sExt)
		Set oFile = Nothing
	End If
Next

'Copy logs from install location to script locatoin
If oFSO.FolderExists(sInstallFolder & "\Logs") Then
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "Copying logs."
	sCMD = "robocopy.exe """ & sInstallFolder & "\Logs"" """ & sBatchFolder & "\Logs"" /mir /z /r:5 /w:5"
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "sCMD:" & sCMD
	oShell.Run sCMD, DontShowWindow, WaitUntilFinished
Else
	If TextOutput Then oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "No logs found."
End If

'Log time metrics if enabled
t4 = Timer()
If TextOutput Then
	oTXT.WriteLine DateStamp(Now(),"HH:mm:ss ") & "SCRIPT EXECUTION COMPLETE!" & vbNewLine
	d1 = FormatNumber(t2-t1,2)
	d2 = FormatNumber(t3-t2,2)
	d3 = FormatNumber(t4-t3,2)
	d4 = FormatNumber(t4-t1,2)
	dl = UDFMax(""&Len(d1)&","&Len(d2)&","&Len(d3)&","&Len(d4)&"")
	oTXT.WriteLine "Prep         : " & PadLeft(d1, dl) & " sec."
	oTXT.WriteLine "Cobra API    : " & PadLeft(d2, dl) & " sec."
	oTXT.WriteLine "Cleanup      : " & PadLeft(d3, dl) & " sec."
	oTXT.WriteLine "TOTAL RUNTIME: " & PadLeft(d4, dl) & " sec."
End If

'Support functions that I wrote years ago. They worked as expected, so I didn't reverse engineer them any further.
Function NameOnly(ByVal sFileName)
	Dim Result, i
	Result = sFileName
	i = InStrRev(sFileName, ".")
	If ( i > 0 ) Then Result = Mid(sFileName, 1, i - 1)
	NameOnly = Result
End Function

Function ExtOnly(ByVal sFileName)
	Dim Result, i
	Result = sFileName
	i = InStrRev(sFileName, ".")
	If ( i > 0 ) Then Result = Right(sFileName, Len(sFileName) - i + 1)
	ExtOnly = Result
End Function

Function DateStamp(aData, sFormatString)
	oStringBuilder.AppendFormat_4 "{0:"&sFormatString&"}", (Array(aData))
	DateStamp = oStringBuilder.ToString()
	oStringBuilder.Length = 0
End Function

Function PadLeft(sText, iTargetLen)
	iInputLen = Len(sText)
	If iInputLen < iTargetLen Then sText = String(iTargetLen - iInputLen, " ") & sText
	PadLeft = sText
End Function

Function UDFMax(sMaxInput)
	aMax = Split(sMaxInput, ",")
	For iMax = LBound(aMax) To UBound(aMax)
		If aMax(iMax) > nMax Then nMax = aMax(iMax)
	Next
	UDFMax = nMax
End Function

Function IsRunning(sProcessName)
	sComputer = "."
	Set oWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2") 
	If oWMIService.ExecQuery("Select * from Win32_Process where name like '" & sProcessName & "'").Count > 0 Then
		IsRunning = True
	Else
		IsRunning = False
	End If
End Function

Sub KillProc(sProcessName)
	sComputer = "."
    Set oWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2") 
    Set cProcs = oWMIService.ExecQuery ("Select * from Win32_Process Where Name like '" & sProcessName & "'")
    For Each oProc in cProcs
        oProc.Terminate             
    Next
End Sub
