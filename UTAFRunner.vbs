Set WshShell = WScript.CreateObject ("WScript.Shell")
Set colProcessList = GetObject("Winmgmts:").ExecQuery ("Select * from Win32_Process")
Set ipList = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set oFso = CreateObject("Scripting.FileSystemObject")
solutionPath = "\\ap8825\UTAFRepo\SriKrishna\CodeBase\UTAF_SAM"
libPath = solutionPath&"\lib\UTAFConfig.qfl"

Call IncludeFile(libPath)
fwVarPath = UTAF_LIB_PATH & UTAF_FWV_FL
testPath = solutionPath&"\UTAF_Test"
Call IncludeFile(fwVarPath)
interPath = UTAF_LIB_PATH & UTAF_API_FL
Call IncludeFile(interPath)
Call apiVDIHealthCheck()

Dim vFound, ipName, pcName
pcName = WshNetwork.Computername
vFound = false
For Each objProcess in colProcessList
	If objProcess.name = "UFT.exe" then
		WshShell.Run "taskkill /im UFT.exe",,True
		WScript.Sleep 5000
		WshShell.Run "taskkill /im UFT.exe",,True
		WScript.Echo "UFT process is terminated..."
		vFound = false
		Exit For
	End if
Next

For Each ipVal in ipList
	if ipVal.IPEnabled then
		ipName = ipVal.IPAddress(0)
		WScript.Echo "IP Address : "&ipName
		WScript.Echo "PC Name : "&pcName
	End If
Next

If vFound = false then
	Dim qtApp, qtTest, objFSO, DoesFolderExist
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	DoesFolderExist = objFSO.FolderExists(testPath)
	WScript.Echo "Loading scripts for App : "&UTAF_SUITE_NAME
	If dbFlag = "N" Then
		If DoesFolderExist Then
			WScript.Echo "Script triggered locally..."
			Set qtApp = CreateObject("QuickTest.Application",ipName)
			qtApp.Launch
			qtApp.Visible = False
			qtApp.Open testPath, False
			WScript.Echo "UFT launched..."
			Set qtTest = qtApp.Test
			qtTest.Run
			qtTest.Close
			qtApp.Quit
		End if
	Else
		objStartFolder = "\\ap9181\prodReports"&UTAF_SUITE_NAME&"\"&pcName
		Dim fso, objFolder, colFiles, Document, dbFileName, dbAckName
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFSO = Nothing
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set objFolder = fso.GetFolder(objStartFolder)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			dbFileName = objFile.Name
			'WScript.Echo dbFileName
			value = split(Trim(dbFileName),".")
			dbAckName = value(0)&".ack"
			'WScript.Echo dbAckName
			If fso.FileExists(objStartFolder&"\"&dbAckName) Then
				'WScript.Echo "Acknowledgment exists for "&dbFileName
			else
				If DoesFolderExist Then
					WScript.Echo "Suite picked up for execution..."
					Set qtApp = CreateObject("QuickTest.Application",ipName)
					qtApp.Launch
					qtApp.Visible = False
					qtApp.Open testPath, False
					WScript.Echo "UFT launched..."
					Set qtTest = qtApp.Test
					qtTest.Run
					qtTest.Close
					qtApp.Quit
					Exit For
				Else
					WScript.Echo "No new suite triggered. Please wait for new execution to be triggered..."
				End if
			End If
		Next
	End If
End If

Function IncludeFile(oFunctionLib)
    Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
    Dim oLibrary : Set oLibrary = oFso.OpenTextFile(oFunctionLib, 1, False, -2)
    Dim sFunctions : sFunctions = oLibrary.ReadAll
    oLibrary.Close
    Set oLibrary = Nothing
    Set oFso = Nothing
    ExecuteGlobal sFunctions
End Function
