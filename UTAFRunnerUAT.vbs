Set WshShell = WScript.CreateObject ("WScript.Shell")
Set colProcessList = GetObject("Winmgmts:").ExecQuery ("Select * from Win32_Process")
Set ipList = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
Set WshNetwork = WScript.CreateObject("WScript.Network")

Dim vFound, ipName, pcName
pcName = WshNetwork.Computername

vFound = false
For Each objProcess in colProcessList
	If objProcess.name = "UFT.exe" then
		vFound = true
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
	testPath = "C:\Data\Automation\UTAF_UFT\UTAF_Test"
	objStartFolder = "\\ap9181\uatReports\BGC-BE-WX300709"
	Dim objFSO, fso, objFolder, colFiles, Document, dbFileName, dbAckName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	DoesFolderExist = objFSO.FolderExists(testPath)
	Set objFSO = Nothing
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFolder = fso.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		dbFileName = objFile.Name
		WScript.Echo dbFileName
		value = split(Trim(dbFileName),".")
		dbAckName = value(0)&".ack"
		WScript.Echo dbAckName
		If fso.FileExists(objStartFolder&"\"&dbAckName) Then
			WScript.Echo "Acknowledgment exists for "&dbFileName
		else
			If DoesFolderExist Then
				WScript.Echo "File picked for Execution : "&dbFileName
				Dim qtApp
				Dim qtTest
				Set qtApp = CreateObject("QuickTest.Application","10.119.188.76")
				qtApp.Launch
				qtApp.Visible = False
				qtApp.Open testPath, False
				Set qtTest = qtApp.Test
				qtTest.Run
				qtTest.Close
				qtApp.Quit
				Exit For
			Else
			End if
	End If
	Next
End If  
