Attribute VB_Name = "loader"
Private Const LOCAL_MODULE_FOLDER = "C:\repositories\payroll"
Private Const REPO_API_URL = "https://api.github.com/repos/SVDotsenko/payroll/contents?ref=main"
Private Const RAW_BASE_URL = "https://raw.githubusercontent.com/SVDotsenko/payroll/main/"

Public Sub UpdatePayrollModules()
	On Error GoTo UpdateFailed

	Set components = ThisWorkbook.VBProject.VBComponents

	' Prefer the local repository; fall back to GitHub when it is not on this machine.
	useLocalFolder = Len(Dir$(LOCAL_MODULE_FOLDER, vbDirectory)) > 0
	If useLocalFolder Then
		Set modulePaths = GetLocalModulePaths()
	Else
		Set modulePaths = DownloadRemoteModules()
		If modulePaths Is Nothing Then
			' Offline: keep the current modules and skip the update silently.
			For Each component In components
				If component.Type = 1 And StrComp(component.Name, "loader", vbTextCompare) <> 0 Then
					ShowUpdateStatus "Payroll modules were not updated: no internet connection. Using the current version."
					Exit Sub
				End If
			Next component
			' Nothing to fall back to on the very first open, so the user has to be told.
			Err.Raise vbObjectError + 3, , "The payroll functions could not be downloaded. " & _
				"Check the internet connection and reopen the workbook."
		End If
	End If

	For Each modulePath In modulePaths
		moduleFile = Mid$(modulePath, InStrRev(modulePath, "\") + 1)
		moduleName = Left$(moduleFile, InStrRev(moduleFile, ".") - 1)
		ReplaceStandardModule components, moduleName, modulePath
		If Not useLocalFolder Then Kill modulePath
	Next modulePath

	ThisWorkbook.Save
	Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!RefreshFunctionDescriptions"
	If useLocalFolder Then source = "the local folder" Else source = "GitHub"
	ShowUpdateStatus "Payroll modules updated from " & source & " (" & modulePaths.Count & " modules)."
	Exit Sub

UpdateFailed:
	errorMessage = Err.Description
	ShowUpdateStatus "Payroll module update failed."
	MsgBox "Payroll module update failed: " & errorMessage & vbCrLf & _
		"The workbook may have partially updated modules in memory. Check the VBA project before saving.", _
		vbExclamation, "Payroll module updater"
End Sub

Public Sub RefreshFunctionDescriptions()
	On Error GoTo RegistrationFailed

	For Each component In ThisWorkbook.VBProject.VBComponents
		For lineNumber = 1 To component.CodeModule.CountOfLines
			codeLine = Trim$(component.CodeModule.Lines(lineNumber, 1))
			signature = vbNullString

			If LCase$(Left$(codeLine, 11)) = "public sub " Then
				signature = Trim$(Mid$(codeLine, 12))
			ElseIf LCase$(Left$(codeLine, 4)) = "sub " Then
				signature = Trim$(Mid$(codeLine, 5))
			End If

			If Len(signature) > 0 Then
				openingParen = InStr(signature, "(")
				If openingParen > 0 Then signature = Left$(signature, openingParen - 1)
				spacePosition = InStr(signature, " ")
				If spacePosition > 0 Then signature = Left$(signature, spacePosition - 1)
				procedureName = signature
				If LCase$(Left$(procedureName, 8)) = "register" Then
					Application.Run "'" & ThisWorkbook.Name & "'!" & _
						component.Name & "." & procedureName
				End If
			End If
		Next lineNumber
	Next component
	Exit Sub

RegistrationFailed:
	MsgBox "The modules were imported, but a function description could not be registered: " & _
		Err.Description, vbExclamation, "Payroll module updater"
End Sub

' Blinks the update result in the status bar to draw attention, then keeps it for 5 seconds.
Private Sub ShowUpdateStatus(ByVal message As String)
	For blinkIndex = 1 To 5
		Application.StatusBar = message
		WaitWithoutFreezing 0.5
		' A space hides the text; False would flash Excel's "Ready" instead.
		Application.StatusBar = " "
		WaitWithoutFreezing 0.3
	Next blinkIndex

	Application.StatusBar = message
	Application.OnTime Now + TimeSerial(0, 0, 5), "'" & ThisWorkbook.Name & "'!ClearUpdateStatus"
End Sub

Private Sub WaitWithoutFreezing(ByVal seconds As Single)
	startTime = Timer
	Do While Timer - startTime < seconds
		DoEvents
	Loop
End Sub

Private Sub ClearUpdateStatus()
	Application.StatusBar = False
End Sub

Private Sub ReplaceStandardModule(ByVal components As Object, _
	ByVal moduleName As String, ByVal modulePath As String)

	For componentIndex = components.Count To 1 Step -1
		Set component = components.Item(componentIndex)
		componentName = component.Name
		If componentName = moduleName Then
			components.Remove component
		ElseIf Left$(componentName, Len(moduleName)) = moduleName Then
			suffix = Mid$(componentName, Len(moduleName) + 1)
			If Len(suffix) > 0 And IsNumeric(suffix) Then
				components.Remove component
			End If
		End If
	Next componentIndex

	components.Import modulePath
End Sub

Private Function GetLocalModulePaths() As Collection
	Set modulePaths = New Collection

	moduleFile = Dir$(LOCAL_MODULE_FOLDER & "\*.bas")
	Do While Len(moduleFile) > 0
		If StrComp(moduleFile, "loader.bas", vbTextCompare) <> 0 Then
			modulePaths.Add LOCAL_MODULE_FOLDER & "\" & moduleFile
		End If
		moduleFile = Dir$
	Loop

	If modulePaths.Count = 0 Then Err.Raise vbObjectError + 1, , "No modules found in " & LOCAL_MODULE_FOLDER

	' Tests are loaded only from the local repository; GitHub users never get them.
	moduleFile = Dir$(LOCAL_MODULE_FOLDER & "\tests\*.bas")
	Do While Len(moduleFile) > 0
		modulePaths.Add LOCAL_MODULE_FOLDER & "\tests\" & moduleFile
		moduleFile = Dir$
	Loop

	Set GetLocalModulePaths = modulePaths
End Function

' Downloads every module before any import, so a network failure leaves the project untouched.
' Returns Nothing when the modules could not be downloaded.
Private Function DownloadRemoteModules() As Collection
	On Error GoTo DownloadFailed

	Set modulePaths = New Collection
	For Each moduleFile In GetRemoteModuleNames()
		modulePaths.Add DownloadModule(moduleFile)
	Next moduleFile

	Set DownloadRemoteModules = modulePaths
	Exit Function

DownloadFailed:
	Resume DiscardDownloads
DiscardDownloads:
	On Error Resume Next
	For Each modulePath In modulePaths
		Kill modulePath
	Next modulePath
End Function

Private Function GetRemoteModuleNames() As Collection
	Set moduleNames = New Collection
	json = HttpGet(REPO_API_URL)

	namePosition = InStr(1, json, """name"":")
	Do While namePosition > 0
		openingQuote = InStr(namePosition + 7, json, """")
		closingQuote = InStr(openingQuote + 1, json, """")
		moduleFile = Mid$(json, openingQuote + 1, closingQuote - openingQuote - 1)
		If LCase$(Right$(moduleFile, 4)) = ".bas" And _
			StrComp(moduleFile, "loader.bas", vbTextCompare) <> 0 Then
			moduleNames.Add moduleFile
		End If
		namePosition = InStr(closingQuote + 1, json, """name"":")
	Loop

	If moduleNames.Count = 0 Then Err.Raise vbObjectError + 1, , "No modules found at " & REPO_API_URL
	Set GetRemoteModuleNames = moduleNames
End Function

Private Function DownloadModule(ByVal moduleFile As String) As String
	moduleText = HttpGet(RAW_BASE_URL & moduleFile)
	' The VBA editor needs CRLF line endings to import a module correctly.
	moduleText = Replace(Replace(moduleText, vbCrLf, vbLf), vbLf, vbCrLf)

	tempFolder = Environ$("TEMP") & "\payroll_modules"
	If Len(Dir$(tempFolder, vbDirectory)) = 0 Then MkDir tempFolder
	modulePath = tempFolder & "\" & moduleFile

	fileNumber = FreeFile
	Open modulePath For Output As #fileNumber
	Print #fileNumber, moduleText;
	Close #fileNumber

	DownloadModule = modulePath
End Function

Private Function HttpGet(ByVal url As String) As String
	Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
	request.Open "GET", url, False
	request.setRequestHeader "User-Agent", "payroll-loader"
	request.setRequestHeader "Cache-Control", "no-cache"
	request.send

	If request.Status <> 200 Then
		Err.Raise vbObjectError + 2, , "GET " & url & " returned " & request.Status & " " & request.statusText
	End If
	HttpGet = request.responseText
End Function