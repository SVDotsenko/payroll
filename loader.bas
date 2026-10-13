Attribute VB_Name = "loader"
Option Explicit

Public Sub UpdatePayrollModules()
	On Error GoTo UpdateFailed

	Dim moduleFile As String
	Dim moduleName As String
	Dim modulePath As String
	Dim components As Object
	Dim errorMessage As String
	Const MODULE_FOLDER As String = "C:\repositories\payroll"

	Set components = ThisWorkbook.VBProject.VBComponents

	moduleFile = Dir$(MODULE_FOLDER & "\*.bas")
	Do While Len(moduleFile) > 0
		moduleName = Left$(moduleFile, InStrRev(moduleFile, ".") - 1)
		If StrComp(moduleName, "loader", vbTextCompare) <> 0 Then
			modulePath = MODULE_FOLDER & Application.PathSeparator & moduleFile
			ReplaceStandardModule components, moduleName, modulePath
		End If
		moduleFile = Dir$
	Loop

	ThisWorkbook.Save
	Application.OnTime Now + TimeSerial(0, 0, 1), _
		"'" & ThisWorkbook.Name & "'!RefreshFunctionDescriptions"
	Exit Sub

UpdateFailed:
	errorMessage = Err.Description
	MsgBox "Payroll module update failed: " & errorMessage & vbCrLf & _
		"The workbook may have partially updated modules in memory. Check the VBA project before saving.", _
		vbExclamation, "Payroll module updater"
End Sub

Public Sub RefreshFunctionDescriptions()
	On Error GoTo RegistrationFailed

	Dim component As Object
	Dim lineNumber As Long
	Dim openingParen As Long
	Dim spacePosition As Long
	Dim codeLine As String
	Dim signature As String
	Dim procedureName As String

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

Private Sub ReplaceStandardModule(ByVal components As Object, _
	ByVal moduleName As String, ByVal modulePath As String)

	Dim componentIndex As Long
	Dim component As Object
	Dim componentName As String
	Dim suffix As String

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
