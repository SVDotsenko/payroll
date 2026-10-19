# Creates a new macro-enabled workbook with the payroll loader installed.
# On every open the workbook runs UpdatePayrollModules, which pulls the latest modules.
param(
	[string]$Path = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Payroll.xlsm')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

# When started via "Run with PowerShell" the console closes immediately, so errors are shown in a dialog.
trap {
	[void][Windows.Forms.MessageBox]::Show("Payroll workbook setup failed:`r`n`r`n$($_.Exception.Message)",
		'Payroll setup', 'OK', 'Error')
	exit 1
}

$LocalLoader = 'C:\repositories\payroll\loader.bas'
$RemoteLoader = 'https://raw.githubusercontent.com/SVDotsenko/payroll/main/loader.bas'
$XlOpenXmlWorkbookMacroEnabled = 52

function Release-ComObject($object) {
	if ($null -ne $object) {
		[void][Runtime.InteropServices.Marshal]::ReleaseComObject($object)
	}
}

$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
$Path = [IO.Path]::ChangeExtension($Path, '.xlsm')
if (Test-Path $Path) {
	throw "File already exists: $Path"
}

# The loader replaces VBA modules, which requires "Trust access to the VBA project object model".
# Excel reads this setting at startup, so it is written before the working instance starts.
$excel = New-Object -ComObject Excel.Application
$excelVersion = $excel.Version
$excel.Quit()
Release-ComObject $excel
$excel = $null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

$securityKey = "HKCU:\Software\Microsoft\Office\$excelVersion\Excel\Security"
if (-not (Test-Path $securityKey)) {
	New-Item -Path $securityKey | Out-Null
}
Set-ItemProperty -Path $securityKey -Name AccessVBOM -Value 1 -Type DWord

# Same lookup order as the loader itself: local repository first, then GitHub.
if (Test-Path $LocalLoader) {
	Write-Host "Using loader from $LocalLoader"
	$loaderText = Get-Content -Path $LocalLoader -Raw
} else {
	Write-Host "Downloading loader from $RemoteLoader"
	$loaderText = [string](Invoke-WebRequest -Uri $RemoteLoader -UseBasicParsing).Content
}

# The VBA editor needs CRLF line endings to import a module correctly.
$loaderText = $loaderText -replace "`r?`n", "`r`n"
$tempLoader = Join-Path $env:TEMP 'loader.bas'
[IO.File]::WriteAllText($tempLoader, $loaderText, [Text.Encoding]::ASCII)

$excel = New-Object -ComObject Excel.Application
$workbook = $null
$project = $null
$thisWorkbook = $null
try {
	$excel.Visible = $false
	$excel.DisplayAlerts = $false

	$workbook = $excel.Workbooks.Add()
	$project = $workbook.VBProject
	[void]$project.VBComponents.Import($tempLoader)

	# CodeName is used instead of "ThisWorkbook" because localized Excel may name it differently.
	$thisWorkbook = $project.VBComponents.Item($workbook.CodeName)
	$thisWorkbook.CodeModule.AddFromString(
		"Private Sub Workbook_Open()`r`n    UpdatePayrollModules`r`nEnd Sub")

	$workbook.SaveAs($Path, $XlOpenXmlWorkbookMacroEnabled)
	$workbook.Close($false)
	$workbook = $null

} finally {
	if ($null -ne $workbook) {
		$workbook.Close($false)
	}
	$excel.Quit()
	Release-ComObject $thisWorkbook
	Release-ComObject $project
	Release-ComObject $workbook
	Release-ComObject $excel
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
	Remove-Item -Path $tempLoader -ErrorAction SilentlyContinue
}

# Windows toast notification. WinRT types are only available in Windows PowerShell 5.1,
# so under PowerShell 7 (e.g. when started via "irm ... | iex") the toast is shown by a 5.1 child process.
$showToast = {
	param([string]$message)
	$ErrorActionPreference = 'Stop'
	$ProgressPreference = 'SilentlyContinue'

	# Register our own app ID so the toast header shows "Payroll setup" instead of a raw ID.
	$appId = 'PayrollSetup'
	$appKey = "HKCU:\Software\Classes\AppUserModelId\$appId"
	[void](New-Item -Path $appKey -Force)
	[void](New-ItemProperty -Path $appKey -Name DisplayName -Value 'Payroll setup' -PropertyType String -Force)

	[void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
	$toastXml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(
		[Windows.UI.Notifications.ToastTemplateType]::ToastText01)
	$textNode = $toastXml.GetElementsByTagName('text').Item(0)
	[void]$textNode.AppendChild($toastXml.CreateTextNode($message))
	[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show(
		[Windows.UI.Notifications.ToastNotification]::new($toastXml))
}

try {
	$toastMessage = "The payroll workbook has been saved to your Desktop: $Path"
	if ($PSVersionTable.PSEdition -eq 'Core') {
		$windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
		& $windowsPowerShell -NoProfile -NonInteractive -Command $showToast -args $toastMessage
		if ($LASTEXITCODE -ne 0) {
			throw "Windows PowerShell could not show the toast (exit code $LASTEXITCODE)."
		}
	} else {
		& $showToast $toastMessage
	}
}
catch {
	# Fallback when toasts are unavailable: a dialog that closes itself after 3 seconds (64 = information icon).
	[void](New-Object -ComObject WScript.Shell).Popup("The payroll workbook has been saved to your Desktop:`r`n`r`n$Path",
		3, 'Payroll setup', 64)
}
