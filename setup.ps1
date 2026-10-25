# Creates a new macro-enabled workbook with the loader and all payroll modules installed.
# Nothing runs when the workbook opens; to update the modules later, run UpdatePayrollModules by hand.
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

$LocalRepository = 'C:\repositories\payroll'
$RepoApiUrl = 'https://api.github.com/repos/SVDotsenko/payroll/contents?ref=main'
$RawBaseUrl = 'https://raw.githubusercontent.com/SVDotsenko/payroll/main/'
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

# Same lookup order as the loader itself: the local repository (with tests) first, then GitHub (without tests).
$moduleTexts = [ordered]@{}
if (Test-Path $LocalRepository) {
	Write-Host "Using modules from $LocalRepository"
	$moduleFiles = @(Get-ChildItem -Path $LocalRepository -Filter '*.bas') +
		@(Get-ChildItem -Path (Join-Path $LocalRepository 'tests') -Filter '*.bas' -ErrorAction SilentlyContinue)
	foreach ($moduleFile in $moduleFiles) {
		$moduleTexts[$moduleFile.Name] = Get-Content -Path $moduleFile.FullName -Raw
	}
} else {
	Write-Host 'Downloading modules from GitHub'
	$entries = Invoke-RestMethod -Uri $RepoApiUrl -Headers @{ 'User-Agent' = 'payroll-setup' }
	foreach ($entry in $entries | Where-Object { $_.type -eq 'file' -and $_.name -like '*.bas' }) {
		$moduleTexts[$entry.name] = [string](Invoke-WebRequest -Uri ($RawBaseUrl + $entry.name) -UseBasicParsing).Content
	}
}
if (-not $moduleTexts.Contains('loader.bas')) {
	throw 'loader.bas was not found among the payroll modules.'
}

# Every module is collected before Excel starts, so a network failure does not leave a half-built workbook.
# The VBA editor needs CRLF line endings to import a module correctly.
$moduleFolder = Join-Path $env:TEMP 'payroll_setup'
Remove-Item -Path $moduleFolder -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $moduleFolder | Out-Null
foreach ($moduleName in $moduleTexts.Keys) {
	$moduleText = $moduleTexts[$moduleName] -replace "`r?`n", "`r`n"
	[IO.File]::WriteAllText((Join-Path $moduleFolder $moduleName), $moduleText, [Text.Encoding]::ASCII)
}

$excel = New-Object -ComObject Excel.Application
$workbook = $null
$project = $null
$components = $null
try {
	$excel.Visible = $false
	$excel.DisplayAlerts = $false

	$workbook = $excel.Workbooks.Add()
	$project = $workbook.VBProject
	$components = $project.VBComponents
	$registerCalls = @()
	foreach ($moduleName in $moduleTexts.Keys) {
		$component = $components.Import((Join-Path $moduleFolder $moduleName))
		# Same rule as RefreshFunctionDescriptions in the loader: every Sub named Register* is run once.
		$pattern = '(?im)^\s*(?:(?:public|private)\s+)?sub\s+(register\w*)'
		foreach ($match in [regex]::Matches($moduleTexts[$moduleName], $pattern)) {
			$registerCalls += "'$($workbook.Name)'!$($component.Name).$($match.Groups[1].Value)"
		}
		Release-ComObject $component
	}
	foreach ($registerCall in $registerCalls) {
		[void]$excel.Run($registerCall)
	}

	$workbook.SaveAs($Path, $XlOpenXmlWorkbookMacroEnabled)
	$workbook.Close($false)
	$workbook = $null

} finally {
	if ($null -ne $workbook) {
		$workbook.Close($false)
	}
	$excel.Quit()
	Release-ComObject $components
	Release-ComObject $project
	Release-ComObject $workbook
	Release-ComObject $excel
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
	Remove-Item -Path $moduleFolder -Recurse -Force -ErrorAction SilentlyContinue
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
