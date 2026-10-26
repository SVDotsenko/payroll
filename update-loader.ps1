# Replaces the loader module in an existing payroll workbook with the latest loader.bas.
# The loader cannot replace itself from VBA, so this script does it from outside Excel.
param(
	# When omitted, a file dialog asks for the workbook.
	[string]$Path
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

# When started via "Run with PowerShell" the console closes immediately, so errors are shown in a dialog.
trap {
	[void][Windows.Forms.MessageBox]::Show("The loader could not be updated:`r`n`r`n$($_.Exception.Message)",
		'Payroll loader update', 'OK', 'Error')
	exit 1
}

$LocalLoader = 'C:\repositories\payroll\loader.bas'
$RemoteLoader = 'https://raw.githubusercontent.com/SVDotsenko/payroll/main/loader.bas'
$MsoAutomationSecurityForceDisable = 3

function Release-ComObject($object) {
	if ($null -ne $object) {
		[void][Runtime.InteropServices.Marshal]::ReleaseComObject($object)
	}
}

if (-not $Path) {
	$dialog = New-Object Windows.Forms.OpenFileDialog
	$dialog.Title = 'Choose the payroll workbook'
	$dialog.Filter = 'Excel macro-enabled workbooks (*.xlsm;*.xlsb)|*.xlsm;*.xlsb|All files (*.*)|*.*'
	$dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
	if ($dialog.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) {
		exit 0
	}
	$Path = $dialog.FileName
}
$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if (-not (Test-Path $Path)) {
	throw "File not found: $Path"
}

# An Excel instance that has the workbook open locks the file; a second instance would open it read-only.
try {
	([IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')).Close()
} catch {
	throw "The workbook is open or locked. Close it in Excel and try again.`r`n`r`n$Path"
}

# Same lookup order as the loader itself: local repository first, then GitHub.
if (Test-Path $LocalLoader) {
	$source = 'the local folder'
	$loaderText = Get-Content -Path $LocalLoader -Raw
} else {
	$source = 'GitHub'
	$loaderText = [string](Invoke-WebRequest -Uri $RemoteLoader -UseBasicParsing).Content
}

# The VBA editor needs CRLF line endings to import a module correctly.
$loaderText = $loaderText -replace "`r?`n", "`r`n"
$tempLoader = Join-Path $env:TEMP 'loader.bas'
[IO.File]::WriteAllText($tempLoader, $loaderText, [Text.Encoding]::ASCII)

# Replacing a VBA module requires "Trust access to the VBA project object model".
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

$backupPath = "$Path.bak"
Copy-Item -Path $Path -Destination $backupPath -Force

$excel = New-Object -ComObject Excel.Application
$workbook = $null
$components = $null
try {
	$excel.Visible = $false
	$excel.DisplayAlerts = $false
	# The workbook's own macros must not run while its loader is being replaced.
	$excel.AutomationSecurity = $MsoAutomationSecurityForceDisable

	$workbook = $excel.Workbooks.Open($Path)
	$components = $workbook.VBProject.VBComponents

	# Removes "loader" and the numbered copies the VBA editor creates on a repeated import, e.g. "loader1".
	for ($index = $components.Count; $index -ge 1; $index--) {
		$component = $components.Item($index)
		if ($component.Name -match '^loader\d*$') {
			$components.Remove($component)
		}
		Release-ComObject $component
	}
	[void]$components.Import($tempLoader)

	$workbook.Save()
	$workbook.Close($false)
	$workbook = $null

} finally {
	if ($null -ne $workbook) {
		$workbook.Close($false)
	}
	$excel.Quit()
	Release-ComObject $components
	Release-ComObject $workbook
	Release-ComObject $excel
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
	Remove-Item -Path $tempLoader -ErrorAction SilentlyContinue
}

[void][Windows.Forms.MessageBox]::Show(
	"The loader has been updated from $source.`r`n`r`n$Path`r`n`r`nA copy of the previous version: $backupPath",
	'Payroll loader update', 'OK', 'Information')
