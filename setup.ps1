# Creates a new macro-enabled workbook with the payroll loader installed.
# On every open the workbook runs UpdatePayrollModules, which pulls the latest modules.
param(
	[string]$Path = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Payroll.xlsm')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

# When started via Win+R the console closes immediately, so errors are shown in a dialog.
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

[void][Windows.Forms.MessageBox]::Show("The payroll workbook has been saved to your Desktop:`r`n`r`n$Path",
	'Payroll setup', 'OK', 'Information')
