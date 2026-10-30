<#
.SYNOPSIS
    Builds InstallationTracker.xlsm from the VBA modules in src\ and runs one update.
    After Excel has closed, tools\scrub_workbook.ps1 removes the folder path Excel saves in the file.

.DESCRIPTION
    Needs Excel for Windows with "Trust access to the VBA project object model" switched on
    (File > Options > Trust Center > Trust Center Settings > Macro Settings).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\build_workbook.ps1
#>
param(
    [string]$Output = (Join-Path (Split-Path $PSScriptRoot -Parent) 'InstallationTracker.xlsm')
)

$ErrorActionPreference = 'Stop'
$src = Join-Path (Split-Path $PSScriptRoot -Parent) 'src'
$Output = [IO.Path]::GetFullPath($Output)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Add()
    Get-ChildItem $src -Filter *.bas | Sort-Object Name | ForEach-Object {
        [void]$wb.VBProject.VBComponents.Import($_.FullName)
    }
    if (Test-Path $Output) { Remove-Item $Output }
    $wb.SaveAs($Output, 52)    # 52 = xlOpenXMLWorkbookMacroEnabled
    $excel.Run('SetupWorkbook')
    $excel.Run('RunUpdate', $false)
    Write-Host ($wb.Worksheets.Item('Dashboard').Range('B2').Text)
    $wb.Save()
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
& (Join-Path $PSScriptRoot 'scrub_workbook.ps1') -Path $Output
Write-Host "Built $Output"
