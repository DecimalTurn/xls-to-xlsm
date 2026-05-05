<#
.SYNOPSIS
    Converts .xls files in the input directory to .xlsm format.

.DESCRIPTION
    Opens each .xls file found in the specified input directory using the
    Excel COM object, then saves it as a macro-enabled workbook (.xlsm) in
    the specified output directory.

.PARAMETER InputDir
    Path to the directory containing .xls files to convert.
    Defaults to ".\input".

.PARAMETER OutputDir
    Path to the directory where converted .xlsm files will be saved.
    Defaults to ".\output".
#>
param(
    [string]$InputDir = ".\input",
    [string]$OutputDir = ".\output"
)

# Resolve to absolute paths
$InputDir  = Resolve-Path $InputDir
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)

# Create output directory if it doesn't exist
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$xlsFiles = Get-ChildItem -Path $InputDir -Filter "*.xls"

if ($xlsFiles.Count -eq 0) {
    Write-Host "No .xls files found in $InputDir"
    exit 0
}

# xlOpenXMLWorkbookMacroEnabled = 52
$xlFormatXlsm = 52

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

try {
    foreach ($file in $xlsFiles) {
        Write-Host "Converting: $($file.Name)"

        $Workbook = $Excel.Workbooks.Open($file.FullName)

        $outputFileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".xlsm"
        $outputPath = Join-Path $OutputDir $outputFileName

        $Workbook.SaveAs($outputPath, $xlFormatXlsm)
        $Workbook.Close($false)

        Write-Host "Saved: $outputPath"
    }
} finally {
    $Excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
}
