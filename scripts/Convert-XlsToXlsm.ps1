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

.PARAMETER ConversionTimeoutSeconds
    Maximum time in seconds allowed for a single file conversion before it
    is considered stuck (e.g. on a dialog) and forcefully terminated.
    Defaults to 120 seconds.
#>
param(
    [string]$InputDir = ".\input",
    [string]$OutputDir = ".\output",
    [int]$ConversionTimeoutSeconds = 120
)

# Resolve to absolute paths
$InputDir  = Resolve-Path $InputDir
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)

# Create output and screenshots directories
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$screenshotDir = Join-Path (Split-Path -Parent $OutputDir) "screenshots"
New-Item -ItemType Directory -Force -Path $screenshotDir | Out-Null

# Import screenshot utility
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/utils/Screenshot.ps1"

$xlsFiles = Get-ChildItem -Path $InputDir -Filter "*.xls"

if ($xlsFiles.Count -eq 0) {
    Write-Host "No .xls files found in $InputDir"
    exit 0
}

# xlOpenXMLWorkbookMacroEnabled = 52
$xlFormatXlsm = 52

foreach ($file in $xlsFiles) {
    Write-Host "Converting: $($file.Name)"
    $fileNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $outputFileName = $fileNameNoExt + ".xlsm"
    $outputPath = Join-Path $OutputDir $outputFileName

    # Take a screenshot before starting the conversion
    Take-Screenshot -OutputPath (Join-Path $screenshotDir "Screenshot_${fileNameNoExt}_start_{{timestamp}}.png")

    # Record any Excel processes already running so we only clean up the one we start
    $preExistingExcelPids = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty Id

    # Run the conversion in a background job so we can enforce a timeout.
    # Excel is made visible so that any dialog that bypasses DisplayAlerts
    # will appear on screen and be captured by a screenshot.
    # The windows-2025 runner provides a virtual desktop, so visibility is supported.
    $conversionJob = Start-Job -ScriptBlock {
        param($filePath, $outPath, $xlFormatXlsm)
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $true
        $Excel.DisplayAlerts = $false
        try {
            $Workbook = $Excel.Workbooks.Open($filePath)
            $Workbook.SaveAs($outPath, $xlFormatXlsm)
            $Workbook.Close($false)
        } finally {
            $Excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
        }
    } -ArgumentList $file.FullName, $outputPath, $xlFormatXlsm

    $completed = Wait-Job -Job $conversionJob -Timeout $ConversionTimeoutSeconds

    if ($null -eq $completed) {
        # Timed out — capture whatever is currently on screen (e.g. a blocking dialog)
        Write-Host "Warning: Conversion of '$($file.Name)' timed out after ${ConversionTimeoutSeconds}s"
        Take-Screenshot -OutputPath (Join-Path $screenshotDir "Screenshot_${fileNameNoExt}_timeout_{{timestamp}}.png")
        Stop-Job  -Job $conversionJob -ErrorAction SilentlyContinue
        Remove-Job -Job $conversionJob -ErrorAction SilentlyContinue
        # Kill only the Excel process(es) started by this job, not pre-existing ones
        Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue |
            Where-Object { $preExistingExcelPids -notcontains $_.Id } |
            Stop-Process -Force
        Write-Error "Conversion timed out for '$($file.Name)'"
        exit 1
    }

    # Take a screenshot after the job finishes (captures any residual dialog)
    Take-Screenshot -OutputPath (Join-Path $screenshotDir "Screenshot_${fileNameNoExt}_end_{{timestamp}}.png")

    $jobOutput = Receive-Job -Job $conversionJob 2>&1
    $jobState  = $conversionJob.State
    Remove-Job  -Job $conversionJob -ErrorAction SilentlyContinue

    if ($jobOutput) {
        $jobOutput | ForEach-Object { Write-Host $_ }
    }

    if ($jobState -eq 'Failed') {
        Write-Error "Conversion of '$($file.Name)' failed"
        exit 1
    }

    Write-Host "Saved: $outputPath"
}
