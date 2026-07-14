# scripts/Import-VbaFromExcel.ps1
# -*- coding: utf-8-with-bom -*-
<#
.SYNOPSIS
    Import VBA modules from work.xlsm to UTF-8 source files.
.DESCRIPTION
    Two-phase encoding model:
      1. Excel exports VBA modules in Windows-1251 encoding
      2. This script converts them to UTF-8 for storage on disk
.PARAMETER ExcelPath
    Path to the Excel workbook. Default: L:\PROject\SysW\work.xlsm
.PARAMETER OutputDir
    Output directory for .bas/.cls files. Default: L:\PROject\SysW
.PARAMETER DryRun
    Show what would be exported without actually doing it.
.EXAMPLE
    .\scripts\Import-VbaFromExcel.ps1
    .\scripts\Import-VbaFromExcel.ps1 -DryRun
    .\scripts\Import-VbaFromExcel.ps1 -ExcelPath "C:\Other\work.xlsm" -OutputDir "C:\Output"
#>

param(
    [string]$ExcelPath = "L:\PROject\SysW\work.xlsm",
    [string]$OutputDir = "L:\PROject\SysW",
    [switch]$DryRun
)

# Mapping of VBA component names to file names on disk
$componentMap = @{
    "Mod_Utils"           = "Mod_Utils.bas"
    "Mod_OrderHeader"     = "Mod_OrderHeader.bas"
    "Mod_Import"          = "Mod_Import.bas"
    "Mod_ButtonDispatcher" = "Mod_ButtonDispatcher.bas"
    "Mod_FullTestRunner"  = "Mod_FullTestRunner.bas"
    "Sheet1"              = "Sheet1_main.cls"
}

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "[OK]   $Message" -ForegroundColor Green
}

function Write-Warning {
    param([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Write-Error {
    param([string]$Message)
    Write-Host "[ERR]  $Message" -ForegroundColor Red
}

# Check if Excel file exists
if (-not (Test-Path -Path $ExcelPath)) {
    Write-Error "File not found: $ExcelPath"
    exit 1
}

# DryRun mode - only show what would be done
if ($DryRun) {
    Write-Info "DryRun mode. Would perform:"
    Write-Info "  Source: $ExcelPath"
    Write-Info "  Target directory: $OutputDir"
    Write-Info "  Components to export:"
    foreach ($entry in $componentMap.GetEnumerator()) {
        $targetFile = Join-Path -Path $OutputDir -ChildPath $entry.Value
        Write-Info "    $($entry.Key) -> $targetFile"
    }
    Write-Info "  Total components: $($componentMap.Count)"
    exit 0
}

# Create temporary directory
$tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "VbaExport_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
try {
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
    Write-Info "Temp directory: $tempDir"

    # Create COM object Excel.Application
    Write-Info "Starting Excel..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Check VBA Project access
    try {
        $vbaAccessCheck = $excel.VBE
    } catch {
        Write-Error "No access to VBA Project Object Model. Enable:"
        Write-Error "  Excel -> Options -> Customize Ribbon -> Developer"
        Write-Error "  Excel -> Options -> Trust Center -> Trust Center Settings ->"
        Write-Error "    Macro Settings -> Trust access to the VBA project object model"
        throw
    }

    # Open workbook
    Write-Info "Opening workbook: $ExcelPath"
    $workbook = $excel.Workbooks.Open($ExcelPath)
    $vbProject = $workbook.VBProject
    Write-Info "VBA Project: $($vbProject.Name)"

    # Export each component
    $exportedCount = 0
    foreach ($entry in $componentMap.GetEnumerator()) {
        $componentName = $entry.Key
        $targetFileName = $entry.Value
        $targetPath = Join-Path -Path $OutputDir -ChildPath $targetFileName

        try {
            # Find component in VBAProject
            $component = $vbProject.VBComponents.Item($componentName)
            if (-not $component) {
                Write-Warning "Component '$componentName' not found in VBAProject. Skipping."
                continue
            }

            # Export to temp directory
            $tempFilePath = Join-Path -Path $tempDir -ChildPath $targetFileName
            $component.Export($tempFilePath)

            # Read in CP1251 and write in UTF-8
            $content = [System.IO.File]::ReadAllText($tempFilePath, [System.Text.Encoding]::GetEncoding(1251))
            [System.IO.File]::WriteAllText($targetPath, $content, [System.Text.Encoding]::UTF8)

            Write-Success "$componentName -> $targetPath"
            $exportedCount++
        } catch {
            Write-Error "Error exporting '$componentName': $_"
        }
    }

    Write-Info "Exported components: $exportedCount of $($componentMap.Count)"

} catch {
    Write-Error "Critical error: $_"
    throw
} finally {
    # Cleanup
    if ($workbook) {
        $workbook.Close($false)
        Write-Info "Workbook closed."
    }
    if ($excel) {
        try {
            $excel.Quit()
            Write-Info "Excel closed."
        } catch {
            Write-Warning "Could not close Excel properly."
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # Remove temp directory
    if (Test-Path -Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force
        Write-Info "Temp directory removed."
    }
}

Write-Info "Done."
