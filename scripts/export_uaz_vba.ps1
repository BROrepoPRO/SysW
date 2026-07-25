# Export VBA from UAZ.xlsm
$UAZPath = "L:\PROject\SysW\base\models\UAZ.xlsm"
$OutputDir = "L:\PROject\SysW\scripts\_uaz_vba_export"

if (-not (Test-Path $UAZPath)) {
    Write-Host "Error: File not found: $UAZPath"
    exit 1
}

# Clean output dir
if (Test-Path $OutputDir) {
    Remove-Item -Path $OutputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

Write-Host "Creating Excel COM object..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening: $UAZPath"
    $wb = $excel.Workbooks.Open($UAZPath)

    Write-Host "Accessing VBA project..."
    $vbProject = $wb.VBProject
    $compCount = $vbProject.VBComponents.Count
    Write-Host "VBA components found: $compCount"

    foreach ($comp in $vbProject.VBComponents) {
        $compName = $comp.Name
        $compType = $comp.Type
        $typeNames = @{1 = "StdModule (.bas)"; 2 = "ClassModule (.cls)"; 3 = "MSForm (.frm)"; 100 = "Document"}
        Write-Host "`n  Component: $compName ($($typeNames[[int]$compType]))"

        $ext = @{1 = ".bas"; 2 = ".cls"; 3 = ".frm"; 100 = ".cls"}[[int]$compType]
        $outFile = Join-Path $OutputDir "$compName$ext"
        $comp.Export($outFile)
        Write-Host "    Exported -> $($outFile)"

        # Read and show first 30 lines
        $content = Get-Content -Path $outFile -Encoding Default
        Write-Host "    Lines of code: $($content.Length)"
        Write-Host "    --- First 30 lines ---"
        for ($i = 0; $i -lt [Math]::Min(30, $content.Length); $i++) {
            Write-Host "    $($content[$i])"
        }
    }

    Write-Host "`n=== All components exported to $OutputDir ==="
}
catch {
    Write-Host "Error: $_"
}
finally {
    if ($wb) { $wb.Close() }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}