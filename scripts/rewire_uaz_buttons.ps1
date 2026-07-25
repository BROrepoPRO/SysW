# rewire_uaz_buttons.ps1
# Rewires UAZ.xlsm buttons from work26.xlsm!* to work.xlsm!Btn_UAZ_*
# Run: powershell -ExecutionPolicy Bypass -File scripts/rewire_uaz_buttons.ps1

$ErrorActionPreference = "Stop"

$uazPath = "L:\PROject\SysW\base\models\UAZ.xlsm"
$workName = "work.xlsm"

$buttonMap = @{
    "UAZ" = @{
        "Button 1" = "Btn_UAZ_Name_Click"
        "Button 2" = "Btn_UAZ_Clear_Click"
        "Button 3" = "Btn_UAZ_Name_Click"
        "Button 4" = "Btn_UAZ_Clear_Click"
        "Button 5" = "Btn_UAZ_Article_Click"
    }
    "UAZw" = @{
        "Button 4" = "Btn_UAZ_Name_Click"
        "Button 5" = "Btn_UAZ_Clear_Click"
        "Button 6" = "Btn_UAZ_Article_Click"
    }
    "z4" = @{
        "Button 1" = "Btn_UAZ_Name_Click"
        "Button 2" = "Btn_UAZ_Clear_Click"
        "Button 3" = "Btn_UAZ_Article_Click"
    }
    "UAZz4" = @{
        "Button 1" = "Btn_UAZ_Name_Click"
        "Button 2" = "Btn_UAZ_Clear_Click"
        "Button 3" = "Btn_UAZ_Article_Click"
    }
}

Write-Host "Opening UAZ.xlsm..." -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($uazPath)
    $totalRewired = 0

    foreach ($sheetName in $buttonMap.Keys) {
        $buttons = $buttonMap[$sheetName]
        Write-Host "Sheet: $sheetName" -ForegroundColor Yellow

        $ws = $wb.Sheets($sheetName)
        if (-not $ws) {
            Write-Host "  [SKIP] Sheet not found" -ForegroundColor Red
            continue
        }

        foreach ($shape in $ws.Shapes) {
            if ($shape.Type -eq 8) {
                $btnName = $shape.Name
                $btn = $shape.OLEFormat.Object

                if ($buttons.ContainsKey($btnName)) {
                    $oldAction = $btn.OnAction
                    $newMacro = "'$workName'!$($buttons[$btnName])"

                    Write-Host "  $btnName : $oldAction -> $newMacro" -ForegroundColor Green
                    $btn.OnAction = $newMacro
                    $totalRewired++
                }
            }
        }
    }

    $wb.Save()
    Write-Host "Done! Total buttons rewired: $totalRewired" -ForegroundColor Cyan
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
}
finally {
    if ($wb) { $wb.Close() }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}