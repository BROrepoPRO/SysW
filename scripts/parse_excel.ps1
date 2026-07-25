[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$output = New-Object System.Collections.ArrayList

# ========================================
# PARSE UAZ.xlsm - sheet UAZw
# ========================================
[void]$output.Add('========================================')
[void]$output.Add('PARSE UAZ.xlsm - sheet UAZw')
[void]$output.Add('========================================')

$wb = $excel.Workbooks.Open('L:\PROject\SysW\base\models\UAZ.xlsm')
$ws = $wb.Sheets('UAZw')

[void]$output.Add('--- Headers (row 3, A-J) ---')
$h = @()
for ($c=1; $c -le 10; $c++) { $h += $ws.Cells.Item(3,$c).Text }
[void]$output.Add(($h -join ' | '))

$lastRow = $ws.UsedRange.Rows.Count
$dataRows = $lastRow - 3
[void]$output.Add("--- Total data rows: $dataRows ---")

[void]$output.Add('--- First 10 rows (4-13) ---')
for ($r=4; $r -le 13; $r++) {
    $row = @()
    for ($c=1; $c -le 10; $c++) { $row += $ws.Cells.Item($r,$c).Text }
    [void]$output.Add(($row -join ' | '))
}

[void]$output.Add('--- Last 5 rows ---')
$start = $lastRow - 4
for ($r=$start; $r -le $lastRow; $r++) {
    $row = @()
    for ($c=1; $c -le 10; $c++) { $row += $ws.Cells.Item($r,$c).Text }
    [void]$output.Add(($row -join ' | '))
}

# ========================================
# PARSE UAZ.xlsm - sheet UAZz4
# ========================================
[void]$output.Add('========================================')
[void]$output.Add('PARSE UAZ.xlsm - sheet UAZz4')
[void]$output.Add('========================================')

$ws = $wb.Sheets('UAZz4')

[void]$output.Add('--- Headers (row 3, A-J) ---')
$h = @()
for ($c=1; $c -le 10; $c++) { $h += $ws.Cells.Item(3,$c).Text }
[void]$output.Add(($h -join ' | '))

$lastRow = $ws.UsedRange.Rows.Count
$dataRows = $lastRow - 3
[void]$output.Add("--- Total data rows: $dataRows ---")

[void]$output.Add('--- First 10 rows (4-13) ---')
for ($r=4; $r -le 13; $r++) {
    $row = @()
    for ($c=1; $c -le 10; $c++) { $row += $ws.Cells.Item($r,$c).Text }
    [void]$output.Add(($row -join ' | '))
}

[void]$output.Add('--- Last 5 rows ---')
$start = $lastRow - 4
for ($r=$start; $r -le $lastRow; $r++) {
    $row = @()
    for ($c=1; $c -le 10; $c++) { $row += $ws.Cells.Item($r,$c).Text }
    [void]$output.Add(($row -join ' | '))
}

$wb.Close($false)

# ========================================
# PARSE work.xlsm - sheet main
# ========================================
[void]$output.Add('========================================')
[void]$output.Add('PARSE work.xlsm - sheet main')
[void]$output.Add('========================================')

$wb = $excel.Workbooks.Open('L:\PROject\SysW\work.xlsm')
$ws = $wb.Sheets('main')

[void]$output.Add('--- Header block B4:B17 ---')
for ($r=4; $r -le 17; $r++) {
    [void]$output.Add("B$r : " + $ws.Cells.Item($r,2).Text)
}

[void]$output.Add('--- Headers (row 3, A-AG = 1-33) ---')
$h = @()
for ($c=1; $c -le 33; $c++) { $h += $ws.Cells.Item(3,$c).Text }
[void]$output.Add(($h -join ' | '))

$lastRow = $ws.UsedRange.Rows.Count
$dataRows = $lastRow - 3
[void]$output.Add("--- Total data rows: $dataRows ---")

[void]$output.Add('--- First 15 rows (4-18) ---')
for ($r=4; $r -le 18; $r++) {
    $row = @()
    for ($c=1; $c -le 33; $c++) { $row += $ws.Cells.Item($r,$c).Text }
    [void]$output.Add(($row -join ' | '))
}

[void]$output.Add('--- Last 5 rows ---')
$start = $lastRow - 4
for ($r=$start; $r -le $lastRow; $r++) {
    $row = @()
    for ($c=1; $c -le 33; $c++) { $row += $ws.Cells.Item($r,$c).Text }
    [void]$output.Add(($row -join ' | '))
}

$wb.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

$output -join "`r`n" | Out-File -FilePath 'L:\PROject\SysW\plans\parsed_data.md' -Encoding UTF8
Write-Host 'Done. Saved to plans/parsed_data.md'