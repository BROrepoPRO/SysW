[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open('L:\PROject\SysW\base\models\UAZ.xlsm')

$output = New-Object System.Collections.ArrayList

# ========================================
# SHEET: UAZw
# ========================================
$null = $output.Add('========================================')
$null = $output.Add('SHEET: UAZw')
$null = $output.Add('========================================')

$ws = $wb.Sheets('UAZw')
$lastCol = $ws.Cells.Item(3, $ws.Columns.Count).End(-4159).Column
$null = $output.Add("Last column (row 3): $lastCol")
$null = $output.Add('')

# Headers
$null = $output.Add('--- ALL COLUMN HEADERS (row 3) ---')
for ($c = 1; $c -le $lastCol; $c++) {
    $letter = [char](64 + $c)
    $val = $ws.Cells.Item(3, $c).Text
    $null = $output.Add("$letter ($c): $val")
}
$null = $output.Add('')

# First 5 data rows
$null = $output.Add('--- FIRST 5 DATA ROWS (4-8) ---')
for ($r = 4; $r -le 8; $r++) {
    $parts = @()
    for ($c = 1; $c -le $lastCol; $c++) {
        $letter = [char](64 + $c)
        $val = $ws.Cells.Item($r, $c).Text
        $parts += "$letter=$val"
    }
    $null = $output.Add(("Row ${r}: ") + ($parts -join ' | '))
}
$null = $output.Add('')

# 3 sample rows with filled I and J
$null = $output.Add('--- 3 SAMPLE ROWS WITH I AND J FILLED ---')
$found = 0
$maxRow = $ws.UsedRange.Rows.Count
for ($r = 4; $r -le $maxRow -and $found -lt 3; $r++) {
    $valI = $ws.Cells.Item($r, 9).Text
    $valJ = $ws.Cells.Item($r, 10).Text
    if ($valI -ne '' -and $valJ -ne '') {
        $parts = @()
        for ($c = 1; $c -le $lastCol; $c++) {
            $letter = [char](64 + $c)
            $val = $ws.Cells.Item($r, $c).Text
            $parts += "$letter=$val"
        }
        $null = $output.Add(("Row ${r}: ") + ($parts -join ' | '))
        $found++
    }
}
$null = $output.Add('')

# ========================================
# SHEET: UAZz4
# ========================================
$null = $output.Add('========================================')
$null = $output.Add('SHEET: UAZz4')
$null = $output.Add('========================================')

$ws2 = $wb.Sheets('UAZz4')
$lastCol2 = $ws2.Cells.Item(3, $ws2.Columns.Count).End(-4159).Column
$null = $output.Add("Last column (row 3): $lastCol2")
$null = $output.Add('')

# Headers
$null = $output.Add('--- ALL COLUMN HEADERS (row 3) ---')
for ($c = 1; $c -le $lastCol2; $c++) {
    $letter = [char](64 + $c)
    $val = $ws2.Cells.Item(3, $c).Text
    $null = $output.Add("$letter ($c): $val")
}
$null = $output.Add('')

# First 5 data rows
$null = $output.Add('--- FIRST 5 DATA ROWS (4-8) ---')
for ($r = 4; $r -le 8; $r++) {
    $parts = @()
    for ($c = 1; $c -le $lastCol2; $c++) {
        $letter = [char](64 + $c)
        $val = $ws2.Cells.Item($r, $c).Text
        $parts += "$letter=$val"
    }
    $null = $output.Add(("Row ${r}: ") + ($parts -join ' | '))
}
$null = $output.Add('')

# 3 sample rows with filled I and J
$null = $output.Add('--- 3 SAMPLE ROWS WITH I AND J FILLED ---')
$found = 0
$maxRow2 = $ws2.UsedRange.Rows.Count
for ($r = 4; $r -le $maxRow2 -and $found -lt 3; $r++) {
    $valI = $ws2.Cells.Item($r, 9).Text
    $valJ = $ws2.Cells.Item($r, 10).Text
    if ($valI -ne '' -and $valJ -ne '') {
        $parts = @()
        for ($c = 1; $c -le $lastCol2; $c++) {
            $letter = [char](64 + $c)
            $val = $ws2.Cells.Item($r, $c).Text
            $parts += "$letter=$val"
        }
        $null = $output.Add(("Row ${r}: ") + ($parts -join ' | '))
        $found++
    }
}
$null = $output.Add('')

# Close
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Save
$output -join "`r`n" | Out-File -FilePath 'L:\PROject\SysW\plans\parsed_data_full.md' -Encoding UTF8
Write-Host 'Done. Saved to plans\parsed_data_full.md'