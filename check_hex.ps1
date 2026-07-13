$path = [System.IO.Path]::GetFullPath("Mod_ButtonDispatcher.bas")
$bytes = [System.IO.File]::ReadAllBytes($path)
$text = [System.Text.Encoding]::UTF8.GetString($bytes)
$lines = $text -split "`n"
$line5 = $lines[4].TrimEnd("`r")
Write-Host ("Line 5 length: " + $line5.Length)
Write-Host ("Line 5 hex bytes around 'Mod_ButtonDispatcher':")
# Find 'Mod_ButtonDispatcher' in the line
$idx = $line5.IndexOf("Mod_ButtonDispatcher")
if ($idx -ge 0) {
    # Show bytes before it (the Russian word)
    $start = [Math]::Max(0, $idx - 20)
    $sub = $line5.Substring($start, $idx - $start)
    Write-Host ("Prefix chars:")
    foreach ($ch in $sub.ToCharArray()) {
        $val = [int]$ch
        Write-Host ("  U+" + $val.ToString("X4") + " = " + $ch)
    }
} else {
    Write-Host "Mod_ButtonDispatcher not found in line 5!"
    Write-Host ("Line 5 raw: [" + $line5 + "]")
    # Show all non-ASCII chars
    Write-Host "Non-ASCII chars in line 5:"
    foreach ($ch in $line5.ToCharArray()) {
        $val = [int]$ch
        if ($val -gt 127) {
            Write-Host ("  U+" + $val.ToString("X4"))
        }
    }
}