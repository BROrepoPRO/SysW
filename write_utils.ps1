#encoding: utf-8
# PowerShell script for writing VBA modules in Windows-1251

$encoding = [System.Text.Encoding]::GetEncoding(1251)
$basePath = 'L:\PROject\SysW'

function Write-File1251($path, $lines) {
    [System.IO.File]::WriteAllLines($path, $lines, $encoding)
    Write-Host ('OK: ' + $path)
}
