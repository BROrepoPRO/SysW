<#
.SYNOPSIS
    Writes a file with Windows-1251 (ANSI Cyrillic) encoding.
.DESCRIPTION
    This script ensures the output file is saved with Windows-1251 encoding,
    which is required for correct display of Cyrillic text in VBA editor (Excel).
.PARAMETER Path
    Absolute or relative path to the output file.
.PARAMETER Content
    String content to write into the file.
.EXAMPLE
    .\Write-Win1251.ps1 -Path "C:\output.bas" -Content "Attribute VB_Name = ""Test"""
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Content
)

try {
    # Create Windows-1251 encoding object
    $win1251 = [System.Text.Encoding]::GetEncoding(1251)

    # Write content to file using Windows-1251 encoding
    [System.IO.File]::WriteAllText($Path, $Content, $win1251)

    Write-Host "[OK] File saved successfully: $Path"
}
catch {
    Write-Host "[ERROR] Failed to write file: $_"
    exit 1
}