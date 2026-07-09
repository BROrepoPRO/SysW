<#
.SYNOPSIS
    Writes a file with UTF-8 encoding including BOM (Byte Order Mark).
.DESCRIPTION
    This script ensures the output file is saved with UTF-8 BOM encoding,
    which is required for correct display of non-ASCII characters (e.g. Cyrillic)
    in PowerShell environments.
.PARAMETER Path
    Absolute or relative path to the output file.
.PARAMETER Content
    String content to write into the file.
.EXAMPLE
    .\Write-UTF8BOM.ps1 -Path "C:\output.ps1" -Content "Write-Host 'Привет, мир!'"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Content
)

try {
    # Create UTF8 encoding object with BOM enabled ($true = emit BOM)
    $utf8WithBom = [System.Text.UTF8Encoding]::new($true)

    # Write content to file using the BOM-enabled encoding
    [System.IO.File]::WriteAllText($Path, $Content, $utf8WithBom)

    Write-Host "[OK] File saved successfully: $Path"
}
catch {
    Write-Host "[ERROR] Failed to write file: $_"
    exit 1
}