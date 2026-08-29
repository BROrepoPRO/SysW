# scripts/config.ps1
# Общая конфигурация путей для PowerShell-скриптов SysW.
# Все скрипты должны dot-source этот файл:
#   . "$PSScriptRoot\config.ps1"

# Версия приложения (единый источник для всей системы)
$Script:AppVersion = "1.0.15"

# Кодировка вывода — UTF-8 (Задача 3, v1.0.17).
# Устраняет «кракозябры» при передаче кириллицы из pwsh в чат SourceCraft,
# который декодирует поток как UTF-8.
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {
    Write-Warning "Не удалось установить UTF-8 для вывода консоли: $($_.Exception.Message)"
}

# Корень проекта — родительская директория scripts/
$Script:ProjectRoot = Resolve-Path "$PSScriptRoot\.."

# Основной Excel-файл
$Script:WorkbookPath = Join-Path $Script:ProjectRoot "work.xlsm"

# Директория исходников VBA
$Script:SrcDir = Join-Path $Script:ProjectRoot "src"