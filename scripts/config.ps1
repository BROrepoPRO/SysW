# scripts/config.ps1
# Общая конфигурация путей для PowerShell-скриптов SysW.
# Все скрипты должны dot-source этот файл:
#   . "$PSScriptRoot\config.ps1"

# Версия приложения (единый источник для всей системы)
$Script:AppVersion = "1.1.0"

# Кодировка вывода — UTF-8 (Задача 3, v1.0.17).
# Устраняет «кракозябры» при передаче кириллицы из pwsh в чат SourceCraft,
# который декодирует поток как UTF-8.
try {
    chcp 65001 > $null
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {
    Write-Warning "Не удалось установить UTF-8 для вывода консоли: $($_.Exception.Message)"
}

# Переменные окружения UTF-8 для Python-скриптов (Задача 3, v1.0.17):
# гарантируют вывод кириллицы скриптами в UTF-8, независимо от системного кодека.
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

# Корень проекта — родительская директория scripts/
$Script:ProjectRoot = Resolve-Path "$PSScriptRoot\.."

# Основной Excel-файл
$Script:WorkbookPath = Join-Path $Script:ProjectRoot "work.xlsm"

# Директория исходников VBA
$Script:SrcDir = Join-Path $Script:ProjectRoot "src"