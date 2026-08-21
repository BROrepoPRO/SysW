# scripts/config.ps1
# Общая конфигурация путей для PowerShell-скриптов SysW.
# Все скрипты должны dot-source этот файл:
#   . "$PSScriptRoot\config.ps1"

# Версия приложения (единый источник для всей системы)
$Script:AppVersion = "1.0.4"

# Корень проекта — родительская директория scripts/
$Script:ProjectRoot = Resolve-Path "$PSScriptRoot\.."

# Основной Excel-файл
$Script:WorkbookPath = Join-Path $Script:ProjectRoot "work.xlsm"

# Директория исходников VBA
$Script:SrcDir = Join-Path $Script:ProjectRoot "src"