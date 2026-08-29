# scripts/monitor_long.ps1
# Мониторинг длительных процессов (Задача 11, v1.1.0).
#
# Назначение:
#   Фоновый опрос состояния процессов (Excel/Python/иных) каждые N секунд с записью
#   статуса в закреплённый лог-файл (например logs/monitor_task6.log). Используется
#   при выполнении длительных операций (COM-сборка, тесты, форматирование шаблонов),
#   чтобы отслеживать прогресс и момент завершения, не занимая контекст ИИ.
#
# Параметры:
#   -LogFile    путь к файлу лога (по умолчанию logs/monitor_long.log)
#   -IntervalSec период опроса в секундах (по умолчанию 60)
#   -MaxPolls   максимальное число опросов (по умолчанию 60; 0 - без ограничения)
#   -Process    массив имён процессов для отслеживания (по умолчанию EXCEL, python)
#
# Примеры:
#   pwsh -NoProfile -File scripts/monitor_long.ps1                       # Excel+python, раз в 60с
#   pwsh -NoProfile -File scripts/monitor_long.ps1 -IntervalSec 180 -LogFile logs/monitor_task6.log
#   pwsh -NoProfile -File scripts/monitor_long.ps1 -Process @("EXCEL","python") -IntervalSec 120
#
# Требование: PowerShell 7 (pwsh). Файл в кодировке UTF-8 with BOM.

param(
    [string]$LogFile = "logs/monitor_long.log",
    [int]$IntervalSec = 60,
    [int]$MaxPolls = 60,
    [string[]]$Process = @("EXCEL", "python")
)

# Создание директории лога при необходимости
$logPath = Join-Path (Get-Location) $LogFile
$logDir = Split-Path -Parent $logPath
if ($logDir -and -not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Force -Path $logDir | Out-Null
}

# Таймстамп текущего времени
function Get-Ts { return (Get-Date -Format "HH:mm:ss") }

Add-Content -Path $logPath -Value ("--- Старт мониторинга [{0}] процессы: {1}; период: {2}с ---" `
    -f (Get-Ts), ($Process -join ","), $IntervalSec)

$poll = 0
while ($true) {
    $poll++
    $counts = foreach ($p in $Process) {
        $n = @(Get-Process -Name $p -ErrorAction SilentlyContinue).Count
        "$p=$n"
    }
    $total = ($counts | Where-Object { $_ -match "=(\d+)$" } |
        ForEach-Object { [int]($_ -replace ".*=(\d+)$", '$1') } |
        Measure-Object -Sum).Sum

    $line = "[{0}] опрос {1}: {2}" -f (Get-Ts), $poll, ($counts -join " ")
    Write-Host $line
    Add-Content -Path $logPath -Value $line

    # Завершение при отсутствии отслеживаемых процессов
    if ($total -le 0) {
        $msg = "ЗАВЕРШЕНО: {0} (все отслеживаемые процессы отсутствуют)" -f (Get-Ts)
        Write-Host $msg
        Add-Content -Path $logPath -Value $msg
        break
    }

    if ($MaxPolls -gt 0 -and $poll -ge $MaxPolls) {
        $msg = "ДОСТИГНУТ ЛИМИТ ОПРОСОВ ({0}) в {1}" -f $MaxPolls, (Get-Ts)
        Write-Host $msg
        Add-Content -Path $logPath -Value $msg
        break
    }

    Start-Sleep -Seconds $IntervalSec
}