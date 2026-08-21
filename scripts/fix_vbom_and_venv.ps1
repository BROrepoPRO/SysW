# scripts/fix_vbom_and_venv.ps1
# Диагностика и исправление двух проблем:
#   1) AccessVBOM слетает после перезагрузки => Excel виснет при COM-открытии
#      .xlsm с VBA. Выставляет Trust access to the VBA project object model.
#   2) Неактивация .venv в терминале VS Code (профиль SourceCraft).
# Запуск: pwsh -NoProfile -File scripts/fix_vbom_and_venv.ps1
# Внимание: правка реестра HKCU (только текущий пользователь).

$ErrorActionPreference = "Stop"

Write-Output "=== [1/3] Диагностика AccessVBOM ==="
$keys = @("16.0", "15.0", "14.0")
foreach ($v in $keys) {
    $base = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
    $avbom = "$base\AccessVBOM"
    if (Test-Path $avbom) {
        $val = (Get-ItemProperty -Path $avbom -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
        Write-Output ("Office {0}: AccessVBOM = {1}" -f $v, $val)
    } else {
        Write-Output ("Office {0}: AccessVBOM ключ ОТСУТСТВУЕТ (выключен)" -f $v)
    }
}

Write-Output ""
Write-Output "=== [2/3] Установка AccessVBOM = 1 (Trust VBA object model) ==="
# Находим установленную версию Office по факту наличия Security-ветки (обычно 16.0)
$targetVersion = $null
foreach ($v in $keys) {
    $base = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
    if (Test-Path $base) { $targetVersion = $v; break }
}
if (-not $targetVersion) {
    # Если ветки нет — создаём для 16.0 (Office 2016/2019/2021/365)
    $targetVersion = "16.0"
}
$avbomKey = "HKCU:\Software\Microsoft\Office\$targetVersion\Excel\Security\AccessVBOM"
if (-not (Test-Path $avbomKey)) {
    New-Item -Path $avbomKey -Force | Out-Null
    Write-Output ("Создан ключ: {0}" -f $avbomKey)
}
# AccessVBOM=1 — разрешить доступ к объектной модели VBA-проекта
Set-ItemProperty -Path $avbomKey -Name AccessVBOM -Value 1 -Type DWord
# VBAWarnings=1 — включать макросы (не спрашивать / не блокировать) — для COM-автоматизации
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$targetVersion\Excel\Security" `
    -Name VBAWarnings -Value 1 -Type DWord
# Дополнительно: VBAOff=0 — не отключать VBA вообще
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$targetVersion\Excel\Security" `
    -Name VBAOff -Value 0 -Type DWord

$check = (Get-ItemProperty -Path $avbomKey -Name AccessVBOM).AccessVBOM
Write-Output ("Office {0}: AccessVBOM установлен = {1}" -f $targetVersion, $check)

Write-Output ""
Write-Output "=== [3/3] Проверка неактивации .venv ==="
$projectRoot = Split-Path -Parent $PSScriptRoot
$venvActivate = Join-Path $projectRoot ".venv\Scripts\Activate.ps1"
if (Test-Path $venvActivate) {
    # Загружаем профиль активации venv в текущую сессию для проверки
    . $venvActivate
    $py = (Get-Command python).Source
    Write-Output ("Активация .venv в текущей сессии: {0}" -f $py)
    if ($py -like "*\SysW\.venv\Scripts\python.exe") {
        Write-Output "OK: .venv активирован (python из .venv)."
    } else {
        Write-Output "ВНИМАНИЕ: python вне .venv — профиль VS Code SourceCraft не активирует окружение."
    }
} else {
    Write-Output ("Не найден Activate.ps1: {0}" -f $venvActivate)
}

Write-Output ""
Write-Output "Готово. Для применения AccessVBOM перезапустите Excel (закройте все EXCEL.EXE)."