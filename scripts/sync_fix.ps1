<#
.SYNOPSIS
    Применяет обновления конфигурации синхронизации Syncthing для проекта SysW
    на текущем устройстве.

.DESCRIPTION
    Скрипт перенастраивает папку синхронизации Syncthing так, чтобы она
    указывала ТОЛЬКО на папку проекта SysW (а не на родительскую папку),
    переименовывает папку в "SysW-syncthing", создаёт маркер .stfolder и
    гарантирует наличие файла исключений .stignore.

    Предназначен для запуска на каждом устройстве, участвующем в синхронизации.
    Требуется PowerShell 7 (pwsh).

.PARAMETER ProjectPath
    Путь к папке проекта SysW. По умолчанию определяется автоматически как
    родительская папка каталога, в котором находится этот скрипт.

.PARAMETER FolderId
    Идентификатор папки Syncthing. По умолчанию "SysW-syncthing".

.PARAMETER OldFolderId
    Старый идентификатор папки Syncthing (для переименования). По умолчанию
    "PROject-syncthing".

.EXAMPLE
    pwsh .\scripts\sync_fix.ps1

    Применяет настройки с автоматическим определением пути к проекту.

.EXAMPLE
    pwsh .\scripts\sync_fix.ps1 -ProjectPath "D:\PROject\SysW"

    Применяет настройки для проекта, расположенного по указанному пути.

.NOTES
    Требуется установленный и запущенный Syncthing (команда syncthing в PATH).
    Скрипт использует `syncthing cli` для изменения конфигурации.
#>

[CmdletBinding()]
param(
    [string]$ProjectPath,
    [string]$FolderId = "SysW-syncthing",
    [string]$OldFolderId = "PROject-syncthing"
)

$ErrorActionPreference = "Stop"

# --- Определяем путь к проекту, если не задан ---
if (-not $ProjectPath) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ProjectPath = Split-Path -Parent $scriptDir
}
$ProjectPath = [System.IO.Path]::GetFullPath($ProjectPath)

Write-Host "=== Настройка синхронизации Syncthing для проекта SysW ===" -ForegroundColor Cyan
Write-Host "Путь к проекту: $ProjectPath" -ForegroundColor Yellow

# --- Проверяем наличие syncthing ---
$syncthing = Get-Command syncthing -ErrorAction SilentlyContinue
if (-not $syncthing) {
    Write-Error "Команда 'syncthing' не найдена в PATH. Установите Syncthing и добавьте его в PATH."
    exit 1
}
Write-Host "Syncthing найден: $($syncthing.Source)" -ForegroundColor Green

# --- Проверяем, что Syncthing запущен ---
try {
    $null = & syncthing cli show system 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Syncthing не отвечает. Убедитесь, что Syncthing запущен."
        exit 1
    }
} catch {
    Write-Error "Не удалось подключиться к Syncthing: $_"
    exit 1
}
Write-Host "Syncthing запущен и отвечает." -ForegroundColor Green

# --- 1. Создаём маркер .stfolder в папке проекта ---
$stfolder = Join-Path $ProjectPath ".stfolder"
if (-not (Test-Path $stfolder)) {
    New-Item -ItemType Directory -Path $stfolder -Force | Out-Null
    Write-Host "Создан маркер .stfolder: $stfolder" -ForegroundColor Green
} else {
    Write-Host "Маркер .stfolder уже существует." -ForegroundColor DarkGray
}

# --- 2. Гарантируем наличие .stignore ---
$stignore = Join-Path $ProjectPath ".stignore"
if (-not (Test-Path $stignore)) {
    Write-Error "Файл .stignore не найден: $stignore. Скопируйте его из репозитория git."
    exit 1
}
Write-Host "Файл .stignore найден." -ForegroundColor Green

# --- 3. Вспомогательная функция: получить путь папки по id ---
function Get-FolderPath {
    param([string]$Id)
    $out = & syncthing cli config folders $Id path get 2>&1
    if ($LASTEXITCODE -eq 0) {
        return $out.Trim()
    }
    return $null
}

# --- 4. Изменяем путь папки (если папка существует) ---
Write-Host "`nПроверка текущей конфигурации папок Syncthing..." -ForegroundColor Cyan

$applied = $false
foreach ($fid in @($FolderId, $OldFolderId)) {
    $currentPath = Get-FolderPath -Id $fid
    if ($null -ne $currentPath) {
        if ($currentPath -ne $ProjectPath) {
            Write-Host "Папка '$fid': изменяю путь с '$currentPath' на '$ProjectPath'..." -ForegroundColor Yellow
            & syncthing cli config folders $fid path set $ProjectPath 2>&1 | Out-Null
            if ($LASTEXITCODE -ne 0) {
                Write-Error "Не удалось изменить путь папки '$fid'."
                exit 1
            }
            Write-Host "Путь папки '$fid' изменён." -ForegroundColor Green
        } else {
            Write-Host "Папка '$fid' уже указывает на правильный путь." -ForegroundColor DarkGray
        }
        $applied = $true
        break
    }
}

if (-not $applied) {
    Write-Warning "Папка с id '$FolderId' или '$OldFolderId' не найдена. Возможно, папка уже переименована или имеет другое имя."
}

# --- 5. Переименовываем папку (id) ---
# Если существует старая папка, переименовываем её в новую
$oldPath = Get-FolderPath -Id $OldFolderId
if ($null -ne $oldPath) {
    Write-Host "Переименовываю папку '$OldFolderId' в '$FolderId'..." -ForegroundColor Yellow
    & syncthing cli config folders $OldFolderId id set $FolderId 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Не удалось переименовать папку '$OldFolderId'."
        exit 1
    }
    Write-Host "Папка переименована в '$FolderId'." -ForegroundColor Green
} else {
    Write-Host "Папка '$OldFolderId' не найдена (уже переименована или отсутствует)." -ForegroundColor DarkGray
}

# --- 6. Проверяем статус конфигурации ---
Write-Host "`nПроверка статуса конфигурации..." -ForegroundColor Cyan
$status = & syncthing cli show config-status 2>&1
# Выводим через Write-Output, чтобы избежать интерпретации фигурных скобок как формата
Write-Output $status

Write-Host "`n=== Готово. Настройка синхронизации применена. ===" -ForegroundColor Green
Write-Host "Проверьте в Syncthing Web UI, что папка '$FolderId' указывает на '$ProjectPath'." -ForegroundColor Yellow