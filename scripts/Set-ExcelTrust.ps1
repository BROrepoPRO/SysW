#!/usr/bin/env pwsh
#Requires -Version 5.1
<#
.SYNOPSIS
    Настройка доверия Excel к VBA-проекту work.xlsm.
.DESCRIPTION
    Устанавливает настройки Trust Center в реестре Windows, чтобы:
    - Папка проекта была доверенным расположением
    - Был включён доступ к VBA Project Object Model
    - Макросы были включены для доверенных расположений

    Эти настройки сохраняются после перезагрузки и очистки системы,
    в отличие от настроек через UI Excel (File -> Options -> Trust Center).

.NOTES
    Требует прав администратора для записи в реестр.
    После применения перезапустите Excel.
#>

#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Конфигурация
$ProjectPath = "L:\PROject\SysW"
$ExcelPath = "L:\PROject\SysW\work.xlsm"

# Версии Excel, которые могут быть установлены
$ExcelVersions = @(
    @{ Name = "Microsoft 365 / Office 2021"; RegPath = "HKCU:\Software\Microsoft\Office\16.0\Excel" },
    @{ Name = "Office 2019";                RegPath = "HKCU:\Software\Microsoft\Office\16.0\Excel" },
    @{ Name = "Office 2016";                RegPath = "HKCU:\Software\Microsoft\Office\16.0\Excel" },
    @{ Name = "Office 2013";                RegPath = "HKCU:\Software\Microsoft\Office\15.0\Excel" },
    @{ Name = "Office 2010";                RegPath = "HKCU:\Software\Microsoft\Office\14.0\Excel" }
)

function Write-Step {
    param([string]$Message, [string]$Status = "INFO")
    $color = @{
        "INFO"    = "Cyan"
        "OK"      = "Green"
        "WARN"    = "Yellow"
        "ERROR"   = "Red"
    }
    Write-Host ("[{0}] {1}" -f $Status.PadRight(5), $Message) -ForegroundColor $color[$Status]
}

function Test-AdminRights {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-TrustedLocation {
    <#
    .SYNOPSIS
        Добавляет папку проекта в доверенные расположения Excel.
        Использует реестр, а не UI, чтобы настройка сохранялась.
    #>
    param([string]$ExcelRegPath)

    $trustedLocationsPath = "$ExcelRegPath\Security\Trusted Locations"

    # Убедимся, что ключ существует
    if (-not (Test-Path $trustedLocationsPath)) {
        New-Item -Path $trustedLocationsPath -Force | Out-Null
        Write-Step "Создан ключ реестра: $trustedLocationsPath" -Status "INFO"
    }

    # Проверяем, есть ли уже наше расположение
    $existingLocations = Get-ItemProperty -Path $trustedLocationsPath
    $alreadyTrusted = $false

    foreach ($prop in $existingLocations.PSObject.Properties) {
        if ($prop.Name -like "Description*" -and $prop.Value -eq "SysW Project") {
            $alreadyTrusted = $true
            break
        }
    }

    if (-not $alreadyTrusted) {
        # Находим следующий доступный номер
        $locationIndex = 1
        while (Test-Path "$trustedLocationsPath\$locationIndex") {
            $locationIndex++
        }

        # Создаём новое доверенное расположение
        $locationPath = "$trustedLocationsPath\$locationIndex"
        New-Item -Path $locationPath -Force | Out-Null
        Set-ItemProperty -Path $locationPath -Name "AllowSubFolders" -Value 1 -Type DWord
        Set-ItemProperty -Path $locationPath -Name "Description" -Value "SysW Project" -Type String
        Set-ItemProperty -Path $locationPath -Name "Path" -Value $ProjectPath -Type String

        Write-Step "Папка '$ProjectPath' добавлена в доверенные расположения (индекс: $locationIndex)" -Status "OK"
    } else {
        Write-Step "Папка '$ProjectPath' уже в доверенных расположениях" -Status "OK"
    }
}

function Enable-VbaObjectModelAccess {
    <#
    .SYNOPSIS
        Включает "Trust access to the VBA project object model".
        Без этого скрипты impVBA.py и export_vba.py не работают.
    #>
    param([string]$ExcelRegPath)

    $securityPath = "$ExcelRegPath\Security"
    if (-not (Test-Path $securityPath)) {
        New-Item -Path $securityPath -Force | Out-Null
    }

    $currentValue = (Get-ItemProperty -Path $securityPath -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
    if ($currentValue -ne 1) {
        Set-ItemProperty -Path $securityPath -Name "AccessVBOM" -Value 1 -Type DWord
        Write-Step "Включён доступ к VBA Project Object Model (AccessVBOM=1)" -Status "OK"
    } else {
        Write-Step "Доступ к VBA Project Object Model уже включён" -Status "OK"
    }
}

function Enable-MacrosForTrustedLocations {
    <#
    .SYNOPSIS
        Включает все макросы для доверенных расположений.
        VBAWarnings = 1 означает "Включить все макросы для доверенных расположений".
    #>
    param([string]$ExcelRegPath)

    $securityPath = "$ExcelRegPath\Security"
    if (-not (Test-Path $securityPath)) {
        New-Item -Path $securityPath -Force | Out-Null
    }

    $currentValue = (Get-ItemProperty -Path $securityPath -Name "VBAWarnings" -ErrorAction SilentlyContinue).VBAWarnings
    if ($currentValue -ne 1) {
        Set-ItemProperty -Path $securityPath -Name "VBAWarnings" -Value 1 -Type DWord
        Write-Step "Включены макросы для доверенных расположений (VBAWarnings=1)" -Status "OK"
    } else {
        Write-Step "Макросы для доверенных расположений уже включены" -Status "OK"
    }
}

function Remove-ZoneIdentifier {
    <#
    .SYNOPSIS
        Удаляет метку "из интернета" с Excel-файла, если она есть.
    #>
    $zoneStream = Get-Item -Path $ExcelPath -Stream Zone.Identifier -ErrorAction SilentlyContinue
    if ($zoneStream) {
        Remove-Item -Path "$ExcelPath`:Zone.Identifier" -Force
        Write-Step "Удалена метка Zone.Identifier с файла $ExcelPath" -Status "OK"
    } else {
        Write-Step "Метка Zone.Identifier отсутствует" -Status "OK"
    }
}

# ===== MAIN =====
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Настройка доверия Excel к VBA-проекту" -ForegroundColor Cyan
Write-Host "  Проект: $ProjectPath" -ForegroundColor Cyan
Write-Host "  Файл:   $ExcelPath" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Проверка прав администратора
if (-not (Test-AdminRights)) {
    Write-Step "Скрипт требует прав администратора. Запустите PowerShell от имени администратора." -Status "ERROR"
    Write-Host ""
    Write-Host "Как запустить от имени администратора:" -ForegroundColor Yellow
    Write-Host "  1. Нажмите Win+X, выберите 'Windows PowerShell (Администратор)'" -ForegroundColor Yellow
    Write-Host "  2. Или: правый клик на PowerShell -> 'Запуск от имени администратора'" -ForegroundColor Yellow
    Write-Host "  3. Затем выполните: powershell -ExecutionPolicy Bypass -File scripts\Set-ExcelTrust.ps1" -ForegroundColor Yellow
    exit 1
}

# Проверка существования файла
if (-not (Test-Path $ExcelPath)) {
    Write-Step "Файл не найден: $ExcelPath" -Status "ERROR"
    exit 1
}

# Применяем настройки для каждой версии Excel
$appliedCount = 0
foreach ($ver in $ExcelVersions) {
    $regPath = $ver.RegPath
    if (Test-Path $regPath) {
        Write-Step "Найдена версия Excel: $($ver.Name)" -Status "INFO"
        Write-Step "  Применение настроек к: $regPath" -Status "INFO"

        try {
            Set-TrustedLocation -ExcelRegPath $regPath
            Enable-VbaObjectModelAccess -ExcelRegPath $regPath
            Enable-MacrosForTrustedLocations -ExcelRegPath $regPath
            $appliedCount++
        } catch {
            Write-Step "  Ошибка: $_" -Status "ERROR"
        }
        Write-Host ""
    }
}

# Удаляем Zone.Identifier с файла
Remove-ZoneIdentifier

# Итог
Write-Host "========================================" -ForegroundColor Cyan
if ($appliedCount -gt 0) {
    Write-Step "Настройки применены для $appliedCount версии(й) Excel" -Status "OK"
    Write-Step "Перезапустите Excel, чтобы изменения вступили в силу" -Status "WARN"
    Write-Step "Проверьте: откройте work.xlsm -> File -> Options -> Trust Center -> Trust Center Settings" -Status "INFO"
    Write-Step "  - Macro Settings: 'Enable all macros' (для доверенных расположений)" -Status "INFO"
    Write-Step "  - Trusted Locations: должна быть L:\PROject\SysW" -Status "INFO"
    Write-Step "  - Macro Settings: 'Trust access to the VBA project object model' - включена" -Status "INFO"
} else {
    Write-Step "Excel не найден в реестре. Возможно, он не установлен." -Status "WARN"
    Write-Step "Убедитесь, что Excel установлен, и запустите скрипт снова." -Status "WARN"
}
Write-Host "========================================" -ForegroundColor Cyan