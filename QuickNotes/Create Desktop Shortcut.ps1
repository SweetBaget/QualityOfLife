#requires -Version 5.1
<#
.SYNOPSIS
    Создаёт на рабочем столе ярлык «Быстрые заметки» со значком приложения.

.DESCRIPTION
    Запустите один раз: правый клик по файлу -> "Выполнить с помощью PowerShell".
    Готовый ярлык можно перетащить на панель задач или закрепить правым кликом —
    он уже содержит иконку AppIcon.ico.
#>

$ErrorActionPreference = 'Stop'

$dir = $PSScriptRoot
if (-not $dir) { $dir = Split-Path -Parent $MyInvocation.MyCommand.Path }

$scriptPath = Join-Path $dir 'QuickNotes.ps1'
$iconPath   = Join-Path $dir 'AppIcon.ico'
if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
    throw "Не найден QuickNotes.ps1 в папке: $dir"
}

$psExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path -LiteralPath $psExe -PathType Leaf)) {
    $psExe = (Get-Command powershell.exe).Source
}

$desktop = [Environment]::GetFolderPath('Desktop')
$lnkPath = Join-Path $desktop 'Быстрые заметки.lnk'

$shell = New-Object -ComObject WScript.Shell
$lnk = $shell.CreateShortcut($lnkPath)
$lnk.TargetPath       = $psExe
$lnk.Arguments        = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $scriptPath
$lnk.WorkingDirectory = $dir
$lnk.WindowStyle      = 7   # свёрнуто: окно консоли не разворачивается на экран
$lnk.Description       = 'Быстрые заметки — текстовые файлы в один клик'
if (Test-Path -LiteralPath $iconPath -PathType Leaf) {
    $lnk.IconLocation = '{0},0' -f $iconPath
}
$lnk.Save()

Write-Host ''
Write-Host ("Ярлык создан: {0}" -f $lnkPath)
Write-Host 'Его можно закрепить на панели задач (правый клик по ярлыку).'
