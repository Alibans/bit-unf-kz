# Выполняет код 1С из файла .bsl в базе отдельным сеансом (через ВыполнитьКодИзФайла.epf рядом со скриптом).
# Итог переменной Результат — в <файл>.out.txt (первая строка OK или ERROR), он же выводится на экран.
# Примеры:
#   .\run-code.ps1 -CodeFile C:\work\model\10-nsi.bsl -Server localhost -Base unf_model -User Администратор
#   .\run-code.ps1 -CodeFile C:\work\model\10-nsi.bsl -FileBase "C:\Bases\unf_model" -User Администратор
param(
    [Parameter(Mandatory)][string]$CodeFile,
    [string]$Server = '',          # сервер 1С (для серверной базы)
    [string]$Base = '',            # имя базы на сервере
    [string]$FileBase = '',        # каталог файловой базы (вместо Server/Base)
    [string]$User = '',
    [string]$Password = '',
    [string]$Platform = '',        # путь к 1cv8c.exe; по умолчанию — самая новая версия в Program Files.
                                   # Для серверной базы версия клиента должна совпадать с версией сервера 1С — тогда укажите явно
    [int]$TimeoutSec = 600
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$epf = Join-Path $PSScriptRoot 'ВыполнитьКодИзФайла.epf'
$CodeFile = (Resolve-Path $CodeFile).Path

# Файл кода — в UTF-8 с BOM (иначе кириллица в 1С ломается)
$text = [IO.File]::ReadAllText($CodeFile, [Text.Encoding]::UTF8)
[IO.File]::WriteAllText($CodeFile, $text, (New-Object Text.UTF8Encoding($true)))

if (-not $Platform) {
    $Platform = Get-ChildItem 'C:\Program Files\1cv8\*\bin\1cv8c.exe', 'C:\Program Files (x86)\1cv8\*\bin\1cv8c.exe' -ErrorAction SilentlyContinue |
        Sort-Object { [version]($_.Directory.Parent.Name -replace '[^\d\.]', '') } -Descending |
        Select-Object -First 1 -ExpandProperty FullName
    if (-not $Platform) { throw 'Не найден 1cv8c.exe — укажите -Platform' }
}

if ($FileBase) { $conn = '/F "' + $FileBase + '"' }
elseif ($Server -and $Base) { $conn = '/S "' + $Server + '\' + $Base + '"' }
else { throw 'Укажите -Server и -Base или -FileBase' }

$out = "$CodeFile.out.txt"; if (Test-Path $out) { Remove-Item $out }
$a = 'ENTERPRISE ' + $conn
if ($User) { $a += ' /N "' + $User + '" /P "' + $Password + '"' }
$a += ' /Execute "' + $epf + '" /C "' + $CodeFile + '" /DisableStartupDialogs /DisableStartupMessages /Out "' + $CodeFile + '.log"'

$p = Start-Process $Platform -ArgumentList $a -PassThru
if (-not $p.WaitForExit($TimeoutSec * 1000)) { Stop-Process -Id $p.Id -Force; 'TIMEOUT'; exit 2 }
if (Test-Path $out) { Get-Content $out -Raw -Encoding UTF8 }
else { "NO OUTPUT (exit $($p.ExitCode))"; if (Test-Path "$CodeFile.log") { Get-Content "$CodeFile.log" -Raw } }
