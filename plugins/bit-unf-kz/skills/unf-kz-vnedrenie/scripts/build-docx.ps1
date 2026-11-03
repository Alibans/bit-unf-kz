# Собирает .docx на бланке «Первый БИТ» из .html через Microsoft Word (отдельный скрытый экземпляр).
# Использование: powershell -ExecutionPolicy Bypass -File build-docx.ps1 "<вход.html>" "<выход.docx>" [-Pdf]
# -Pdf — рядом сохранить PDF. Ширины колонок таблиц задаются по названию колонки в шапке (словарь $weights ниже);
# для новой таблицы с другими названиями колонок — дописать их веса в словарь.
param(
    [Parameter(Mandatory = $true)][string]$HtmlPath,
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [string]$BlankPath = '',
    [switch]$Pdf
)
$ErrorActionPreference = 'Stop'
if (-not $BlankPath) { $BlankPath = Join-Path $PSScriptRoot '..\assets\бланк Первый БИТ.docx' }
$HtmlPath = (Resolve-Path $HtmlPath).Path
$BlankPath = (Resolve-Path $BlankPath).Path
$DocxPath = [IO.Path]::GetFullPath($DocxPath)
$brand = 226 + 0 * 256 + 110 * 65536   # #E2006E в формате Word (BGR)

try { $word = New-Object -ComObject Word.Application }
catch {
    Write-Output "Word не найден. Откройте $HtmlPath в Word и сохраните как .docx."
    exit 2
}
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add($BlankPath)
    $doc.Content.Delete() | Out-Null
    $doc.Content.InsertFile($HtmlPath)

    # Стили бланка: Calibri 11, заголовки фирменным цветом
    $st = $doc.Styles
    $st.Item(-1).Font.Name = 'Calibri'; $st.Item(-1).Font.Size = 11          # Обычный
    foreach ($id in -2, -3, -4) {                                               # Заголовки 1–3
        $h = $st.Item($id); $h.Font.Name = 'Calibri'; $h.Font.Color = $brand; $h.Font.Bold = $true
    }

    # Ширины колонок по названию в шапке (Word при вставке HTML их не берёт); неизвестная колонка — вес 20
    $weights = @{
        '№' = 6; 'Функция' = 27; 'Как сейчас' = 20; 'Как в 1С' = 33; 'Заключение' = 17; 'Комментарий' = 20
        'Дата' = 15; 'Что смотрели' = 50; 'Участники' = 35
        'Что нужно Заказчику' = 30; 'Как в типовой' = 28; 'Что предлагаем' = 36
        'Что передать или решить' = 52; 'Строки реестра' = 13; 'Кто' = 18; 'Срок' = 11
        'Роль' = 34; 'ФИО' = 34; 'Подпись' = 16
        'Требование Заказчика' = 29; 'Реализация в Программе' = 31; 'Примечание' = 18
        'Термин' = 22; 'Определение' = 78
        'Бизнес-процесс' = 22; 'Как выполняется сейчас' = 29; 'Исполнитель' = 18; 'Периодичность' = 17; 'Приоритет' = 13
        'Вопрос' = 52; 'Кто отвечает' = 17; 'Ответ' = 27; 'Критерий' = 31; '1С:Розница' = 23; '1С:Управление торговлей' = 23; '1С:Управление нашей фирмой' = 23
        'Наименование' = 40; 'Содержание' = 42; 'Основание' = 14; 'Ответственный' = 18
        'Зачем нам' = 36; 'Маршрут' = 18; 'Признаки' = 44; 'Что дальше' = 32; 'Развилка' = 34; 'Как в программе' = 50
        'Работа' = 60; 'Ориентир' = 30; 'Документ' = 30; 'Кому' = 14; 'Что в нём' = 50; 'Шаг' = 52; 'Что на выходе' = 30; 'Участок' = 16; 'Что есть в программе' = 46; 'Что спросить у клиента' = 38
    }
    $ps = $doc.PageSetup
    $ps.BottomMargin = 85                   # 3 см: плашка 1bit.kz в нижнем колонтитуле выше поля бланка (2 см)
    $total = $ps.PageWidth - $ps.LeftMargin - $ps.RightMargin - 2
    foreach ($t in $doc.Tables) {
        $t.AllowAutoFit = $false
        $t.Range.ParagraphFormat.SpaceBefore = 0
        $t.Range.ParagraphFormat.SpaceAfter = 0
        $t.Range.ParagraphFormat.LineSpacingRule = 0   # одинарный
        $t.TopPadding = 1; $t.BottomPadding = 1; $t.LeftPadding = 3; $t.RightPadding = 3
        $t.Rows.AllowBreakAcrossPages = $false   # строка целиком на одной странице
        $t.Rows.Item(1).HeadingFormat = -1  # шапка таблицы повторяется на каждой странице

        $head = $t.Rows.Item(1).Cells
        $n = $head.Count
        $w = @()
        for ($i = 1; $i -le $n; $i++) {
            $name = ($head.Item($i).Range.Text -replace '[\r\a\n]', '').Trim()
            if ($weights.ContainsKey($name)) { $w += $weights[$name] } else { $w += 20 }
        }
        $sum = ($w | Measure-Object -Sum).Sum
        $pt = $w | ForEach-Object { $total * $_ / $sum }
        foreach ($row in $t.Rows) {
            $cells = $row.Cells
            if ($cells.Count -eq $n) {
                for ($i = 1; $i -le $n; $i++) { $cells.Item($i).Width = $pt[$i - 1] }
            }
            elseif ($cells.Count -eq 1) { $cells.Item(1).Width = $total; $row.Range.ParagraphFormat.KeepWithNext = $true }
            else {
                # строка группы: первая колонка как в шапке, остальное — одной ячейкой; не отрывать от следующей строки
                $row.Range.ParagraphFormat.KeepWithNext = $true
                $cells.Item(1).Width = $pt[0]
                for ($i = 2; $i -le $cells.Count; $i++) { $cells.Item($i).Width = ($total - $pt[0]) / ($cells.Count - 1) }
            }
        }
    }

    $doc.SaveAs2($DocxPath, 16)             # wdFormatDocumentDefault
    if ($Pdf) { $doc.ExportAsFixedFormat([IO.Path]::ChangeExtension($DocxPath, '.pdf'), 17) }
    Write-Output "Готово: $DocxPath, страниц: $($doc.ComputeStatistics(2))"
    $doc.Close(0)
}
finally {
    $word.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
