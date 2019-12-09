#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Import-Module -Name "$scriptPath\ExcelData\ExcelData.psm1"
Import-Module -Name "$scriptPath\ExcelUtils\ExcelUtils.psm1"

Function Main() {
    $excelFilePath = "$scriptPath\..\$ExcelFilename"
    $incomingDir = "$scriptPath\..\$IncomingFolder"

    Write-Host "#" -ForegroundColor Yellow
    Write-Host "# Читаем текстовые файлы из Системы Город из папки $incomingDir" -ForegroundColor Yellow
    Write-Host "#" -ForegroundColor Yellow

    $excel = New-Object -ComObject excel.application
    $excel.visible = $true
    $workbook = $excel.Workbooks.Open($excelFilePath)

    $worksheet = GetSheet $workbook $IncomingLogSheetName
    if (-Not$worksheet) {
        exit
    }

    $groupsSheet = GetSheet $workbook $CommonListSheetName
    if (-Not$groupsSheet) {
        exit
    }

    $currentRow = $worksheet.UsedRange.SpecialCells($xlCellTypeLastCell).Row + 1
    $filesCount = 0
    $paymentAddedCount = 0
    $paymentCount = 0
    $skipedCount = 0
    Get-ChildItem $incomingDir -Filter *.txt |
            Foreach-Object {
                # Обработка файла
                $filesCount++
                $lines = Get-Content $_.FullName
                $metadata = New-Object System.Collections.ArrayList

                # Обработка строк в файле
                $lines | Foreach-Object {
                    if ($_ -Match '#') {
                        $value = $_.substring(1)
                        $metadata.Add($value) > $null
                    }

                    if ($_ -NotMatch '#') {
                        $errorMessage = WriteLine $_ $worksheet $groupsSheet $currentRow
                        if ($errorMessage) {
                            Write-Host $errorMessage -ForegroundColor Magenta
                            $skipedCount++
                        } else {
                            Write-Host "добавлена строка $currentRow" -ForegroundColor Green
                            $currentRow++
                            $paymentAddedCount++
                        }
                        $paymentCount++
                    }
                }
            }

    Write-Host "Итого:" -ForegroundColor Green
    Write-Host "Обработано $filesCount файлов, $paymentCount строк" -ForegroundColor Green
    Write-Host "Добавлено  $paymentAddedCount строк" -ForegroundColor Green
    Write-Host "Пропущено  $skipedCount строк" -ForegroundColor Green
    Write-Host "Лист '$IncomingLogSheetName' файла $excelFilePath" -ForegroundColor Magenta

    #saving & closing the file
    #adjusting the column width so all data's properly visible
    $excel.DisplayAlerts = $false
    # $workbook.SaveAs($outputpath, 51, [Type]::Missing, [Type]::Missing, $false, $false, 1, 2)

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($groupsSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Remove-Variable -Name excel

    Write-Host "Для завершения нажмите Enter" -ForegroundColor Blue
    Read-Host
}

Function WriteLine($line, $worksheet, $groupsSheet, $currentRow) {
    $parts = $line -split ";"

    $partCounter = 2

    $childName = $parts[$SGParts."childName"]
    $paymentId = $parts[$SGParts."paymentId"]
    $date = $parts[$SGParts."date"]

    $range = GetColumnRange($PaymentIdCell)
    $FindedCell = $worksheet.Range($range).EntireColumn.Find($paymentId)

    If ($FindedCell) {
        return "$paymentId - уже существует, строка $( $FindedCell.Row ), $childName";
    }
    $Found = $groupsSheet.Cells.Find($childName)

    $groupNumber = "?"
    $groupCell = $worksheet.Cells.Item($currentRow, $partCounter)
    if ($Found) {
        $groupNumber = $groupsSheet.Cells.Item($Found.Row, $GroupNumberCell).Text
    } else {
        Write-Host "$paymentId - не найдена группа $childName" -ForegroundColor Magenta
        $groupCell.Interior.ColorIndex = 3
    }
    $groupCell.Value = $groupNumber

    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter).Value2 = [datetime]::parseexact($date, 'dd/MM/yyyy', $null)

    # пишем все атрибуты платежа
    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $parts[0] # фио

    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $parts[1] # адрес

    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $parts[2] # номер договора

    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $parts[3] # сумма

    # обычно в этих ячейках пусто, но на всякий случай проверим
    if ($parts[4] -or $parts[5] -or $parts[6] -or $parts[7]){
        Write-Host "В строке $currentRow в пропущенных ячейках что-то есть!" -BackgroundColor DarkRed
    }

    # в 8 ячейке неформатированную дату пропускаем

    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $parts[9] # остаток задолженности

    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $parts[10] # номер квитанции

    # 10 индекс - пеня, пропускаем
    if ($parts[11] -ne "0.00"){
        Write-Host "В строке $currentRow какая-то странная пеня!" -BackgroundColor DarkRed
    }

    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $parts[12] # номер квитанции

    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $parts[13].Text # номер транзакции

    # с этой позиции начинаем добавлять метаданные
    $partCounter++
    $m = $metadata[0] -split ";" # номер реестра
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[0]
    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[1].Replace("Номер реестра", "")

    $partCounter++
    $m = $metadata[1] -split ";" # сумма реестра
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[0] # описание пропускаем

    $partCounter++
    $m = $metadata[2] -split ";" # пеня
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[0] # описание пропускаем

    $partCounter++
    $m = $metadata[3] -split ";" # удержанная сумма
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[0] # описание пропускаем

    $partCounter++
    $m = $metadata[4] -split ";" # сумма к перечислению
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[0] # описание пропускаем

    $partCounter++
    $m = $metadata[5] -split ";" # число записей в Реестре
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[0] # описание пропускаем

    $partCounter++
    $m = $metadata[6] -split ";" # код агента
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[0]
    $partCounter++
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[1].Replace("Код агента", "")

    $partCounter++
    $m = $metadata[7] -split ";" # номер услуги
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[0] # описание пропускаем

    $partCounter++
    $m = $metadata[8] -split ";" # дата формирования реестра
    $worksheet.Cells.Item($currentRow, $partCounter).Value2 = [datetime]::ParseExact($m[0], 'dd/MM/yyyy HH:mm:ss', $null)  # TODO не работает описание пропускаем  04/09/2019 13:25:04

    # начало и конец диапазона дат документов, входящих в реестр пропускаем

    # $partCounter++
    $m = $metadata[11] -split ";" # Примечание
    $worksheet.Cells.Item($currentRow, $partCounter) = $m[0] # описание пропускаем

}

Main



