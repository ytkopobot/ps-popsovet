#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Import-Module -Name "$scriptPath\ExcelData\ExcelData.psm1"
Import-Module -Name "$scriptPath\ExcelUtils\ExcelUtils.psm1"

Function Main() {
    $excelFilePath = "$scriptPath\..\$ExcelFilename"

    Write-Host "#" -ForegroundColor Yellow
    Write-Host "# Обрабатываем полученные взносы из листа '$IncomingLogSheetName' файла $excelFilePath" -ForegroundColor Yellow
    Write-Host "# Записываем их в лист $GroupSheetName соответствующей группы" -ForegroundColor Yellow
    Write-Host "# Месяц вычисляем из даты платежа, введённые значения отмечаются жёлтым фоном, т.к. требуют проверки." -ForegroundColor Yellow
    Write-Host "#" -ForegroundColor Yellow

    if (-Not [System.IO.File]::Exists($excelFilePath)) {
        Write-Host "Файл не найден $excelFilePath"
        exit
    }

    $excel = New-Object -ComObject excel.application
    $excel.visible = $true
    $excel.DisplayAlerts = $true
    $workbook = $excel.Workbooks.Open($excelFilePath)

    $worksheet = GetSheet $workbook $IncomingLogSheetName
    if (-Not$worksheet) {
        exit
    }

    Write-Host "Введите номер строки из листа '$IncomingLogSheetName' с которой начнётся обработка" -ForegroundColor Green
    [uint16] $startRow = Read-Host

    $endRow = $worksheet.UsedRange.SpecialCells($xlCellTypeLastCell).Row

    $allGroupSheets = @{}
    $incomingsCount = 0
    $addedIncomings = 0
    $editedFilesCount = @{}

    for ($i = $startRow; $i -le $endRow; $i++)
    {
        $incomingName = $worksheet.Cells.Item($i, $IncomingNameCell).Text
        if (-Not$incomingName) {
            continue
        }
        $incomingsCount++
        [uint16] $groupNumber = $worksheet.Cells.Item($i, $IncomingGroupCell).Value2
        if (-Not$groupNumber) {
            Write-Host "Не найдена группа для $incomingName, строка $i колонка $IncomingGroupCell" -ForegroundColor Magenta
            continue
        }

        $payment = $worksheet.Cells.Item($i, $IncomingPaymentCell).Value2
        if (-Not$payment) {
            Write-Host "Не найдена сумма платежа для $incomingName, строка $i колонка $IncomingPaymentCell" -ForegroundColor Magenta
            continue
        }

        $groupSheet = $allGroupSheets["$groupNumber"]
        if (-Not$groupSheet) {
            $groupSheet = $allGroupSheets["$groupNumber"] = OpenGroup $excel $groupNumber
            if (-Not$groupSheet) {
                Write-Host "Невозможно записать взнос для $incomingName группа $groupNumber, строка $i - файл группы не доступен" -ForegroundColor Magenta
                continue
            }
        }
        $paymentDate = $worksheet.Cells.Item($i, $IncomingPaymentDateCell).Text
        if (-Not$paymentDate) {
            Write-Host "Дата взноса не известна $incomingName, строка $i колонка $IncomingPaymentDateCell" -ForegroundColor Magenta
            continue
        }
        $monthNumber = [datetime]::parseexact($paymentDate, 'dd.MM.yyyy', $null).Month
        $monthName = GetMonthName $monthNumber


        $FindedMonth = $groupSheet.Cells.Item($GroupMonthRow, 1).EntireRow.Find($monthName)
        if (-Not$FindedMonth) {
            Write-Host "Для группы $groupNumber на листе $GroupSheetName в строке $GroupMonthRow не найден месяц $monthName, взнос для $incomingName будет пропущен" -ForegroundColor Magenta
            continue
        }

        $FindedName = $groupSheet.Cells.Item(1, $NameColumn).EntireColumn.Find($incomingName)
        if (-Not$FindedName) {
            Write-Host "Для группы $groupNumber на листе $GroupSheetName в колонке $NameColumn не найдено фио $incomingName, взнос будет пропущен" -ForegroundColor Magenta
            continue
        }

        $currentCommonFondValueCell = $groupSheet.Cells.Item($FindedName.Row, $FindedMonth.Column + 1)
        $currentCommonFondValue = $currentCommonFondValueCell.Value2

        $currentGroupValueCell = $groupSheet.Cells.Item($FindedName.Row, $FindedMonth.Column + 2)
        $currentGroupValue = $currentGroupValueCell.Value2

        if ($currentCommonFondValue -or $currentGroupValue) {
            $currentCommonFondValueCell.Interior.ColorIndex = 3
            $currentGroupValueCell.Interior.ColorIndex = 3
            Write-Host ""
            Write-Host "Для группы $groupNumber для $incomingName в месяце $monthName уже есть значение: $currentCommonFondValue и $currentGroupValue, взнос $payment будет пропущен!" -ForegroundColor Red
            Write-Host ""
            continue
        }

        $CommonFondSum = $groupsheet.Cells.Item($FindedName.Row, $CommonFondColumn).Value2
        $GroupFondSum = $groupsheet.Cells.Item($FindedName.Row, $GroupFondColumn).Value2


        if ($payment -eq ($CommonFondSum + $GroupFondSum)){
            Write-Host "Для $incomingName $groupNumber гр. в месяце '$monthName' сумма равна ежемесячному взносу, поэтому она сразу будет разбита на две колонки" -ForegroundColor Green
            $currentCommonFondValueCell.Interior.ColorIndex = 4
            $currentCommonFondValueCell.Value2 = $CommonFondSum

            $currentGroupValueCell.Interior.ColorIndex = 4
            if($GroupFondSum){
                $currentGroupValueCell.Value2 = $GroupFondSum
            }

        }else {
            Write-Host "Для $incomingName $groupNumber гр. в месяце '$monthName' сумма не равна ежемесячному взносу, поэтому она будет записана в фонд сада для дальнейшей ручной разбивки" -ForegroundColor Cyan
            $currentCommonFondValueCell.Interior.ColorIndex = 6
            $currentCommonFondValueCell.Value2 = $payment
        }

        $addedIncomings++

        if (-Not$editedFilesCount["$groupNumber"]) {
            $editedFilesCount["$groupNumber"] = 0
        }
        $editedFilesCount["$groupNumber"] = $editedFilesCount["$groupNumber"] + 1
    }

    Write-Host
    Write-Host "Итого:" -BackgroundColor Green
    Write-Host "Обработано $incomingsCount строк, с $startRow по $endRow" -ForegroundColor Green
    Write-Host "Добавлено  $addedIncomings взносов" -ForegroundColor Green
    Write-Host "Открыто  $( $allGroupSheets.count ) файлов групп" -ForegroundColor Green
    Write-Host "Изменено  $( $editedFilesCount.count ) файлов групп" -ForegroundColor Green

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
    for ($i = 1; $i -le 12; $i++){
        if ($allGroupSheets["$i"]) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($allGroupSheets["$i"]) | Out-Null
        }
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Remove-Variable -Name excel
    Pause

}

Function OpenAllGroupFiles($excel) {
    $hash = @{}
    for ($i = 1; $i -le 12; $i++){
        $hash["$i"] = OpenGroup $excel $i
    }
    Write-Host $hash
    return $hash
}

Function OpenGroup($excel, $groupNumber) {
    $groupFilePath = "$scriptPath\..\$($GroupExcel.Replace("N", $groupNumber) )"
    if (-Not [System.IO.File]::Exists($groupFilePath)) {
        Write-Host "Файл не найден $groupFilePath"
        return $null
    }
    if (CheckFileOpen $groupFilePath) {
        exit
    }
    $groupbook = $excel.Workbooks.Open($groupFilePath)
    $groupsheet = GetSheet $groupbook $GroupSheetName
    if (-Not$groupsheet) {
        return $null
    }
    return $groupsheet
}

Main



