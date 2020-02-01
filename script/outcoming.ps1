#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Import-Module -Name "$scriptPath\ExcelData\ExcelData.psm1"
Import-Module -Name "$scriptPath\ExcelUtils\ExcelUtils.psm1"


#
# Выставляем начисления в Систему Город в виде текстовых файлов
#
Function Main() {
    Write-Host "#" -ForegroundColor Yellow
    Write-Host "# Создаём текстовые файлы для загрузки в Систему Город " -ForegroundColor Yellow
    Write-Host "#" -ForegroundColor Yellow
    Write-Host "Месяц для начислений [1-12]" -ForegroundColor Green
    [uint16] $month = Read-Host
    $monthName = GetMonthName $month
    Write-Host "$monthName"

    if (-Not$monthName) {
        exit
    }

    Write-Host "Номер группы [1-12]" -ForegroundColor Green
    $groupNumber = Read-Host

    $excelFilePath = "$scriptPath\..\$ExcelFilename"
    if (-Not [System.IO.File]::Exists($excelFilePath)) {
        Write-Host "Файл $ExcelFilename не найден $excelFilePath" -ForegroundColor Red
        exit
    }

    $groupFilePath = "$scriptPath\..\$($GroupExcel.Replace("N", $groupNumber) )"
    if (-Not [System.IO.File]::Exists($groupFilePath)) {
        Write-Host "Файл группы не найден $groupFilePath" -ForegroundColor Red
        exit
    }

    $outcomingDir = "$scriptPath\..\$OutcomingFolder"

    $excel = New-Object -ComObject excel.application
    $excel.visible = $true
    $excel.DisplayAlerts = $true
    $workbook = $excel.Workbooks.Open($excelFilePath)
    $groupbook = $excel.Workbooks.Open($groupFilePath)

    $groupsheet = GetSheet $groupbook $GroupSheetName
    if (-Not$groupsheet) {
        exit
    }

    $commonListSheet = GetSheet $workbook $CommonListSheetName
    if (-Not$commonListSheet) {
        exit
    }

    $startRow = $GroupStartRow
    $endRow = $groupsheet.UsedRange.SpecialCells($xlCellTypeLastCell).Row

    $newFile = "$outcomingDir\$($OutcomingFilename.Replace("N", $groupNumber).Replace("M", $monthName) )"
    if (-Not [System.IO.File]::Exists($newFile)) {
        New-Item $newFile | Out-Null
    }

    $overalSum = 0
    $skips = 0
    $n = 1
    Write-Host ""
    $lines = New-Object System.Collections.ArrayList

    $FindedMonth = $groupSheet.Cells.Item($GroupMonthRow, 1).EntireRow.Find($monthName)
    if (-Not$FindedMonth) {
        Write-Host "Для группы на листе $GroupSheetName в строке $GroupMonthRow не найден месяц $monthName" -ForegroundColor Magenta
        exit
    }

    $groupSheet.Columns($FindedMonth.Column).Hidden = $false
    $groupSheet.Columns($FindedMonth.Column + 1).Hidden = $false
    $groupSheet.Columns($FindedMonth.Column + 2).Hidden = $false

    for ($i = $startRow; $i -le $endRow; $i++)
    {
        $CommonFondSum = $groupsheet.Cells.Item($i, $CommonFondColumn).Value2
        $GroupFondSum = $groupsheet.Cells.Item($i, $GroupFondColumn).Value2
        $currentDebt = $groupsheet.Cells.Item($i, $DebtColumn).Value2
        $Name = $groupsheet.Cells.Item($i, $NameColumn).Value2
        $Tag = $groupsheet.Cells.Item($i, $TagColumn).Value2
        if (-Not$Name) {
            continue;
        }
        if ( $Name.StartsWith("#")) {
            continue
        }
        if($Tag -eq "н") {
            Write-Host "$n. $Name   начисление пропущено   указана 'н'" -ForegroundColor DarkGray
            $skips++
            $n++
            continue
        }

        $Adress = FindAdress $Name $commonListSheet
        if (-Not$Adress) {
            Write-Host "$n. $Name   начисление пропущено   адрес не найден" -ForegroundColor Magenta
            $skips++
            $n++
            continue
        }
        $Contract = FindContract $Name $commonListSheet
        if (-Not$Contract) {
            Write-Host "$n. $Name   начисление пропущено   номер договора не найден" -ForegroundColor Magenta
            $skips++
            $n++
            continue
        }

        if($Tag -eq 0){
            Write-Host "$n. $Name   начислено 0   строка помечена 0" -ForegroundColor Cyan
            $lines.Add("$Name;$Adress;$Contract;0.00") > $null
            $n++
            continue;
        }


        # Делаем начисления в указанный месяц
        $currentMonthOutcomeCell = $groupSheet.Cells.Item($i, $FindedMonth.Column)
        if(-Not$currentMonthOutcomeCell.Value2){
            $currentMonthOutcomeCell.Value2 = $CommonFondSum + $GroupFondSum
        }

        if ((-Not$currentDebt) -or ($currentDebt -eq 0)) {
            Write-Host "$n. $Name   начислено 0   долга нет" -ForegroundColor Cyan
            $lines.Add("$Name;$Adress;$Contract;0.00") > $null
            $n++
            continue;
        }

        if ($currentDebt -lt 0) {
            Write-Host "$n. $Name   начислено 0   долг отрицательный" -ForegroundColor Cyan
            $lines.Add("$Name;$Adress;$Contract;0.00") > $null
            $n++
            continue;
        }

        $currentSum = $currentDebt
        $fondSum = $CommonFondSum + $GroupFondSum
        if ($fondSum -le 0) {
            Write-Host "$n. $Name    начислено $currentSum   сумма взноса не установлена, начислен только долг"  -ForegroundColor Cyan
        } else {
            if ($currentDebt -gt $fondSum * 3) {
                $currentSum = $fondSum * 3
                Write-Host "$n. $Name   начислено $currentSum   сумма долга превышает трёхмесячный размер платежей, необходимо списать " -ForegroundColor Cyan  -nonewline
                Write-Host "$( $currentDebt - $fondSum * 3 )" -BackgroundColor DarkRed  -ForegroundColor White
            }else{
                Write-Host "$n. $Name   начислено $currentSum" -ForegroundColor Green
            }
        }

        $formatted = FormatNumber $currentSum
        $lines.Add("$Name;$Adress;$Contract;$formatted") > $null
        $overalSum = $overalSum + $currentSum
        $n++
    }

    Clear-Content $newFile
    Add-Content $newFile "#FILESUM $(FormatNumber $overalSum)"
    Add-Content $newFile "#TYPE 7"
    Add-Content $newFile "#SERVICE 40334"

    foreach ($element in $lines) {
        Add-Content $newFile $element
    }

    Write-Host
    Write-Host "Итого:" -BackgroundColor Green
    Write-Host "Добавлено  $($lines.Count) строк" -ForegroundColor Green
    Write-Host "Пропущено  $skips строк" -ForegroundColor Green
    Write-Host "Общая сумма  $overalSum" -ForegroundColor Green
    Write-Host "Результат $newFile" -ForegroundColor Green

    #saving & closing the file
    #adjusting the column width so all data's properly visible
    $excel.DisplayAlerts = $false

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($groupsheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($commonListSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Remove-Variable -Name excel

    Pause
}

Function FindAdress($Name, $commonListSheet) {
    $range = GetColumnRange($NameCell)
    $FindedCell = $commonListSheet.Range($range).EntireColumn.Find($Name)
    $Address = $commonListSheet.Cells.Item($FindedCell.Row, $AddressCell).Value2
    if (-Not$Address) {
        return $null
    }
    return $Address
}

Function FindContract($Name, $commonListSheet) {
    $range = GetColumnRange($NameCell)
    $FindedCell = $commonListSheet.Range($range).EntireColumn.Find($Name)
    $Contract = $commonListSheet.Cells.Item($FindedCell.Row, $ContractCell).Value2
    if (-Not$Contract) {
        return $null
    }
    return $Contract
}

Function FormatNumber($number){
    return $([double]$number).ToString("#.00").Replace(",", ".")
}
Main



