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
    Write-Host "Месяц для начислений [1-12]" -ForegroundColor Green
    [uint16] $month = Read-Host
    $monthName = GetMonthName $month
    Write-Host "$monthName"

    Write-Host "Номер группы [1-12]" -ForegroundColor Green
    $group = Read-Host

    $excelFilePath = "$scriptPath\$ExcelFilename"
    $outcomingDir = "$scriptPath\$OutcomingFolder"
    $groupFilePath = "$scriptPath\$($GroupExcel.Replace("N", $group) )"

    if (-Not [System.IO.File]::Exists($excelFilePath)) {
        Write-Host "Файл не найден $excelFilePath"
        exit
    }

    if (-Not [System.IO.File]::Exists($groupFilePath)) {
        Write-Host "Файл не найден $groupFilePath"
        exit
    }

    $excel = New-Object -ComObject excel.application
    $excel.visible = $false
    $excel.DisplayAlerts = $false
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

    $newFile = "$outcomingDir\$($OutcomingFilename.Replace("N", $group).Replace("M", $monthName) )"
    if (-Not [System.IO.File]::Exists($newFile)){
        New-Item $newFile | Out-Null
    }
    Add-Content $newFile "#TYPE 7"
    Add-Content $newFile "#SERVICE 40334"

    $overalSum = 0
    $rows = 0

    for ($i = $startRow; $i -le $endRow; $i++)
    {
        $CommonFondSum = $groupsheet.Cells.Item($i, $CommonFondColumn).Value2
        $GroupFondSum = $groupsheet.Cells.Item($i, $GroupFondColumn).Value2
        $Name = $groupsheet.Cells.Item($i, $NameColumn).Text
        if ( $Name.StartsWith("#")) {
            Write-Host "Строка $i пропущена, т.к. начинается с #"
            continue
        }
        $Adress = FindAdress $Name $commonListSheet
        if (-Not$Adress) {
            continue
        }
        $Contract = FindContract $Name $commonListSheet
        if (-Not$Contract) {
            continue
        }
        if ($CommonFondSum -and $GroupFondSum -and $Name) {
            Add-Content $newFile "$Name;$Adress;$Contract;$( $CommonFondSum + $GroupFondSum )"
            $rows++
            $overalSum = $overalSum + $CommonFondSum + $GroupFondSum
        }
    }

    Add-Content $newFile "#FILESUM $overalSum"

    Write-Host "Итого:"
    Write-Host "Добавлено  $rows строк" -ForegroundColor Green
    Write-Host "Общая сумма  $overalSum" -ForegroundColor Green

    #saving & closing the file
    #adjusting the column width so all data's properly visible
    $usedRange = $groupsheet.UsedRange
    $excel.DisplayAlerts = $false
    $excel.Quit()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($groupsheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Remove-Variable -Name excel
}

Function FindAdress($Name, $commonListSheet) {
    $FindedCell = $commonListSheet.Range($NameCell2 + ":" + $NameCell2).EntireColumn.Find($Name)
    $Address = $commonListSheet.Cells.Item($FindedCell.Row, $AddressCell).Value2
    if (-Not $Address){
        Write-Host "Адрес не найден для $Name" -ForegroundColor Magenta
        return $null
    }
    return $Address
}

Function FindContract($Name, $commonListSheet) {
    $column =  GetColumn($NameCell)
    $FindedCell = $commonListSheet.Range("$($column):$($column)").EntireColumn.Find($Name)
    $Contract = $commonListSheet.Cells.Item($FindedCell.Row, $ContractCell).Value2
    if (-Not $Contract){
        Write-Host "Номер договора не найден для $Name"
        return $null
    }
    return $Contract
}
Main



