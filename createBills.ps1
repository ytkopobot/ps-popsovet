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
    $month = Read-Host
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
    $excel.visible = $true
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
    New-Item $newFile
    Add-Content $newFile "#TYPE 7"
    Add-Content $newFile "#SERVICE 40334"

    $overalSum = 0
    $rows = 0

    for ($i = $startRow; $i -le $endRow; $i++)
    {
        $CommonFondSum = $groupsheet.Cells.Item($i, $CommonFondColumn).Value2
        $GroupFondSum = $groupsheet.Cells.Item($i, $GroupFondColumn).Value2
        $Name = $groupsheet.Cells.Item($i, $NameColumn).Text
        $Adress = FindAdress
        if(-Not $Adress){
            continue
        }
        $Contract = FindContract
        if(-Not $Contract){
            continue
        }
        if ($Name.StartsWith("#")){
            Write-Host "Строка $i пропущена, т.к. начинается с #"
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

    #saving & closing the file
    #adjusting the column width so all data's properly visible
    $usedRange = $groupsheet.UsedRange
    $excel.DisplayAlerts = $false

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($groupsheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Remove-Variable -Name excel
}

Function FindAdress(){
    return "Сиреневая"
}

Function FindContract(){
    return "123"
}
Main



