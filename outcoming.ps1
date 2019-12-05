#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

Import-Module -Name "$scriptPath\ExcelData\ExcelData.psm1"

$excelFilePath = "$scriptPath\$ExcelFilename"
$outcomingDir = "$scriptPath\$OutcomingFolder"

$excel = New-Object -ComObject excel.application
$excel.visible = $true
$workbook = $excel.Workbooks.Open($excelFilePath)

$groupsSheet = $Workbook.Worksheets | Where-Object {
    $_.name -eq $CommonListSheetName
}
if(-Not $groupsSheet){
    Write-Host "Не найден лист $CommonListSheetName"
    exit
}

$endRow = $groupsSheet.UsedRange.SpecialCells($xlCellTypeLastCell).Row

for($i = 1; $i -le 12; $i++){
    New-Item "$outcomingDir\$($OutcomingFilename.Replace("N", $i)).txt"
}

$currentSum = 0
for ($i = 0; $i -le $endRow; $i++)
{
    $newFile = "$outcomingDir\123.txt"
    New-Item $newFile

    $groupTitle = $groupsSheet.Cells.Item($i, $GroupTitleCell).Text
    If($groupTitle.StartsWith("#")){
        # завершаем предыдущие подсчеты и создаём новый файл, обнуляем суммы
        # https://powershellexplained.com/2018-10-15-Powershell-arrays-Everything-you-wanted-to-know/
        $currentSum = 0

    }
    $name = $groupsSheet.Cells.Item($i, $NameCell).Text
    $group = $groupsSheet.Cells.Item($i, $GroupNumberCell).Text
    $payment = $groupsSheet.Cells.Item($i, $PaymentCell).Text
    Add-Content $newFile $name
    Add-Content $newFile $payment
}

Write-Host "Итого:"
Write-Host "Обработано $( $endRow - $startRow + 1 ) чел." -ForegroundColor Green

#saving & closing the file
$excel.DisplayAlerts = $false
$excel.Quit()

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($groupsSheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Remove-Variable -Name excel

