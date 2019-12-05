[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

$excelFilePath = "$( $scriptPath )\2019.xlsx" # Update source File Path

Write-Host "$( $excelFilePath )"

$month = Read-Host -Prompt 'Месяц для начислений'
$group = Read-Host -Prompt 'Номер группы'

$excel = New-Object -ComObject excel.application
$excel.visible = $true
$workbook = $excel.Workbooks.Open($excelFilePath)

$listSheet = $Workbook.Worksheets | Where-Object {
    $_.name -eq "Общий список"
}

$groupSheet = $Workbook.Worksheets | Where-Object {
    $_.name -eq "$group группа"
}

$xlCellTypeLastCell = 11
$startRow = 4
$endRow = $listSheet.UsedRange.SpecialCells($xlCellTypeLastCell).Row
Write-Host "$endRow"

$monthRange = $groupSheet.Range("A4:Z4")
$monthCell = $monthRange.Find($month)

$childRange = $groupSheet.Range("B:B")

$billCount = 0
for ($i = $startRow; $i -le $endRow; $i++)
{
    if ($listSheet.Cells.Item($i, 3).Text -eq $group) {
        $child = $listSheet.Cells.Item($i, 2).Text
        $foundChild = $childRange.Find($child)

        if($foundChild){
            $groupSheet.Cells.Item($foundChild.Row, $monthCell.Column) = $listSheet.Cells.Item($i, 4).Text
            $billCount++
        }else{
            Write-Host "$child не найден"
        }
    }
}

Write-Host "Итого:"
Write-Host "Добавлено $billCount начислений" -ForegroundColor Green

#saving & closing the file
$excel.DisplayAlerts = $false

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($listSheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Remove-Variable -Name excel


