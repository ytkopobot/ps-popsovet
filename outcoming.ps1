#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

$excelFilePath = "$( $scriptPath )\2019.xlsx" # Update source File Path
$outcomingDir = "$( $scriptPath )\outcoming"

Write-Host "$( $excelFilePath )"

$excel = New-Object -ComObject excel.application
$excel.visible = $true
$workbook = $excel.Workbooks.Open($excelFilePath)

$groupsSheet = $Workbook.Worksheets | Where-Object {$_.name -eq "Группы"}

$xlCellTypeLastCell = 11
$startRow = 4
$endRow = $groupsSheet.UsedRange.SpecialCells($xlCellTypeLastCell).Row

$newFileName = Get-Date -Format "yyMMdd-HHmm"
$newFile = "$($outcomingDir)\$($newFileName).txt"
New-Item $newFile
for ($i = $startRow; $i -le $endRow; $i++)
{
    $name = $groupsSheet.Cells.Item($i, 1).Text
    $group = $groupsSheet.Cells.Item($i, 2).Text
    $payment = $groupsSheet.Cells.Item($i, 3).Text
    Add-Content $newFile $name
    Add-Content $newFile $payment
}

#saving & closing the file
$excel.DisplayAlerts = $false
$excel.Quit()

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($groupsSheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

Remove-Variable -Name excel

