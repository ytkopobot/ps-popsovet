#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

$excelFilePath = "$( $scriptPath )\2019.xlsx" # Update source File Path
$fileDir = "$( $scriptPath )\incoming"

Write-Host "$( $excelFilePath )"

$excel = New-Object -ComObject excel.application
$excel.visible = $true
$workbook = $excel.Workbooks.Open($excelFilePath)

$worksheet = $workbook.Worksheets.Item(1)
$groupsSheet = $Workbook.Sheets.Item(2)

$row = 1
$Column = 1
$worksheet.Cells.Item($row, $column) = 'Title'

#merging a few cells on the top row to make the title look nicer
$MergeCells = $worksheet.Range("A1:G1")
$MergeCells.Select()
$MergeCells.MergeCells = $true
$worksheet.Cells(1, 1).HorizontalAlignment = -4108

$worksheet.Cells.Item(1, 1).Font.Size = 18
$worksheet.Cells.Item(1, 1).Font.Bold = $True
$worksheet.Cells.Item(1, 1).Font.Name = "Cambria"
$worksheet.Cells.Item(1, 1).Font.ThemeFont = 1
$worksheet.Cells.Item(1, 1).Font.ThemeColor = 4
$worksheet.Cells.Item(1, 1).Font.ColorIndex = 55
$worksheet.Cells.Item(1, 1).Font.Color = 8210719

#create the column headers
$worksheet.Cells.Item(3, 1) = 'Группы'
$worksheet.Cells.Item(3, 2) = 'Имя'
$worksheet.Cells.Item(3, 3) = 'Адрес'
$worksheet.Cells.Item(3, 4) = 'Номер договора'
$worksheet.Cells.Item(3, 5) = 'Сумма'
$worksheet.Cells.Item(3, 10) = 'Дата'
$worksheet.Cells.Item(3, 15) = '??'

$currentRow = 4
Get-ChildItem $fileDir -Filter *.txt |
        Foreach-Object {
            $lines = Get-Content $_.FullName
            $metadata = New-Object System.Collections.ArrayList
            $lines | Foreach-Object {
                if ($_ -Match '#') {
                    $value = $_.substring(1)
                    $metadata.Add($value) > $null
                }
                $metadata |  Foreach-Object {
                    Write-Host "$($_)"
                }
                if ($_ -NotMatch '#') {
                    $parts = $_ -split ";"

                    $partCounter = 1
                    $Found = $groupsSheet.Cells.Find($parts[0])
                    $worksheet.Cells.Item($currentRow, $partCounter) = $groupsSheet.Cells.Item($Found.Row, $Found.Column + 1)

                    $partCounter = 2
                    $parts |  Foreach-Object {
                        $worksheet.Cells.Item($currentRow, $partCounter) = $_
                        $partCounter++
                    }
                    $partCounter = 15
                    $metadata |  Foreach-Object {
                        $m = $_ -split ";"
                        $worksheet.Cells.Item($currentRow, $partCounter) = $m[0]
                        $partCounter++
                        $worksheet.Cells.Item($currentRow, $partCounter) = $m[1]
                        $partCounter++

                    }
                    $currentRow++
                }

            }
        }

#saving & closing the file
#adjusting the column width so all data's properly visible
$usedRange = $worksheet.UsedRange
$usedRange.EntireColumn.AutoFit() | Out-Null
$excel.DisplayAlerts = $true
$workbook.SaveAs($outputpath, 51, [Type]::Missing, [Type]::Missing, $false, $false, 1, 2)

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

Remove-Variable -Name excel

