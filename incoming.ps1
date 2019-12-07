#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Import-Module -Name "$scriptPath\ExcelData\ExcelData.psm1"
Import-Module -Name "$scriptPath\ExcelUtils\ExcelUtils.psm1"

Function Main() {
    $excelFilePath = "$scriptPath\$ExcelFilename"
    $incomingDir = "$scriptPath\$IncomingFolder"

    Write-Host "#" -ForegroundColor Yellow
    Write-Host "# ������ ��������� ����� �� ������� ����� �� ����� $incomingDir" -ForegroundColor Yellow
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
    Get-ChildItem $incomingDir -Filter *.txt |
            Foreach-Object {
                # ��������� �����
                $filesCount++
                $lines = Get-Content $_.FullName
                $metadata = New-Object System.Collections.ArrayList

                # ��������� ����� � �����
                $lines | Foreach-Object {
                    if ($_ -Match '#') {
                        $value = $_.substring(1)
                        $metadata.Add($value) > $null
                    }

                    if ($_ -NotMatch '#') {
                        $errorMessage = WriteLine $_ $worksheet $groupsSheet $currentRow
                        if ($errorMessage) {
                            Write-Host $errorMessage -ForegroundColor Magenta
                        } else {
                            Write-Host "��������� ������ $currentRow" -ForegroundColor Green
                            $currentRow++
                            $paymentAddedCount++
                        }
                        $paymentCount++
                    }
                }
            }

    Write-Host "�����:" -ForegroundColor Green
    Write-Host "���������� $filesCount ������, $paymentCount �����" -ForegroundColor Green
    Write-Host "���������  $paymentAddedCount �����" -ForegroundColor Green
    Write-Host "���� '$IncomingLogSheetName' ����� $excelFilePath" -ForegroundColor Magenta

    #saving & closing the file
    #adjusting the column width so all data's properly visible
    $usedRange = $worksheet.UsedRange
    $usedRange.EntireColumn.AutoFit() | Out-Null
    $excel.DisplayAlerts = $false
    # $workbook.SaveAs($outputpath, 51, [Type]::Missing, [Type]::Missing, $false, $false, 1, 2)

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($groupsSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Remove-Variable -Name excel

    Write-Host "��� ���������� ������� Enter" -ForegroundColor Blue
    Read-Host
}

Function WriteLine($line, $worksheet, $groupsSheet, $currentRow) {
    $parts = $line -split ";"

    $partCounter = 1

    $childName = $parts[$SGParts."childName"]
    $paymentId = $parts[$SGParts."paymentId"]
    $date = $parts[$SGParts."date"]

    $range = GetColumnRange($PaymentIdCell)
    $FindedCell = $worksheet.Range($range).EntireColumn.Find($paymentId)

    If ($FindedCell) {
        return "$paymentId - ��� ����������, ������ $( $FindedCell.Row ), $childName";
    }
    $Found = $groupsSheet.Cells.Find($childName)
    $groupNumber = "?"
    if ($Found) {
        $groupNumber = $groupsSheet.Cells.Item($Found.Row, $GroupNumberCell).Text
    } else {
        return "$paymentId - �� ������� ������ $childName";
    }

    $worksheet.Cells.Item($currentRow, $partCounter) = $groupNumber

    $partCounter = 2
    $worksheet.Cells.Item($currentRow, $partCounter) = [datetime]::parseexact($date, 'dd/MM/yyyy', $null).ToString('dd.MM.yyyy')

    $partCounter = 3 # ����� ��� �������� �������
    $parts |  Foreach-Object {
        $worksheet.Cells.Item($currentRow, $partCounter) = $_
        $partCounter++
    }
    $partCounter++ # � ���� ������� �������� ��������� ����������
    $metadata |  Foreach-Object {
        $m = $_ -split ";"
        $worksheet.Cells.Item($currentRow, $partCounter) = $m[0]
        $partCounter++
        $worksheet.Cells.Item($currentRow, $partCounter) = $m[1]
        $partCounter++

    }
}

Main



