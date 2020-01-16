#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Import-Module -Name "$scriptPath\ExcelData\ExcelData.psm1"
Import-Module -Name "$scriptPath\ExcelUtils\ExcelUtils.psm1"

Function Main() {
    $excelFilePath = "$scriptPath\..\$ExcelFilename"

    Write-Host "#" -ForegroundColor Yellow
    Write-Host "# ������������ ���������� ������ �� ����� '$IncomingLogSheetName' ����� $excelFilePath" -ForegroundColor Yellow
    Write-Host "# ���������� �� � ���� $GroupSheetName ��������������� ������" -ForegroundColor Yellow
    Write-Host "# ����� ��������� �� ���� �������, �������� �������� ���������� ����� �����, �.�. ������� ��������." -ForegroundColor Yellow
    Write-Host "#" -ForegroundColor Yellow

    if (-Not [System.IO.File]::Exists($excelFilePath)) {
        Write-Host "���� �� ������ $excelFilePath"
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

    Write-Host "������� ����� ������ �� ����� '$IncomingLogSheetName' � ������� ������� ���������" -ForegroundColor Green
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
            Write-Host "�� ������� ������ ��� $incomingName, ������ $i ������� $IncomingGroupCell" -ForegroundColor Magenta
            continue
        }

        $payment = $worksheet.Cells.Item($i, $IncomingPaymentCell).Value2
        if (-Not$payment) {
            Write-Host "�� ������� ����� ������� ��� $incomingName, ������ $i ������� $IncomingPaymentCell" -ForegroundColor Magenta
            continue
        }

        $groupSheet = $allGroupSheets["$groupNumber"]
        if (-Not$groupSheet) {
            $groupSheet = $allGroupSheets["$groupNumber"] = OpenGroup $excel $groupNumber
            if (-Not$groupSheet) {
                Write-Host "���������� �������� ����� ��� $incomingName ������ $groupNumber, ������ $i - ���� ������ �� ��������" -ForegroundColor Magenta
                continue
            }
        }
        $paymentDate = $worksheet.Cells.Item($i, $IncomingPaymentDateCell).Text
        if (-Not$paymentDate) {
            Write-Host "���� ������ �� �������� $incomingName, ������ $i ������� $IncomingPaymentDateCell" -ForegroundColor Magenta
            continue
        }
        $monthNumber = [datetime]::parseexact($paymentDate, 'dd.MM.yyyy', $null).Month
        $monthName = GetMonthName $monthNumber


        $FindedMonth = $groupSheet.Cells.Item($GroupMonthRow, 1).EntireRow.Find($monthName)
        if (-Not$FindedMonth) {
            Write-Host "��� ������ $groupNumber �� ����� $GroupSheetName � ������ $GroupMonthRow �� ������ ����� $monthName, ����� ��� $incomingName ����� ��������" -ForegroundColor Magenta
            continue
        }

        $FindedName = $groupSheet.Cells.Item(1, $NameColumn).EntireColumn.Find($incomingName)
        if (-Not$FindedName) {
            Write-Host "��� ������ $groupNumber �� ����� $GroupSheetName � ������� $NameColumn �� ������� ��� $incomingName, ����� ����� ��������" -ForegroundColor Magenta
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
            Write-Host "��� ������ $groupNumber ��� $incomingName � ������ $monthName ��� ���� ��������: $currentCommonFondValue � $currentGroupValue, ����� $payment ����� ��������!" -ForegroundColor Red
            Write-Host ""
            continue
        }

        $CommonFondSum = $groupsheet.Cells.Item($FindedName.Row, $CommonFondColumn).Value2
        $GroupFondSum = $groupsheet.Cells.Item($FindedName.Row, $GroupFondColumn).Value2


        if ($payment -eq ($CommonFondSum + $GroupFondSum)){
            Write-Host "��� $incomingName $groupNumber ��. � ������ '$monthName' ����� ����� ������������ ������, ������� ��� ����� ����� ������� �� ��� �������" -ForegroundColor Green
            $currentCommonFondValueCell.Interior.ColorIndex = 4
            $currentCommonFondValueCell.Value2 = $CommonFondSum

            $currentGroupValueCell.Interior.ColorIndex = 4
            if($GroupFondSum){
                $currentGroupValueCell.Value2 = $GroupFondSum
            }

        }else {
            Write-Host "��� $incomingName $groupNumber ��. � ������ '$monthName' ����� �� ����� ������������ ������, ������� ��� ����� �������� � ���� ���� ��� ���������� ������ ��������" -ForegroundColor Cyan
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
    Write-Host "�����:" -BackgroundColor Green
    Write-Host "���������� $incomingsCount �����, � $startRow �� $endRow" -ForegroundColor Green
    Write-Host "���������  $addedIncomings �������" -ForegroundColor Green
    Write-Host "�������  $( $allGroupSheets.count ) ������ �����" -ForegroundColor Green
    Write-Host "��������  $( $editedFilesCount.count ) ������ �����" -ForegroundColor Green

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
        Write-Host "���� �� ������ $groupFilePath"
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



