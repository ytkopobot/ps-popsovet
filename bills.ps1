#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Import-Module -Name "$scriptPath\ExcelData\ExcelData.psm1"
Import-Module -Name "$scriptPath\ExcelUtils\ExcelUtils.psm1"


#
# ���������� ���������� � ������� ����� � ���� ��������� ������
#
Function Main() {
    Write-Host "#" -ForegroundColor Yellow
    Write-Host "# ������ ��������� ����� ��� �������� � ������� ����� " -ForegroundColor Yellow
    Write-Host "#" -ForegroundColor Yellow
    Write-Host "����� ��� ���������� [1-12]" -ForegroundColor Green
    [uint16] $month = Read-Host
    $monthName = GetMonthName $month
    Write-Host "$monthName"

    if (-Not$monthName) {
        exit
    }

    Write-Host "����� ������ [1-12]" -ForegroundColor Green
    $groupNumber = Read-Host

    $excelFilePath = "$scriptPath\$ExcelFilename"
    if (-Not [System.IO.File]::Exists($excelFilePath)) {
        Write-Host "���� $ExcelFilename �� ������ $excelFilePath" -ForegroundColor Red
        exit
    }

    $groupFilePath = "$scriptPath\$($GroupExcel.Replace("N", $groupNumber) )"
    if (-Not [System.IO.File]::Exists($groupFilePath)) {
        Write-Host "���� ������ �� ������ $groupFilePath" -ForegroundColor Red
        exit
    }

    $outcomingDir = "$scriptPath\$OutcomingFolder"

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

    $newFile = "$outcomingDir\$($OutcomingFilename.Replace("N", $groupNumber).Replace("M", $monthName) )"
    if (-Not [System.IO.File]::Exists($newFile)) {
        New-Item $newFile | Out-Null
    }
    Add-Content $newFile "#TYPE 7"
    Add-Content $newFile "#SERVICE 40334"

    $overalSum = 0
    $rows = 0
    $skips = 0

    for ($i = $startRow; $i -le $endRow; $i++)
    {
        $CommonFondSum = $groupsheet.Cells.Item($i, $CommonFondColumn).Value2
        $GroupFondSum = $groupsheet.Cells.Item($i, $GroupFondColumn).Value2
        $currentDebt = $groupsheet.Cells.Item($i, $DebtColumn).Value2
        $Name = $groupsheet.Cells.Item($i, $NameColumn).Text
        if (-Not$Name) {
            continue;
        }
        if ( $Name.StartsWith("#")) {
            continue
        }
        $Adress = FindAdress $Name $commonListSheet
        if (-Not$Adress) {
            $skips++
            continue
        }
        $Contract = FindContract $Name $commonListSheet
        if (-Not$Contract) {
            $skips++
            continue
        }

        $ForDebt = $groupsheet.Cells.Item($i, 1).Text # �� ������� TODO ���������� ���� ���� �����

        $currentSum = 0
        if ($ForDebt -ieq "�") {
            # ������� � ������������
            if ((-Not$currentDebt) -or ($currentDebt -eq 0)) {
                Write-Host "����� ���, ���������� ����� ��������� ��� $Name"
                $skips++
                continue;
            }
            if ($currentDebt -lt 0) {
                Write-Host "���� �������������, ���������� ����� ��������� ��� $Name"
                $skips++
                continue;
            }
            $currentSum = $currentDebt
        } else {
            $currentSum = $CommonFondSum + $GroupFondSum
            if ($currentSum -le 0) {
                Write-Host "����� ������ �� �����������, ���������� ����� ��������� ��� $Name"
                $skips++
                continue;
            }
        }
        Add-Content $newFile "$Name;$Adress;$Contract;$currentSum"
        $rows++
        $overalSum = $overalSum + $currentSum
    }

    Add-Content $newFile "#FILESUM $overalSum"

    Write-Host "�����:" -ForegroundColor Green
    Write-Host "���������  $rows �����" -ForegroundColor Green
    Write-Host "���������  $skips �����" -ForegroundColor Green
    Write-Host "����� �����  $overalSum" -ForegroundColor Green
    Write-Host "��������� $newFile" -ForegroundColor Green

    #saving & closing the file
    #adjusting the column width so all data's properly visible
    $excel.DisplayAlerts = $false
    $excel.Quit()

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($groupsheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($commonListSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Remove-Variable -Name excel

    Write-Host "��� ���������� ������� Enter" -ForegroundColor Blue
    Read-Host
}

Function FindAdress($Name, $commonListSheet) {
    $range = GetColumnRange($NameCell)
    $FindedCell = $commonListSheet.Range($range).EntireColumn.Find($Name)
    $Address = $commonListSheet.Cells.Item($FindedCell.Row, $AddressCell).Value2
    if (-Not$Address) {
        Write-Host "����� �� ������ ��� $Name" -ForegroundColor Magenta
        return $null
    }
    return $Address
}

Function FindContract($Name, $commonListSheet) {
    $range = GetColumnRange($NameCell)
    $FindedCell = $commonListSheet.Range($range).EntireColumn.Find($Name)
    $Contract = $commonListSheet.Cells.Item($FindedCell.Row, $ContractCell).Value2
    if (-Not$Contract) {
        Write-Host "����� �������� �� ������ ��� $Name"
        return $null
    }
    return $Contract
}
Main



