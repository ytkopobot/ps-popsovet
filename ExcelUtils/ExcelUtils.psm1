$months = @("������", "�������", "����", "������", "���", "����", "����", "������", "��������", "�������", "������", "�������")
$letters = @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

function GetSheet($workbook, $sheetName) {
    $worksheet = $workbook.Worksheets | Where-Object {
        $_.name -eq $sheetName
    }
    if (-Not$worksheet) {
        Write-Host "�� ������ ���� $sheetName"
        return $null
    }
    return $worksheet
}

function GetMonthName($month) {
    if (-Not$month) {
        Write-Host "����� �� ������"
        return $null
    }
    if (($month -lt 1) -or ($month -gt 12)) {
        Write-Host "�������� �����: $month"
        return $null
    }
    return  $months[$month - 1]
}


function GetColumn($index) {
    if (-Not$index) {
        Write-Host "�� ������ ������"
        return $null
    }
    return $letters[$index-1]
}

function GetColumnRange($index){
    $column =  GetColumn($index)
    return "$($column):$($column)"
}
Export-ModuleMember -Function 'GetSheet', 'GetMonthName', 'GetColumn', 'GetColumnRange'
