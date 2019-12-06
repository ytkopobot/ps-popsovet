function GetSheet ($workbook, $sheetName){
    $worksheet = $workbook.Worksheets | Where-Object {
        $_.name -eq $sheetName
    }
    if(-Not $worksheet){
        Write-Host "�� ������ ���� $sheetName"
        return $null
    }
    return $worksheet
}

function GetMonthName ($month){
    if (-Not $month){
        Write-Host "����� �� ������"
        return $null
    }
    if(($month -lt 1) -or ($month -gt 12)){
        Write-Host "�������� �����: $month"
        return $null
    }
    return  @("������", "�������", "����", "������", "���", "����", "����", "������", "��������", "�������", "������", "�������")[$month-1]
}

Export-ModuleMember -Function 'GetSheet', 'GetMonthName'
