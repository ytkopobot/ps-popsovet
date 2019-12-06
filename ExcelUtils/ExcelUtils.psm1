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
Export-ModuleMember -Function 'GetSheet'
