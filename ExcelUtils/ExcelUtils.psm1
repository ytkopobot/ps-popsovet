function GetSheet ($workbook, $sheetName){
    $worksheet = $workbook.Worksheets | Where-Object {
        $_.name -eq $sheetName
    }
    if(-Not $worksheet){
        Write-Host "Ќе найден лист $sheetName"
        return $null
    }
    return $worksheet
}

function GetMonthName ($month){
    if (-Not $month){
        Write-Host "ћес€ц не указан"
        return $null
    }
    if(($month -lt 1) -or ($month -gt 12)){
        Write-Host "Ќеверный мес€ц: $month"
        return $null
    }
    return  @("€нварь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сент€брь", "окт€брь", "но€брь", "декабрь")[$month-1]
}

Export-ModuleMember -Function 'GetSheet', 'GetMonthName'
