#$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Specify the path to the Excel file and the WorkSheet Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
Import-Module -Name "$scriptPath\ExcelData\ExcelData.psm1"
Import-Module -Name "$scriptPath\ExcelUtils\ExcelUtils.psm1"

Function Main() {
    Write-Host "#" -ForegroundColor Yellow
    Write-Host "# Копирование файлов групп для публичного доступа " -ForegroundColor Yellow
    Write-Host "#" -ForegroundColor Yellow
    Write-Host "Введите номер группы, для которой нужно скопировать файл (1-12, 0 - если нужно скопировать все):" -ForegroundColor Green
    [uint16] $groupNumber = Read-Host
    $dateSuffix = Get-Date -Format "yyyy-MM-dd"
    Write-Host ""
    if ($groupNumber -eq 0){
        for ($i = 1; $i -le 12; $i++){
            CopyFile $i $dateSuffix
        }
    }else{
        CopyFile $groupNumber $dateSuffix
    }

    Write-Host ""
    Pause
}

Function CopyFile($groupNumber, $dateSuffix) {
    $groupFilePath = "$scriptPath\..\$($GroupExcel.Replace("N", $groupNumber) )"
    $groupFilePathForPublic = "$scriptPath\..\..\Public\$groupNumber\$($GroupExcelForPublic.Replace("N", $groupNumber).Replace("D", $dateSuffix) )"

    if (-Not [System.IO.File]::Exists($groupFilePath)) {
        Write-Host "Файл группы не найден $groupFilePath" -ForegroundColor Red
        exit
    }

    Copy-Item -Path $groupFilePath -Destination $groupFilePathForPublic
    Write-Host "Скопировали $($groupFilePath | Resolve-Path)    в    $($groupFilePathForPublic | Resolve-Path)" -ForegroundColor Cyan

    $excel = New-Object -ComObject excel.application
    $excel.visible = $true
    $excel.DisplayAlerts = $true
    $groupbook = $excel.Workbooks.Open($groupFilePathForPublic)

    $groupsheet = GetSheet $groupbook $GroupSheetName
    if (-Not$groupsheet) {
        exit
    }
    $groupSheet.Columns($CalcCommonDebtColumn).Hidden = $true
    $groupSheet.Columns($CalcGroupDebtColumn).Hidden = $true

    $excel.DisplayAlerts = $false
    $groupbook.Close($true)
    $excel.Quit()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($groupsheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable -Name excel
}

Main
