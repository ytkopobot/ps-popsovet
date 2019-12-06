#Файл общего списка с нужной нам информацией
$ExcelFilename = "2019.xlsx"

$IncomingLogSheetName = "Взносы"

$CommonListSheetName = "Общий список"
$GroupNumberCell = 1
$NameCell = 3
$NameCell2 = "C"
$ContractCell = 5
$AddressCell = 6

#Файл с группами содержит информацию для групп
$GroupExcel = "N группа.xlsx"
$GroupSheetName = "Взносы"
$GroupStartRow = 6 # c какой строчки начинаются фамилии
$NameColumn = 2   # Сумма взноса в фонд сада
$CommonFondColumn = 4   # Сумма взноса в фонд сада
$GroupFondColumn = 5  # Сумма взноса в фонд группы



$SGParts = New-Object -Type PSObject -Property @{
    'childName'   = 0
    'date' = 8
    'paymentId' = 10
}

$OutcomingFolder = "outcoming"
$IncomingFolder = "incoming"
$OutcomingFilename = "N группа M 19-20.txt"
$GroupCount = 12 # TODO проверить, возможно уже не нужна


$xlCellTypeLastCell = 11
#TODO некоторые имена очень похожи!!!
Export-ModuleMember -Variable ExcelFilename, OutcomingFolder, CommonListSheetName, IncomingFolder, IncomingLogSheetName,`
    GroupNumberCell, xlCellTypeLastCell, NameCell, NameCell2, PaymentCell, ContractCell, AddressCell, `
    OutcomingFilename, GroupCount, GroupTitleCell, SGParts,`
    GroupExcel, GroupStartRow, CommonFondColumn, GroupFondColumn, GroupSheetName, NameColumn
