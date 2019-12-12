#Файл общего списка с нужной нам информацией
$ExcelFilename = "2019.xlsx"
$IncomingLogSheetName = "Взносы"
$IncomingGroupCell = 2
$IncomingPaymentDateCell = 3  # Дата когда пришел взнос
$IncomingNameCell = 4
$IncomingPaymentCell = 7
$PaymentIdCell = 9

$CommonListSheetName = "Общий список"
$GroupNumberCell = 2
$NameCell = 4
$ContractCell = 5
$AddressCell = 6

#Файл с группами содержит информацию для групп
$GroupExcel = "N группа.xlsx"
$GroupSheetName = "Взносы"
$GroupStartRow = 6 # c какой строчки начинаются фамилии
$TagColumn = 1   # тэг для строчки
$NameColumn = 2   # Сумма взноса в фонд сада
$CommonFondColumn = 4   # Сумма взноса в фонд сада
$GroupFondColumn = 5  # Сумма взноса в фонд группы
$DebtColumn = 21  # Сумма долга
$GroupMonthRow = 4


#Формат от Системы Город (индексы полей)
$SGParts = New-Object -Type PSObject -Property @{
    'childName'   = 0
    'date' = 8
    'paymentId' = 10
}

$OutcomingFolder = "outcoming"
$IncomingFolder = "incoming"
$OutcomingFilename = "N группа M 19-20.txt"


$xlCellTypeLastCell = 11
#TODO некоторые имена очень похожи!!!
Export-ModuleMember -Variable ExcelFilename, OutcomingFolder, CommonListSheetName, `
    IncomingFolder, PaymentIdCell, IncomingNameCell, IncomingGroupCell, IncomingLogSheetName, IncomingPaymentDateCell, IncomingPaymentCell, `
    GroupNumberCell, xlCellTypeLastCell, TagColumn, NameCell, ContractCell, AddressCell, `
    OutcomingFilename, GroupTitleCell, SGParts,`
    GroupExcel, GroupStartRow, CommonFondColumn, GroupFondColumn, GroupSheetName, NameColumn, DebtColumn, GroupMonthRow
