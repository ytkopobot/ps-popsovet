#Файл общего списка с нужной нам информацией
$ExcelFilename = "2019.xlsx"
$IncomingLogSheetName = "Взносы"
$IncomingGroupCell = 1
$IncomingPaymentDateCell = 2  #Дата когда пришел взнос
$IncomingNameCell = 3
$PaymentIdCell = 13
$IncomingPaymentCell = 6

$CommonListSheetName = "Общий список"
$GroupNumberCell = 1
$NameCell = 3
$ContractCell = 4
$AddressCell = 5

#Файл с группами содержит информацию для групп
$GroupExcel = "N группа.xlsx"
$GroupSheetName = "Взносы"
$GroupStartRow = 6 # c какой строчки начинаются фамилии
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
    GroupNumberCell, xlCellTypeLastCell, NameCell, ContractCell, AddressCell, `
    OutcomingFilename, GroupTitleCell, SGParts,`
    GroupExcel, GroupStartRow, CommonFondColumn, GroupFondColumn, GroupSheetName, NameColumn, DebtColumn, GroupMonthRow
