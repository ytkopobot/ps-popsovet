$ExcelFilename = "2019.xlsx"
$OutcomingFolder = "outcoming"
$IncomingFolder = "incoming"

$IncomingLogSheetName = "Взносы"

$CommonListSheetName = "Общий список"
$GroupNumberCell = 1
$NameCell = 3
$GroupTitleCell = 3
$PaymentCell = 4

$OutcomingFilename = "N группа 19-20.txt"

$GroupCount = 12


$SGParts = New-Object -Type PSObject -Property @{
    'childName'   = 0
    'date' = 8
    'paymentId' = 10
}


$xlCellTypeLastCell = 11
Export-ModuleMember -Variable ExcelFilename, OutcomingFolder, CommonListSheetName, IncomingFolder, IncomingLogSheetName, GroupNumberCell, xlCellTypeLastCell, NameCell, PaymentCell, OutcomingFilename, GroupCount, GroupTitleCell, SGParts
