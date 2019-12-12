#���� ������ ������ � ������ ��� �����������
$ExcelFilename = "2019.xlsx"
$IncomingLogSheetName = "������"
$IncomingGroupCell = 2
$IncomingPaymentDateCell = 3  # ���� ����� ������ �����
$IncomingNameCell = 4
$IncomingPaymentCell = 7
$PaymentIdCell = 9

$CommonListSheetName = "����� ������"
$GroupNumberCell = 2
$NameCell = 4
$ContractCell = 5
$AddressCell = 6

#���� � �������� �������� ���������� ��� �����
$GroupExcel = "N ������.xlsx"
$GroupSheetName = "������"
$GroupStartRow = 6 # c ����� ������� ���������� �������
$TagColumn = 1   # ��� ��� �������
$NameColumn = 2   # ����� ������ � ���� ����
$CommonFondColumn = 4   # ����� ������ � ���� ����
$GroupFondColumn = 5  # ����� ������ � ���� ������
$DebtColumn = 21  # ����� �����
$GroupMonthRow = 4


#������ �� ������� ����� (������� �����)
$SGParts = New-Object -Type PSObject -Property @{
    'childName'   = 0
    'date' = 8
    'paymentId' = 10
}

$OutcomingFolder = "outcoming"
$IncomingFolder = "incoming"
$OutcomingFilename = "N ������ M 19-20.txt"


$xlCellTypeLastCell = 11
#TODO ��������� ����� ����� ������!!!
Export-ModuleMember -Variable ExcelFilename, OutcomingFolder, CommonListSheetName, `
    IncomingFolder, PaymentIdCell, IncomingNameCell, IncomingGroupCell, IncomingLogSheetName, IncomingPaymentDateCell, IncomingPaymentCell, `
    GroupNumberCell, xlCellTypeLastCell, TagColumn, NameCell, ContractCell, AddressCell, `
    OutcomingFilename, GroupTitleCell, SGParts,`
    GroupExcel, GroupStartRow, CommonFondColumn, GroupFondColumn, GroupSheetName, NameColumn, DebtColumn, GroupMonthRow
