#���� ������ ������ � ������ ��� �����������
$ExcelFilename = "2019.xlsx"

$IncomingLogSheetName = "������"

$CommonListSheetName = "����� ������"
$GroupNumberCell = 1
$NameCell = 3
$GroupTitleCell = 3
$PaymentCell = 4

#���� � �������� �������� ���������� ��� �����
$GroupExcel = "N ������.xlsx"
$GroupSheetName = "������"
$GroupStartRow = 6 # c ����� ������� ���������� �������
$NameColumn = 2   # ����� ������ � ���� ����
$CommonFondColumn = 4   # ����� ������ � ���� ����
$GroupFondColumn = 5  # ����� ������ � ���� ������



$SGParts = New-Object -Type PSObject -Property @{
    'childName'   = 0
    'date' = 8
    'paymentId' = 10
}

$OutcomingFolder = "outcoming"
$IncomingFolder = "incoming"
$OutcomingFilename = "N ������ M 19-20.txt"
$GroupCount = 12 # TODO ���������, �������� ��� �� �����


$xlCellTypeLastCell = 11
#TODO ��������� ����� ����� ������!!!
Export-ModuleMember -Variable ExcelFilename, OutcomingFolder, CommonListSheetName, IncomingFolder, IncomingLogSheetName,`
    GroupNumberCell, xlCellTypeLastCell, NameCell, PaymentCell, OutcomingFilename, GroupCount, GroupTitleCell, SGParts,`
    GroupExcel, GroupStartRow, CommonFondColumn, GroupFondColumn, GroupSheetName, NameColumn
