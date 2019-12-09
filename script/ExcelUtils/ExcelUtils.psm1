$months = @("������", "�������", "����", "������", "���", "����", "����", "������", "��������", "�������", "������", "�������")
$letters = @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

function GetSheet($workbook, $sheetName) {
    $worksheet = $workbook.Worksheets | Where-Object {
        $_.name -eq $sheetName
    }
    if (-Not$worksheet) {
        Write-Host "�� ������ ���� $sheetName"
        return $null
    }
    return $worksheet
}

function GetMonthName($month) {
    if (-Not$month) {
        Write-Host "����� �� ������"
        return $null
    }
    if (($month -lt 1) -or ($month -gt 12)) {
        Write-Host "�������� �����: $month"
        return $null
    }
    return  $months[$month - 1]
}


function GetColumn($index) {
    if (-Not$index) {
        Write-Host "�� ������ ������"
        return $null
    }
    return $letters[$index - 1]
}

function GetColumnRange($index) {
    $column = GetColumn($index)
    return "$( $column ):$( $column )"
}

function CheckFileOpen {
    param ([parameter(Mandatory = $true)][string]$Path)

    $oFile = New-Object System.IO.FileInfo $Path

    if ((Test-Path -Path $Path) -eq $false) {
        return $false
    }

    try {
        $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)

        if ($oStream) {
            $oStream.Close()
        }
        $false
    } catch {
        # file is locked by a process.
        Write-Host "���� $Path ��� ������. ��� ���������� ��������� ��������� ��� ��������� � �������� ���" -ForegroundColor Cyan
        return $true
    }
}
Export-ModuleMember -Function 'GetSheet', 'GetMonthName', 'GetColumn', 'GetColumnRange', "CheckFileOpen"
