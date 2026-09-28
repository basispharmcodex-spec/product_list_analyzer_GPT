param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$SheetName = "",
    [int]$StartRow = 1,
    [int]$MaxRows = 50,
    [int]$MaxColumns = 20
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $true)

    $sheetNames = @()
    foreach ($sheet in $workbook.Worksheets) {
        $sheetNames += [string]$sheet.Name
    }

    if ([string]::IsNullOrWhiteSpace($SheetName)) {
        $worksheet = $workbook.Worksheets.Item(1)
    }
    else {
        $worksheet = $workbook.Worksheets.Item($SheetName)
    }

    $lastCell = $worksheet.Cells.Find('*', $worksheet.Cells.Item(1, 1), -4123, 2, 1, 2, $false, $false, $false)
    if ($null -eq $lastCell) {
        $lastRow = 0
        $lastColumn = 0
    }
    else {
        $lastRow = [int]$lastCell.Row
        $lastColumnCell = $worksheet.Cells.Find('*', $worksheet.Cells.Item(1, 1), -4123, 2, 2, 2, $false, $false, $false)
        $lastColumn = [int]$lastColumnCell.Column
    }

    $rows = @()
    $endRow = [Math]::Min($lastRow, $StartRow + $MaxRows - 1)
    $endColumn = [Math]::Min($lastColumn, $MaxColumns)
    for ($rowNumber = $StartRow; $rowNumber -le $endRow; $rowNumber++) {
        $values = @()
        for ($columnNumber = 1; $columnNumber -le $endColumn; $columnNumber++) {
            $value = [string]$worksheet.Cells.Item($rowNumber, $columnNumber).Text
            $values += ($value -replace "[\r\n\t]+", " ")
        }
        $rows += [ordered]@{
            row = $rowNumber
            values = $values
        }
    }

    $result = [ordered]@{
        file_path = [System.IO.Path]::GetFullPath($Path)
        format = "xlsb"
        reader = "Microsoft Excel COM"
        sheets = $sheetNames
        selected_sheet = [string]$worksheet.Name
        last_row = $lastRow
        last_column = $lastColumn
        rows = $rows
    }

    $result | ConvertTo-Json -Compress -Depth 8
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
