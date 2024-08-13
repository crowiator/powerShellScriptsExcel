# Define the path to the Excel file
$excelFilePath = "C:\Users\User\Desktop\sca_checkmarks.xlsx"

# Specify the columns you want to keep
$columnsToKeep = @("Name", "Version", "CriticalVulnerabilityCount", "HighVulnerabilityCount", "MediumVulnerabilityCount", "LowVulnerabilityCount", "NoneVulnerabilityCount")

# Load the Excel file
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelFilePath)
$worksheet = $workbook.Sheets.Item(1)

# Get the used range of the worksheet
$usedRange = $worksheet.UsedRange
$headerRow = 1

# Get the column headers
$headers = @()
for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
    $headers += $worksheet.Cells.Item($headerRow, $col).Value()
}

# Find the columns that are not in the keep list and hide them
for ($col = $usedRange.Columns.Count; $col -ge 1; $col--) {
    if (-not ($columnsToKeep -contains $headers[$col - 1])) {
        $worksheet.Columns.Item($col).Delete()
    }
}

# Save the modified workbook
$workbook.SaveAs("C:\Users\User\Desktop\sca_checkmarks1.xlsx")

# Cleanup
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable -Name excel