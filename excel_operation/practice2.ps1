# Set working directory
$working_directory = (Convert-Path .)

# Launch excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $true

# Get excel workbook
$book = $excel.Workbooks.Add()

# Get sheet
$sheet = $book.ActiveSheet
$sheet.Name = 'hoge'

# Some processes...
for ($i = 1; $i -le 10; $i++) {
  $sheet.Cells.Item($i, 1).Value2 = ($i)
  $sheet.Cells.Item($i, 2).Value2 = ($i * $i)
}

$table = $sheet.Range("A1", "B10")
$table.Columns.AutoFit() | Out-Null
$sheet.Range("A1", "B10").Font.Bold = $true
$sheet.Range("A1", "C1").Interior.ColorIndex = 15

$position_x = 200
$position_y = 10
$width = 600
$height = 400

$chart = $sheet.ChartObjects().Add($position_x, $position_y, $width, $height).Chart
$chart.SetSourceData($table) | Out-Null

$chart.ChartType = -4169

$chart.HasTitle = $true
$chart.ChartTitle.Text = 'Powered by PowerShell'

# Save book
$book.SaveAs("$working_directory\practice2.xlsx")

# Terminate excel
$excel.Quit()
$excel = $null
$book = $null
$sheet = $null
$table = $null
$chart = $null

[GC]::Collect()