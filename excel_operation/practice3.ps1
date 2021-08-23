# Set working directory
$working_directory = (Convert-Path .)

# Launch Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $true

# Get excel workbook
$book = $excel.Workbooks.Add()

# Get sheet
$sheet = $book.ActiveSheet
$sheet.Name = 'Sheet1'

$sheet.Cells.Item(1, 1) = "Data1"
$sheet.Cells.Item(1, 2) = "Data2"

# Import csv files
$data1 = Import-Csv .\practice8_a.csv -Encoding UTF8
$data2 = Import-Csv .\practice8_b.csv -Encoding UTF8

[int]$data_length = $data1.Count

# Repost each data
for ($i = 0; $i -lt $data_length; $i++) {
  $sheet.Cells.Item($i+2, 1) = $data1[$i].data
  $sheet.Cells.Item($i+2, 2) = $data2[$i].data
}

# Create chart
$table = $sheet.Range("A1", "B$data_length")  # Set data source table
$chart = $sheet.ChartObjects().Add(200, 10, 600, 400).Chart # Create chart object
$chart.SetSourceData($table) | Out-Null # Set data
$chart.ChartType = 4  # Line graph
$chart.HasTitle = $true # Add graph title
$chart.ChartTitle.Text = 'Powered by PowerShell'  # Set graph title

# Save book
$book.SaveAs("$working_directory\practice8.xlsx")

# Terminate excel
$excel.Quit()
$excel = $null
$book = $null
$sheet = $null
$table = $null
$chart = $null

[GC]::Collect()