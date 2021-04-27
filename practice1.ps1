$working_directory = (Convert-Path .)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $true

$book = $excel.Workbooks.Add()

$sheet = $book.ActiveSheet

$sheet.Name = 'hoge'

$sheet.Range("B2") = "M200"

$sheet.Range("A1", "A10") = 10
$sheet.Range("A3", "B3") = 5,10

$book.SaveAs("$working_directory\practice1.xlsx")
$excel.Quit()
$excel = $null

[GC]::Collect()