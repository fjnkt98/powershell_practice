Set-PSDebug -strict
Add-Type -AssemblyName Microsoft.VisualBasic

# Set working directory
$workingDirectory = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

# Launch Excel application
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $true
$excel.ScreenUpdating = $false

# Open target Excel workbook
$fileName = "practice5.xlsx"
try {
  $book = $excel.Workbooks.Open($workingDirectory + "\" + $fileName)

  $source_sheet = $book.Sheets("data")
  [int]$length = 3
  for ($i = 0; $i -lt 3; $i++) {
    $array = [System.Collections.Generic.List[int]]::new()
    for ($j = 1; $j -le 10; $j += $length) {
      $range = $source_sheet.Range($source_sheet.Cells($i + 1, $j), $source_sheet.Cells($i + 1, $j + $length - 1))
      $array.Add($excel.WorksheetFunction.Max($range))
    }

    foreach($item in $array) {
      Write-Host $item
    }
  }

  $book.Save()
} finally {
  $excel.Quit()
}

Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable
[GC]::Collect()
