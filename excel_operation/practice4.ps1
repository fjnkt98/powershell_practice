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
$fileName = "practice4.xlsx"
try {
  $book = $excel.Workbooks.Open($workingDirectory + "\" + $fileName)

  [string[]]$targetSheets = @("A", "B", "C", "D")

  foreach ($sheetName in $targetSheets) {
    $sheet = $book.Sheets($sheetName)
    $sheet.Range("A1").Clear() | Out-Null

    [int]$chartPositionX = 20
    [int]$chartPositionY = 20
    [int]$chartWidth = 370
    [int]$chartHeight = 200

    [int]$topRow = 17
    [int]$bottomRow = $sheet.Range("A$topRow").End([Microsoft.Office.Interop.Excel.XlDirection]::xlDown).Row

    $xRange = $sheet.Range($sheet.Cells($topRow, 1), $sheet.Cells($bottomRow, 1))
    $dataRange1 = $sheet.Range($sheet.Cells($topRow, 2), $sheet.Cells($bottomRow, 2))
    $dataRange2 = $sheet.Range($sheet.Cells($topRow, 3), $sheet.Cells($bottomRow, 3))

    # Delete existing charts
    $sheet.ChartObjects().Delete()

    # Create chart object
    $chart = $sheet.ChartObjects().Add(
      $chartPositionX, $chartPositionY,
      $chartWidth, $chartHeight
    ).Chart

    # Configure chart title
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = "Chart " + $sheet.Name
    $chart.ChartTitle.Font.Bold = $false

    # Configure chart legend
    $chart.HasLegend = $true
    $chart.Legend.Position = [Microsoft.Office.Interop.Excel.XlLegendPosition]::xlLegendPositionBottom

    # Add data series
    $dataSeries1 = $chart.SeriesCollection().NewSeries()
    $dataSeries1.ChartType = [Microsoft.Office.Interop.Excel.XlChartType]::xlColumnClustered
    $dataSeries1.AxisGroup = [Microsoft.Office.Interop.Excel.XlAxisGroup]::xlPrimary
    $dataSeries1.XValues = $xRange
    $dataSeries1.Values = $dataRange1
    $dataSeries1.Format.Fill.ForeColor.RGB = [Microsoft.VisualBasic.Information]::RGB(0, 176, 240)
    $dataSeries1.Format.Line.ForeColor.RGB = [Microsoft.VisualBasic.Information]::RGB(0, 0, 0)
    $dataSeries1.Format.Line.Visible = $true
    $dataSeries1.Format.Line.Transparency = 0

    $dataSeries2 = $chart.SeriesCollection().NewSeries()
    $dataSeries2.ChartType = [Microsoft.Office.Interop.Excel.XlChartType]::xlLine
    $dataSeries2.AxisGroup = [Microsoft.Office.Interop.Excel.XlAxisGroup]::xlSecondary
    $dataSeries2.XValues = $xRange
    $dataSeries2.Values = $dataRange2
    $dataSeries2.Format.Fill.ForeColor.RGB = [Microsoft.VisualBasic.Information]::RGB(0, 176, 240)
    $dataSeries2.Format.Line.ForeColor.RGB = [Microsoft.VisualBasic.Information]::RGB(0, 0, 0)
    $dataSeries2.Format.Line.Visible = $true
    $dataSeries2.Format.Line.Transparency = 0

    $chart.ChartGroups([Microsoft.Office.Interop.Excel.XlAxisGroup]::xlPrimary).GapWidth = 0
    $chart.ChartGroups([Microsoft.Office.Interop.Excel.XlAxisGroup]::xlPrimary).Overlap = 0
  }

  $sheet = $null
  $chart = $null
  $xRange = $null
  $dataRange1 = $null
  $dataRange2 = $null
  $dataSeries1 = $null
  $dataSeries2 = $null

  $book.Save()
} catch [System.Runtime.InteropServices.COMException] {
  Write-Error "The target file does'nt exit."
} finally {
  $excel.Quit()

  $excel = $null
  $book = $null
}

[GC]::Collect()
