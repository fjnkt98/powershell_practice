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

    $xRange = $sheet.Range("A$topRow", "A$bottomRow")
    $dataRange1 = $sheet.Range("B$topRow", "B$bottomRow")
    $dataRange2 = $sheet.Range("C$topRow", "C$bottomRow")

    # Delete existing charts
    $chartObjects = $sheet.ChartObjects()
    if ($chartObjects.Count -gt 0) {
      $chartObjects.Delete()
    }

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

  $book.Save()
} finally {
  $excel.Quit()
}

Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable
[GC]::Collect()
