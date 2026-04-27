$ErrorActionPreference = 'Stop'

$root      = 'c:\Users\user\My Drive\KAIST MASc 2021\Laboratory Work\Protocol\overpotential calculation\For paper SI'
$srcWbPath = Join-Path $root 'LSV_comparison.xlsx'
$outFile   = Join-Path $root 'LSV_reference_style.xlsx'
$pngDir    = Join-Path $root 'LSV_reference_style_png'
if (-not (Test-Path $pngDir)) { [void](New-Item -ItemType Directory -Path $pngDir) }

# Catalyst -> sample type chosen for the reference figure.
$catalystChoice = [ordered]@{
  'CN' = @{ Sample='BM';     LineStyle='solid'  }   # green solid
  'KB' = @{ Sample='BM';     LineStyle='solid'  }   # light blue solid
  'VC' = @{ Sample='Polyol'; LineStyle='dotted' }   # black dotted
  'AB' = @{ Sample='Polyol'; LineStyle='dotted' }   # orange dotted
}

# Excel-COM RGB integers (BGR-encoded). Pre-computed from reference colors.
function Rgb([int]$r,[int]$g,[int]$b){ return ($b -shl 16) -bor ($g -shl 8) -bor $r }
$catalystColor = @{
  'CN' = Rgb 45  160 130   # teal-green
  'KB' = Rgb 120 175 220   # powder blue
  'VC' = Rgb 0   0   0     # black
  'AB' = Rgb 245 180 30    # amber
}
$catalystOrder = @('CN','KB','VC','AB')

# Data sheet column blocks in LSV_comparison.xlsx (row 1=label, row 2=hdr, row3+=data).
$condBlock = [ordered]@{
  'Air_0bp'  = @{ col=1;  label='Air, 0 bar';   gas='Air'; bar='0 bar_g'   }
  'Air_15bp' = @{ col=4;  label='Air, 1.5 bar'; gas='Air'; bar='1.5 bar_g' }
  'O2_0bp'   = @{ col=7;  label='O2, 0 bar';    gas='O2';  bar='0 bar_g'   }
  'O2_15bp'  = @{ col=10; label='O2, 1.5 bar';  gas='O2';  bar='1.5 bar_g' }
}

# --- Read all required data series from LSV_comparison.xlsx ---
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false

$srcWb = $xl.Workbooks.Open($srcWbPath, 0, $true)

# series[$cat][$cond] = @{ I=[double[]], V=[double[]], P=[double[]], MaxIdx=int, MaxI, MaxV, MaxP }
$series = @{}
foreach ($cat in $catalystOrder) {
  $series[$cat] = @{}
  $sample = $catalystChoice[$cat].Sample
  $sheetName = "${cat}_${sample}"
  $ws = $srcWb.Worksheets.Item($sheetName)
  foreach ($condKey in $condBlock.Keys) {
    $col = $condBlock[$condKey].col
    $lastRow = $ws.Cells.Item($ws.Rows.Count, $col).End(-4162).Row
    if ($lastRow -lt 3) {
      $series[$cat][$condKey] = @{ I=@(); V=@(); P=@(); MaxIdx=-1 }
      continue
    }
    $rng = $ws.Range($ws.Cells.Item(3, $col), $ws.Cells.Item($lastRow, $col + 1))
    $vals = $rng.Value2
    $n = $lastRow - 2
    $I = New-Object 'double[]' $n
    $V = New-Object 'double[]' $n
    $P = New-Object 'double[]' $n
    $maxP = -1.0; $maxIdx = -1
    for ($k = 1; $k -le $n; $k++) {
      $iVal = [double]$vals[$k,1]
      $vVal = [double]$vals[$k,2]
      $pVal = $iVal * $vVal   # mA/cm^2 * V = mW/cm^2
      $I[$k-1] = $iVal
      $V[$k-1] = $vVal
      $P[$k-1] = $pVal
      if ($pVal -gt $maxP) { $maxP = $pVal; $maxIdx = $k - 1 }
    }
    $series[$cat][$condKey] = @{
      I = $I; V = $V; P = $P
      MaxIdx = $maxIdx
      MaxI   = if ($maxIdx -ge 0) { $I[$maxIdx] } else { [double]::NaN }
      MaxV   = if ($maxIdx -ge 0) { $V[$maxIdx] } else { [double]::NaN }
      MaxP   = if ($maxIdx -ge 0) { $P[$maxIdx] } else { [double]::NaN }
    }
  }
}
$srcWb.Close($false)

Write-Output 'Series read; max-power points:'
foreach ($cat in $catalystOrder) {
  foreach ($condKey in $condBlock.Keys) {
    $s = $series[$cat][$condKey]
    if ($s.MaxIdx -ge 0) {
      Write-Output ("  {0,-3} {1,-9}: i*V max -> P={2,7:F1} mW/cm^2 at i={3,7:F1} mA/cm^2, V={4,5:F3}" -f $cat, $condKey, $s.MaxP, $s.MaxI, $s.MaxV)
    }
  }
}

# --- Build output workbook ---
$outWb = $xl.Workbooks.Add()
while ($outWb.Worksheets.Count -gt 1) { $outWb.Worksheets.Item($outWb.Worksheets.Count).Delete() }
$readme = $outWb.Worksheets.Item(1)
$readme.Name = 'README'
$readme.Cells.Item(1,1) = 'LSV reference-style figures'
$readme.Cells.Item(2,1) = 'Series: CN BM (green), KB BM (light blue), VC Polyol (black, dotted), AB Polyol (amber, dotted)'
$readme.Cells.Item(3,1) = 'Left panel: Cell Voltage vs Current Density. Right panel: Power Density vs Current Density.'
$readme.Cells.Item(4,1) = 'Markers on LSV indicate the V at peak power.'
$readme.Cells.Item(5,1) = 'Generated: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm')

# Excel constants we need.
$xlScatterSmoothNoMarkers = 73
$xlNone        = -4142
$xlCircle      = 8
$xlContinuous  = 1
$xlLeft        = -4131
$xlRight       = -4152
$xlTop         = -4160
$xlBottom      = -4107
$xlInside      = 2
$xlOutsideEnd  = 4
$msoSolid      = 1
$msoLineSolid  = 1
$msoLineRoundDot = 3
$xlPrimary     = 1
$xlValue       = 2
$xlCategory    = 1

function Set-ChartCommon($chart, $title, $xMax, $yMin, $yMax, $yMajor, $yTitle) {
  $xMaxD   = [double]$xMax
  $yMinD   = [double]$yMin
  $yMaxD   = [double]$yMax
  $yMajorD = [double]$yMajor

  $chart.HasTitle = $false
  $chart.ChartArea.Format.Line.Visible = $false
  $chart.PlotArea.Format.Line.Visible = $true
  $chart.PlotArea.Format.Line.ForeColor.RGB = 0
  $chart.PlotArea.Format.Line.Weight = 1.25

  # X axis
  $xAxis = $chart.Axes($xlCategory, $xlPrimary)
  try { $xAxis.MaximumScale = $xMaxD } catch { Write-Output ("X MaxScale set failed: " + $_.Exception.Message) }
  try { $xAxis.MinimumScale = [double]0.0 } catch { Write-Output ("X MinScale set failed: " + $_.Exception.Message) }
  try { $xAxis.MajorUnit = [double]500.0 } catch {}
  try { $xAxis.MinorUnit = [double]100.0 } catch {}
  $xAxis.HasMajorGridlines = $false
  $xAxis.HasMinorGridlines = $false
  $xAxis.HasTitle = $true
  $xAxis.AxisTitle.Text = 'Current Density (mAcm' + [char]0x207B + [char]0x00B2 + ')'
  $xAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = $true
  $xAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
  $xAxis.TickLabels.Font.Bold = $true
  $xAxis.TickLabels.Font.Size = 10
  $xAxis.MajorTickMark = $xlInside
  $xAxis.Format.Line.ForeColor.RGB = 0
  $xAxis.Format.Line.Weight = 1.25

  # Y axis
  $yAxis = $chart.Axes($xlValue, $xlPrimary)
  try { $yAxis.MaximumScale = $yMaxD } catch { Write-Output ("Y MaxScale set failed: " + $_.Exception.Message) }
  try { $yAxis.MinimumScale = $yMinD } catch { Write-Output ("Y MinScale set failed: " + $_.Exception.Message) }
  try { $yAxis.MajorUnit = $yMajorD } catch { Write-Output ("Y MajorUnit set failed: " + $_.Exception.Message) }
  $yAxis.HasMajorGridlines = $false
  $yAxis.HasMinorGridlines = $false
  $yAxis.HasTitle = $true
  $yAxis.AxisTitle.Text = $yTitle
  $yAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = $true
  $yAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
  $yAxis.TickLabels.Font.Bold = $true
  $yAxis.TickLabels.Font.Size = 10
  $yAxis.MajorTickMark = $xlInside
  $yAxis.Format.Line.ForeColor.RGB = 0
  $yAxis.Format.Line.Weight = 1.25
  $chart.ChartArea.Format.Fill.ForeColor.RGB = 0xFFFFFF
}

function Add-LineSeries($chart, $name, $xRng, $yRng, $color, $isDotted, [bool]$hideFromLegend = $false) {
  $sNew = $chart.SeriesCollection().NewSeries()
  $sNew.Name = $name
  $sNew.XValues = $xRng
  $sNew.Values  = $yRng
  $sNew.Format.Line.Visible = $true
  $sNew.Format.Line.ForeColor.RGB = $color
  $sNew.Format.Line.Weight = 2.25
  if ($isDotted) {
    $sNew.Format.Line.DashStyle = $msoLineRoundDot
  } else {
    $sNew.Format.Line.DashStyle = $msoLineSolid
  }
  try { $sNew.MarkerStyle = $xlNone } catch {}
  if ($hideFromLegend) {
    try { $chart.Legend.LegendEntries($chart.SeriesCollection().Count).Delete() } catch {}
  }
}

function Add-MarkerSeries($chart, $name, $xRng, $yRng, $color) {
  $sNew = $chart.SeriesCollection().NewSeries()
  $sNew.Name = $name
  $sNew.XValues = $xRng
  $sNew.Values  = $yRng
  $sNew.Format.Line.Visible = $false
  try { $sNew.MarkerStyle = $xlCircle } catch {}
  try { $sNew.MarkerSize  = 8 } catch {}
  try { $sNew.Format.Fill.Visible = $true; $sNew.Format.Fill.ForeColor.RGB = $color } catch {}
  try { $sNew.MarkerForegroundColor = $color; $sNew.MarkerBackgroundColor = $color } catch {}
  try { $chart.Legend.LegendEntries($chart.SeriesCollection().Count).Delete() } catch {}
}

# Process each condition: build sheet, write data, draw two charts, export PNGs.
foreach ($condKey in $condBlock.Keys) {
  $info = $condBlock[$condKey]
  $sheetName = $condKey -replace '_',' '
  $ws = $outWb.Worksheets.Add()
  [void]$ws.Move([System.Reflection.Missing]::Value, $outWb.Worksheets.Item($outWb.Worksheets.Count))
  $ws.Name = $sheetName

  # Layout per catalyst: 3 cols (i, V, P) + 1 spacer.
  $colMap = @{}
  $col = 1
  foreach ($cat in $catalystOrder) {
    $sample = $catalystChoice[$cat].Sample
    $hdrLabel = "$cat $sample"
    $ws.Cells.Item(1, $col) = $hdrLabel
    $ws.Cells.Item(1, $col).Font.Bold = $true
    [void]$ws.Range($ws.Cells.Item(1,$col), $ws.Cells.Item(1,$col+2)).Merge()
    $ws.Cells.Item(2, $col)   = 'i (mA/cm^2)'
    $ws.Cells.Item(2, $col+1) = 'V (V)'
    $ws.Cells.Item(2, $col+2) = 'P (mW/cm^2)'
    $s = $series[$cat][$condKey]
    $n = $s.I.Count
    if ($n -gt 0) {
      $arr = New-Object 'object[,]' $n, 3
      for ($i = 0; $i -lt $n; $i++) {
        $arr[$i,0] = $s.I[$i]
        $arr[$i,1] = $s.V[$i]
        $arr[$i,2] = $s.P[$i]
      }
      [void]($ws.Range($ws.Cells.Item(3, $col), $ws.Cells.Item(2 + $n, $col + 2)).Value2 = $arr)
    }
    $colMap[$cat] = $col
    $col += 4
  }

  # Max-P marker block: cols starting at $col (col 17 by default).
  $markerStart = $col + 1
  $ws.Cells.Item(1, $markerStart) = 'Max-P markers'
  $ws.Cells.Item(1, $markerStart).Font.Bold = $true
  [void]$ws.Range($ws.Cells.Item(1, $markerStart), $ws.Cells.Item(1, $markerStart + 2)).Merge()
  $ws.Cells.Item(2, $markerStart    ) = 'i_at_maxP'
  $ws.Cells.Item(2, $markerStart + 1) = 'V_at_maxP'
  $ws.Cells.Item(2, $markerStart + 2) = 'P_max'
  $r = 3
  foreach ($cat in $catalystOrder) {
    $s = $series[$cat][$condKey]
    $ws.Cells.Item($r, $markerStart    ) = $s.MaxI
    $ws.Cells.Item($r, $markerStart + 1) = $s.MaxV
    $ws.Cells.Item($r, $markerStart + 2) = $s.MaxP
    $ws.Cells.Item($r, $markerStart + 4) = $cat
    $r++
  }

  # ---- LSV chart (left) ----
  $widthPx = 520; $heightPx = 380; $gapPx = 40
  $leftLSV = 5; $topCharts = 30
  $shapeLSV = $ws.Shapes.AddChart2(-1, $xlScatterSmoothNoMarkers, $leftLSV, $topCharts, $widthPx, $heightPx)
  $chartLSV = $shapeLSV.Chart
  while ($chartLSV.SeriesCollection().Count -gt 0) { [void]$chartLSV.SeriesCollection().Item(1).Delete() }

  # Add line series, then marker series for each catalyst.
  foreach ($cat in $catalystOrder) {
    $sample = $catalystChoice[$cat].Sample
    $col = $colMap[$cat]
    $s = $series[$cat][$condKey]
    $n = $s.I.Count
    if ($n -lt 1) { continue }
    $xRng = $ws.Range($ws.Cells.Item(3, $col), $ws.Cells.Item(2 + $n, $col))
    $yRng = $ws.Range($ws.Cells.Item(3, $col + 1), $ws.Cells.Item(2 + $n, $col + 1))
    $name = "$cat $sample"
    Add-LineSeries $chartLSV $name $xRng $yRng $catalystColor[$cat] ($catalystChoice[$cat].LineStyle -eq 'dotted')
  }
  # Add max-P marker series (after all lines so they render on top).
  $r = 3
  foreach ($cat in $catalystOrder) {
    $xR = $ws.Range($ws.Cells.Item($r, $markerStart    ), $ws.Cells.Item($r, $markerStart    ))
    $yR = $ws.Range($ws.Cells.Item($r, $markerStart + 1), $ws.Cells.Item($r, $markerStart + 1))
    Add-MarkerSeries $chartLSV "$cat maxP" $xR $yR $catalystColor[$cat]
    $r++
  }

  $chartLSV.HasLegend = $true
  $chartLSV.Legend.Position = $xlTop
  $chartLSV.Legend.Font.Size = 10
  $chartLSV.Legend.Font.Bold = $true
  # Try moving legend to top-left inside plot.
  try {
    $chartLSV.Legend.Left   = 50
    $chartLSV.Legend.Top    = 25
    $chartLSV.Legend.Width  = 130
    $chartLSV.Legend.Height = 70
  } catch {}

  Set-ChartCommon $chartLSV $info.label 3500 0.2 1.2 0.2 'Cell Voltage (V)'

  # ---- Power chart (right) ----
  $leftPow = $leftLSV + $widthPx + $gapPx
  $shapePow = $ws.Shapes.AddChart2(-1, $xlScatterSmoothNoMarkers, $leftPow, $topCharts, $widthPx, $heightPx)
  $chartPow = $shapePow.Chart
  while ($chartPow.SeriesCollection().Count -gt 0) { [void]$chartPow.SeriesCollection().Item(1).Delete() }

  foreach ($cat in $catalystOrder) {
    $sample = $catalystChoice[$cat].Sample
    $col = $colMap[$cat]
    $s = $series[$cat][$condKey]
    $n = $s.I.Count
    if ($n -lt 1) { continue }
    $xRng = $ws.Range($ws.Cells.Item(3, $col), $ws.Cells.Item(2 + $n, $col))
    $pRng = $ws.Range($ws.Cells.Item(3, $col + 2), $ws.Cells.Item(2 + $n, $col + 2))
    $name = "$cat $sample"
    Add-LineSeries $chartPow $name $xRng $pRng $catalystColor[$cat] ($catalystChoice[$cat].LineStyle -eq 'dotted')
  }
  $chartPow.HasLegend = $false
  Set-ChartCommon $chartPow '' 3500 0 1200 200 'Power Density (mWcm' + [char]0x207B + [char]0x00B2 + ')'

  # ---- Annotation textbox in LSV chart (top-right) ----
  $gas = $info.gas
  $bar = $info.bar
  $annot = "H" + [char]0x2082 + "/" + $gas + "`r`nRH100`r`nPt 5 wt%`r`n0.05 mg" + [char]0x209A + [char]0x209C + "/cm" + [char]0x00B2 + "`r`nIC 0.8, N212`r`n" + $bar
  # Add the textbox as a shape on the LSV chart's chart object.
  $tb = $chartLSV.Shapes.AddTextbox(1, 360, 25, 150, 110)  # msoTextOrientationHorizontal=1
  $tb.TextFrame.Characters().Text = $annot
  $tb.TextFrame.Characters().Font.Bold = $true
  $tb.TextFrame.Characters().Font.Size = 10
  $tb.Line.Visible = $false
  $tb.Fill.Visible = $false

  # Export each chart as PNG.
  $cleanCond = $condKey
  $pLSV = Join-Path $pngDir ("LSV_{0}_LSV.png" -f $cleanCond)
  $pPow = Join-Path $pngDir ("LSV_{0}_Power.png" -f $cleanCond)
  [void]$chartLSV.Export($pLSV, 'PNG')
  [void]$chartPow.Export($pPow, 'PNG')

  Write-Output ("Built sheet '{0}': LSV+Power charts and PNGs exported." -f $sheetName)
}

# Save and close.
if (Test-Path -LiteralPath $outFile) { Remove-Item -LiteralPath $outFile -Force }
$outWb.SaveAs($outFile, 51)
$outWb.Close($false)
$xl.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

# --- Stitch each (LSV, Power) pair into a single PNG side-by-side. ---
Add-Type -AssemblyName System.Drawing
foreach ($condKey in $condBlock.Keys) {
  $pLSV = Join-Path $pngDir ("LSV_{0}_LSV.png"   -f $condKey)
  $pPow = Join-Path $pngDir ("LSV_{0}_Power.png" -f $condKey)
  $pOut = Join-Path $pngDir ("LSV_{0}_combined.png" -f $condKey)
  if (-not (Test-Path $pLSV) -or -not (Test-Path $pPow)) { continue }
  $imgL = [System.Drawing.Image]::FromFile($pLSV)
  $imgR = [System.Drawing.Image]::FromFile($pPow)
  $gap = 24
  $w = $imgL.Width + $imgR.Width + $gap
  $h = [Math]::Max($imgL.Height, $imgR.Height)
  $bmp = New-Object System.Drawing.Bitmap($w, $h)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::White)
  $g.DrawImage($imgL, 0, 0)
  $g.DrawImage($imgR, $imgL.Width + $gap, 0)
  $g.Dispose()
  $bmp.Save($pOut, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose(); $imgL.Dispose(); $imgR.Dispose()
  Write-Output ("Stitched: " + (Split-Path $pOut -Leaf))
}

Write-Output ''
Write-Output ("Saved workbook: " + $outFile)
Write-Output ("PNG output dir: " + $pngDir)
