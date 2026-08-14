param([int]$Width = 1280, [int]$Height = 400)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$config = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'banner-config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$outDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'assets'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$outPath = Join-Path $outDir 'banner.png'

$bitmap = New-Object System.Drawing.Bitmap($Width, $Height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

$bgRect = New-Object System.Drawing.Rectangle(0, 0, $Width, $Height)
$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
  $bgRect,
  [System.Drawing.Color]::FromArgb(255, 13, 27, 42),
  [System.Drawing.Color]::FromArgb(255, 0, 105, 100),
  45.0)
$graphics.FillRectangle($brush, $bgRect)

$yahei = 'Microsoft YaHei'
$titleFont = New-Object System.Drawing.Font($yahei, 64, [System.Drawing.FontStyle]::Bold)
$taglineFont = New-Object System.Drawing.Font($yahei, 30, [System.Drawing.FontStyle]::Bold)
$subtitleFont = New-Object System.Drawing.Font($yahei, 20, [System.Drawing.FontStyle]::Regular)
$chipFont = New-Object System.Drawing.Font($yahei, 16, [System.Drawing.FontStyle]::Regular)

$white = [System.Drawing.Brushes]::White
$soft = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 200, 230, 225))

$graphics.DrawString($config.title, $titleFont, $white, 48, 42)
$graphics.DrawString($config.tagline, $taglineFont, $white, 52, 140)
$graphics.DrawString($config.subtitle, $subtitleFont, $soft, 54, 190)

$chipBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(70, 255, 255, 255))
$chipPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200, 255, 255, 255), 1)
$x = 48
$y = 300
$chipHeight = 34
foreach ($chip in $config.chips) {
  $size = $graphics.MeasureString($chip, $chipFont)
  $chipWidth = [int]($size.Width + 28)
  if ($x + $chipWidth -gt ($Width - 24)) {
    $x = 48
    $y += $chipHeight + 10
  }
  $rect = New-Object System.Drawing.Rectangle($x, $y, $chipWidth, $chipHeight)
  $path = New-Object System.Drawing.Drawing2D.GraphicsPath
  $radius = 17
  $diameter = $radius * 2
  $path.AddArc($rect.X, $rect.Y, $diameter, $diameter, 180, 90)
  $path.AddArc($rect.Right - $diameter, $rect.Y, $diameter, $diameter, 270, 90)
  $path.AddArc($rect.Right - $diameter, $rect.Bottom - $diameter, $diameter, $diameter, 0, 90)
  $path.AddArc($rect.X, $rect.Bottom - $diameter, $diameter, $diameter, 90, 90)
  $path.CloseFigure()
  $graphics.FillPath($chipBrush, $path)
  $graphics.DrawPath($chipPen, $path)
  $graphics.DrawString($chip, $chipFont, $soft, ($x + 14), ($y + 6))
  $x += $chipWidth + 10
}

$graphics.Flush()
$bitmap.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()
Write-Output "banner written to $outPath"
