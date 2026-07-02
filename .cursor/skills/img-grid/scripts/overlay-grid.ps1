#Requires -Version 5.1
<#
.SYNOPSIS
    Overlay a numbered grid on an image to help determine column/row proportions.

.DESCRIPTION
    Рисует пронумерованную сетку на изображении печатной формы.
    Помогает определить границы колонок, их пропорции и span-ы
    для генерации макета табличного документа (MXL).

    Числа выводятся в поле отступа (вне области изображения),
    чтобы не перекрывать содержимое формы.

.PARAMETER Image
    Путь к исходному изображению (PNG, JPG).

.PARAMETER Cols
    Количество вертикальных делений (по умолчанию: 50).

.PARAMETER Rows
    Количество горизонтальных делений (0 = авто: квадратные ячейки).

.PARAMETER Output
    Путь для сохранения результата (по умолчанию: <name>-grid.<ext>).

.EXAMPLE
    .\overlay-grid.ps1 -Image "form.png" -Cols 50
    .\overlay-grid.ps1 -Image "form.png" -Cols 40 -Output "result.png"
#>
param(
    [Parameter(Mandatory)][string]$Image,
    [int]$Cols   = 50,
    [int]$Rows   = 0,
    [string]$Output = ""
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$MARGIN_TOP  = 20
$MARGIN_LEFT = 24

$ImagePath = Resolve-Path $Image

$src = [System.Drawing.Image]::FromFile($ImagePath.Path)
$sw  = $src.Width
$sh  = $src.Height

$stepX = $sw / $Cols
$actualRows = if ($Rows -eq 0) { [Math]::Round($sh / $stepX) } else { $Rows }
$stepY = $sh / $actualRows

$cw = $MARGIN_LEFT + $sw
$ch = $MARGIN_TOP  + $sh

$canvas = New-Object System.Drawing.Bitmap($cw, $ch, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g = [System.Drawing.Graphics]::FromImage($canvas)
$g.SmoothingMode    = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
$g.Clear([System.Drawing.Color]::White)
$g.DrawImage($src, $MARGIN_LEFT, $MARGIN_TOP, $sw, $sh)
$src.Dispose()

$font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Regular)

# Вертикальные линии + номера сверху
for ($i = 0; $i -le $Cols; $i++) {
    $x     = $MARGIN_LEFT + [Math]::Round($i * $stepX)
    $major = ($i % 10) -eq 0
    $mid   = ($i % 5)  -eq 0

    $alpha = if ($major) { 160 } elseif ($mid) { 110 } else { 40 }
    $lw    = if ($major -or $mid) { 2 } else { 1 }
    $pen   = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb($alpha, 220, 0, 0), $lw)
    $g.DrawLine($pen, $x, $MARGIN_TOP, $x, $ch)
    $pen.Dispose()

    if ($major -or $mid -or $stepX -ge 20) {
        $label  = "$i"
        $sz     = $g.MeasureString($label, $font)
        $tx     = $x - $sz.Width / 2
        $bAlpha = if ($major -or $mid) { 220 } else { 160 }
        $brush  = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($bAlpha, 200, 0, 0))
        $g.DrawString($label, $font, $brush, $tx, 2)
        $brush.Dispose()
    }
}

# Горизонтальные линии + номера слева
for ($j = 0; $j -le $actualRows; $j++) {
    $y     = $MARGIN_TOP + [Math]::Round($j * $stepY)
    $major = ($j % 10) -eq 0
    $mid   = ($j % 5)  -eq 0

    $alpha = if ($major) { 160 } elseif ($mid) { 110 } else { 20 }
    $lw    = if ($major -or $mid) { 2 } else { 1 }
    $pen   = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb($alpha, 0, 0, 210), $lw)
    $g.DrawLine($pen, $MARGIN_LEFT, $y, $cw, $y)
    $pen.Dispose()

    if ($major -or $mid -or $stepY -ge 20) {
        $label  = "$j"
        $sz     = $g.MeasureString($label, $font)
        $tx     = $MARGIN_LEFT - $sz.Width - 2
        $ty     = $y - $sz.Height / 2
        $bAlpha = if ($major -or $mid) { 220 } else { 160 }
        $brush  = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($bAlpha, 0, 0, 200))
        $g.DrawString($label, $font, $brush, $tx, $ty)
        $brush.Dispose()
    }
}

$g.Dispose()
$font.Dispose()

if ($Output -eq "") {
    $dir  = [System.IO.Path]::GetDirectoryName($ImagePath.Path)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($ImagePath.Path)
    $ext  = [System.IO.Path]::GetExtension($ImagePath.Path)
    $Output = Join-Path $dir "$name-grid$ext"
}

$canvas.Save($Output)
$canvas.Dispose()

Write-Host "Grid: $Cols x $actualRows cells"
Write-Host ("Cell size: {0:F1} x {1:F1} px" -f $stepX, $stepY)
Write-Host "Image: $sw x $sh px"
Write-Host "Saved: $Output"
