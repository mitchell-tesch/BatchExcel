<#
.SYNOPSIS
    Generates the BatchExcel application icon (PNG + multi-resolution ICO).

.DESCRIPTION
    Modern flat-ish design:
        * Excel-green rounded-square tile with a soft diagonal gradient and a
          subtle top-light glass highlight.
        * Floating white spreadsheet "card" with a coloured Excel-green
          header bar and a clean 4-col x 5-row grid (no row/column letter
          labels - they alias badly at small sizes). Two cells are tinted
          to read as "input" / "output" without being noisy.
        * Bold yellow-orange lightning bolt with a soft glow, a thin dark
          contour for definition, and a bright inner highlight stripe.

    Master canvas is 2048 x 2048 by default for crisp downscaling. The ICO
    embeds PNG-encoded frames up to 512 px (Windows uses up to 256, larger
    frames are kept so shell extensions / future scaling can use them).

    Outputs:
        BatchExcel/Resources/app_icon.png   (MasterSize, default 2048)
        BatchExcel/Resources/app_icon.ico   (16, 20, 24, 32, 40, 48, 64, 96, 128, 256, 512)
        docs/app_icon.png                   (copy of the master PNG)
#>

[CmdletBinding()]
param(
    [int] $MasterSize = 2048
)

Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

$repoRoot     = Split-Path -Parent $PSScriptRoot
$resourcesDir = Join-Path $repoRoot 'BatchExcel\Resources'
$docsDir      = Join-Path $repoRoot 'docs'
$pngPath      = Join-Path $resourcesDir 'app_icon.png'
$icoPath      = Join-Path $resourcesDir 'app_icon.ico'
$docsPngPath  = Join-Path $docsDir 'app_icon.png'

New-Item -ItemType Directory -Force -Path $resourcesDir, $docsDir | Out-Null

function Get-RoundedRectPath {
    param([single]$X, [single]$Y, [single]$W, [single]$H, [single]$R)
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $p.AddArc($X,           $Y,           $R, $R, 180, 90)
    $p.AddArc($X + $W - $R, $Y,           $R, $R, 270, 90)
    $p.AddArc($X + $W - $R, $Y + $H - $R, $R, $R,   0, 90)
    $p.AddArc($X,           $Y + $H - $R, $R, $R,  90, 90)
    $p.CloseFigure()
    return $p
}

function New-IconBitmap {
    param([int] $Size)

    # Always render the artwork at the master canvas size then downscale to
    # $Size with HighQualityBicubic. This keeps thin elements (grid lines,
    # bolt contour) anti-aliased instead of dropping pixels when drawn
    # directly at tiny sizes like 16 or 24.
    $canvasSize = [Math]::Max($Size, 1024)

    $bigBmp = New-Object System.Drawing.Bitmap $canvasSize, $canvasSize, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g      = [System.Drawing.Graphics]::FromImage($bigBmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::Transparent)

    # All coordinates are expressed on a 1024 design canvas and scaled.
    $s = $canvasSize / 1024.0
    $px = { param([double]$v) [single]($v * $s) }

    # --- Tile background ------------------------------------------------------
    $bgX = & $px 40;  $bgY = & $px 40
    $bgW = & $px 944; $bgH = & $px 944
    $bgR = & $px 230   # generous corner radius for a modern Fluent / iOS feel

    # Soft drop shadow under the tile.
    $shadowOffset = & $px 18
    $shadowPath = Get-RoundedRectPath $bgX ($bgY + $shadowOffset) $bgW $bgH $bgR
    $shBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(95, 0, 0, 0))
    $g.FillPath($shBrush, $shadowPath)
    $shBrush.Dispose(); $shadowPath.Dispose()

    # Main tile - vivid Excel-green diagonal gradient.
    $bgPath = Get-RoundedRectPath $bgX $bgY $bgW $bgH $bgR
    $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.PointF $bgX, $bgY),
        (New-Object System.Drawing.PointF ($bgX + $bgW), ($bgY + $bgH)),
        ([System.Drawing.Color]::FromArgb(255,  46, 178,  98)),    # bright Excel green
        ([System.Drawing.Color]::FromArgb(255,  16,  92,  52)))    # deep forest
    $g.FillPath($bgBrush, $bgPath)
    $bgBrush.Dispose()

    # Top-light glass highlight - a soft white wedge in the upper half.
    $glassState = $g.Save()
    $g.SetClip($bgPath)
    $glassPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $glassPath.AddEllipse((& $px -100), (& $px -480), (& $px 1224), (& $px 760))
    $glassBrush = New-Object System.Drawing.Drawing2D.PathGradientBrush $glassPath
    $glassBrush.CenterColor = [System.Drawing.Color]::FromArgb(70, 255, 255, 255)
    $glassBrush.SurroundColors = ,([System.Drawing.Color]::FromArgb(0, 255, 255, 255))
    $g.FillPath($glassBrush, $glassPath)
    $glassBrush.Dispose(); $glassPath.Dispose()
    $g.Restore($glassState)

    # Inner hairline highlight on the tile edge.
    $innerHl = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(70, 255, 255, 255)), (& $px 3)
    $g.DrawPath($innerHl, $bgPath)
    $innerHl.Dispose()

    # --- Spreadsheet card -----------------------------------------------------
    # Centred white card with its own shadow, a green header bar and a grid.
    $cardX = & $px 160
    $cardY = & $px 215
    $cardW = & $px 704
    $cardH = & $px 614
    $cardR = & $px 50

    # Card shadow.
    $cardShadow = Get-RoundedRectPath $cardX ($cardY + (& $px 16)) $cardW $cardH $cardR
    $cardShBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(85, 0, 0, 0))
    $g.FillPath($cardShBrush, $cardShadow)
    $cardShBrush.Dispose(); $cardShadow.Dispose()

    # Card body.
    $cardPath = Get-RoundedRectPath $cardX $cardY $cardW $cardH $cardR
    $cardBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 252, 253, 251))
    $g.FillPath($cardBrush, $cardPath)
    $cardBrush.Dispose()

    # Clip subsequent decoration to the card's rounded shape.
    $cardState = $g.Save()
    $g.SetClip($cardPath)

    # Header bar (slightly darker green than the tile so it stays distinct).
    $headerH = & $px 110
    $headerBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.PointF $cardX, $cardY),
        (New-Object System.Drawing.PointF ($cardX + $cardW), $cardY),
        ([System.Drawing.Color]::FromArgb(255, 30, 130, 75)),
        ([System.Drawing.Color]::FromArgb(255, 22, 102, 58)))
    $g.FillRectangle($headerBrush, $cardX, $cardY, $cardW, $headerH)
    $headerBrush.Dispose()

    # Three traffic-light dots on the header - very subtle modern flourish.
    $dotR = & $px 16
    $dotY = $cardY + ($headerH / 2) - ($dotR / 2)
    $dotColors = @(
        [System.Drawing.Color]::FromArgb(255, 255, 100,  95),
        [System.Drawing.Color]::FromArgb(255, 255, 196,  60),
        [System.Drawing.Color]::FromArgb(255, 130, 220, 130))
    for ($d = 0; $d -lt 3; $d++) {
        $dotBrush = New-Object System.Drawing.SolidBrush $dotColors[$d]
        $dotX = $cardX + (& $px 40) + ($d * (& $px 44))
        $g.FillEllipse($dotBrush, $dotX, $dotY, $dotR, $dotR)
        $dotBrush.Dispose()
    }

    # Grid area below the header.
    $gridPadX    = & $px 36
    $gridPadTop  = & $px 30
    $gridPadBot  = & $px 40
    $gridLeft    = $cardX + $gridPadX
    $gridRight   = $cardX + $cardW - $gridPadX
    $gridTop     = $cardY + $headerH + $gridPadTop
    $gridBottom  = $cardY + $cardH   - $gridPadBot

    $cols = 4
    $rows = 5
    $cellW = ($gridRight  - $gridLeft) / $cols
    $cellH = ($gridBottom - $gridTop)  / $rows

    # Header strip immediately under the title bar (column-header band).
    $colHeadH = & $px 56
    $colHeadBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 230, 244, 235))
    $g.FillRectangle($colHeadBrush, $gridLeft, $gridTop, ($gridRight - $gridLeft), $colHeadH)
    $colHeadBrush.Dispose()

    # Highlighted cells (input = soft blue, output = soft green).
    $inBrush  = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220, 196, 222, 245))
    $outBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220, 178, 226, 188))
    # Inputs in column 0 (rows 2 and 3 - below the column-header strip).
    $g.FillRectangle($inBrush,
        $gridLeft + 0 * $cellW,
        $gridTop + $colHeadH + 0 * ($cellH - $colHeadH / $rows),
        $cellW,
        ($gridBottom - ($gridTop + $colHeadH)) / ($rows - 1))
    # Outputs in last column (rows 2 and 4).
    $usableTop = $gridTop + $colHeadH
    $usableH   = $gridBottom - $usableTop
    $rowDataH  = $usableH / ($rows - 1)

    # Repaint cleanly: clear the simplistic highlight above and use a proper grid.
    $g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 252, 253, 251))),
        $gridLeft, $usableTop, ($gridRight - $gridLeft), $usableH)

    # Now lay out the data cells: 4 columns x 4 rows (under the col header).
    $dataRows = 4
    $dataRowH = $usableH / $dataRows

    # Tinted cells (chosen to look like a couple of inputs and outputs).
    $tinted = @(
        @{ Col=0; Row=0; Brush=$inBrush  },
        @{ Col=0; Row=2; Brush=$inBrush  },
        @{ Col=3; Row=1; Brush=$outBrush },
        @{ Col=3; Row=3; Brush=$outBrush }
    )
    foreach ($t in $tinted) {
        $g.FillRectangle($t.Brush,
            ($gridLeft + $t.Col * $cellW),
            ($usableTop + $t.Row * $dataRowH),
            $cellW, $dataRowH)
    }
    $inBrush.Dispose(); $outBrush.Dispose()

    # Faint zebra striping on column-header strip cells (subtle dividers).
    $headerDivPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 195, 220, 205)), (& $px 2)
    for ($cc = 1; $cc -lt $cols; $cc++) {
        $x = $gridLeft + $cc * $cellW
        $g.DrawLine($headerDivPen, $x, $gridTop, $x, ($gridTop + $colHeadH))
    }
    $headerDivPen.Dispose()

    # Grid lines (clean, light).
    $gridPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 200, 215, 205)), (& $px 2.5)
    for ($cc = 0; $cc -le $cols; $cc++) {
        $x = $gridLeft + $cc * $cellW
        $g.DrawLine($gridPen, $x, $usableTop, $x, $gridBottom)
    }
    for ($rr = 0; $rr -le $dataRows; $rr++) {
        $y = $usableTop + $rr * $dataRowH
        $g.DrawLine($gridPen, $gridLeft, $y, $gridRight, $y)
    }
    # Stronger divider under the column-header band.
    $bandPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 110, 165, 135)), (& $px 3)
    $g.DrawLine($bandPen, $gridLeft, $usableTop, $gridRight, $usableTop)
    $bandPen.Dispose()
    $gridPen.Dispose()

    $g.Restore($cardState)

    # Crisp card border.
    $cardEdge = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 70, 130, 95)), (& $px 3)
    $g.DrawPath($cardEdge, $cardPath)
    $cardEdge.Dispose()
    $cardPath.Dispose()
    $bgPath.Dispose()

    # --- Lightning bolt overlay ----------------------------------------------
    # Tilted modern bolt with a soft glow and a bright inner highlight.
    $boltPts = @(
        (New-Object System.Drawing.PointF (& $px 640), (& $px 120)),
        (New-Object System.Drawing.PointF (& $px 300), (& $px 605)),
        (New-Object System.Drawing.PointF (& $px 505), (& $px 605)),
        (New-Object System.Drawing.PointF (& $px 380), (& $px 970)),
        (New-Object System.Drawing.PointF (& $px 760), (& $px 470)),
        (New-Object System.Drawing.PointF (& $px 555), (& $px 470)),
        (New-Object System.Drawing.PointF (& $px 700), (& $px 120))
    )
    $boltPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $boltPath.AddPolygon([System.Drawing.PointF[]] $boltPts)
    $boltPath.CloseFigure()

    # Glow: draw the path repeatedly with a thick translucent yellow pen.
    for ($gi = 6; $gi -ge 1; $gi--) {
        $glowPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb([int](16 + $gi * 6), 255, 220, 90)), (& $px (8 * $gi))
        $glowPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
        $g.DrawPath($glowPen, $boltPath)
        $glowPen.Dispose()
    }

    # Hard drop shadow under the bolt.
    $shadowPts = foreach ($p in $boltPts) {
        New-Object System.Drawing.PointF ($p.X + (& $px 10)), ($p.Y + (& $px 14))
    }
    $boltShadow = New-Object System.Drawing.Drawing2D.GraphicsPath
    $boltShadow.AddPolygon([System.Drawing.PointF[]] $shadowPts)
    $boltShBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(120, 0, 0, 0))
    $g.FillPath($boltShBrush, $boltShadow)
    $boltShBrush.Dispose(); $boltShadow.Dispose()

    # Bolt body - golden gradient.
    $boltBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.PointF (& $px 300), (& $px 120)),
        (New-Object System.Drawing.PointF (& $px 760), (& $px 970)),
        ([System.Drawing.Color]::FromArgb(255, 255, 240, 110)),
        ([System.Drawing.Color]::FromArgb(255, 255, 138,  18)))
    $g.FillPath($boltBrush, $boltPath)
    $boltBrush.Dispose()

    # Inner highlight stripe (offset, partial) for a glossy 3D feel.
    $hlState = $g.Save()
    $g.SetClip($boltPath)
    $hlPts = @(
        (New-Object System.Drawing.PointF (& $px 612), (& $px 130)),
        (New-Object System.Drawing.PointF (& $px 360), (& $px 580)),
        (New-Object System.Drawing.PointF (& $px 470), (& $px 580)),
        (New-Object System.Drawing.PointF (& $px 410), (& $px 760))
    )
    $hlPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $hlPath.AddPolygon([System.Drawing.PointF[]] $hlPts)
    $hlBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.PointF (& $px 350), (& $px 120)),
        (New-Object System.Drawing.PointF (& $px 600), (& $px 760)),
        ([System.Drawing.Color]::FromArgb(180, 255, 255, 230)),
        ([System.Drawing.Color]::FromArgb(30,  255, 255, 230)))
    $g.FillPath($hlBrush, $hlPath)
    $hlBrush.Dispose(); $hlPath.Dispose()
    $g.Restore($hlState)

    # Thin dark contour for definition (kept thin so the bolt looks modern, not cartoonish).
    $boltEdge = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 110, 60, 0)), (& $px 5)
    $boltEdge.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    $g.DrawPath($boltEdge, $boltPath)
    $boltEdge.Dispose()
    $boltPath.Dispose()

    $g.Dispose()

    if ($canvasSize -eq $Size) {
        return $bigBmp
    }

    # Downscale to the requested size for crisp small-icon rendering.
    $finalBmp = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $fg = [System.Drawing.Graphics]::FromImage($finalBmp)
    $fg.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $fg.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $fg.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $fg.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $fg.Clear([System.Drawing.Color]::Transparent)
    $fg.DrawImage($bigBmp, (New-Object System.Drawing.Rectangle 0, 0, $Size, $Size))
    $fg.Dispose()
    $bigBmp.Dispose()
    return $finalBmp
}

function Save-Png {
    param([System.Drawing.Bitmap]$Bitmap, [string]$Path)
    $Bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
}

function Save-Ico {
    param([string]$Path, [int[]]$Sizes)

    $entries = @()
    foreach ($sz in $Sizes) {
        $bmp = New-IconBitmap -Size $sz
        $ms  = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        $entries += ,@{ Size = $sz; Bytes = $ms.ToArray() }
        $ms.Dispose()
    }

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    $bw = New-Object System.IO.BinaryWriter $fs
    try {
        $bw.Write([uint16]0)
        $bw.Write([uint16]1)
        $bw.Write([uint16]$entries.Count)

        $offset = 6 + (16 * $entries.Count)
        foreach ($e in $entries) {
            $dim = if ($e.Size -ge 256) { 0 } else { $e.Size }
            $bw.Write([byte]$dim)
            $bw.Write([byte]$dim)
            $bw.Write([byte]0)
            $bw.Write([byte]0)
            $bw.Write([uint16]1)
            $bw.Write([uint16]32)
            $bw.Write([uint32]$e.Bytes.Length)
            $bw.Write([uint32]$offset)
            $offset += $e.Bytes.Length
        }
        foreach ($e in $entries) { $bw.Write($e.Bytes) }
    }
    finally {
        $bw.Dispose()
        $fs.Dispose()
    }
}

Write-Host "Rendering master ($MasterSize x $MasterSize)..."
$master = New-IconBitmap -Size $MasterSize

Write-Host "Writing $pngPath"
Save-Png -Bitmap $master -Path $pngPath

Write-Host "Writing $docsPngPath"
Save-Png -Bitmap $master -Path $docsPngPath

$master.Dispose()

Write-Host "Writing $icoPath"
Save-Ico -Path $icoPath -Sizes @(16, 20, 24, 32, 40, 48, 64, 96, 128, 256, 512)

Write-Host 'Done.' -ForegroundColor Green


