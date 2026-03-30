param(
    [string]$IcoPath = (Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "Start-ReconcileTool.ico"),
    [string]$PreviewPath = (Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "Start-ReconcileTool-icon-preview.png")
)

Set-StrictMode -Version 3
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

function New-Color {
    param(
        [string]$Hex,
        [int]$Alpha = 255
    )

    $base = [System.Drawing.ColorTranslator]::FromHtml($Hex)
    return [System.Drawing.Color]::FromArgb($Alpha, $base.R, $base.G, $base.B)
}

function New-RoundedRectPath {
    param(
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [float]$Radius
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $diameter = [float]($Radius * 2)

    $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
    $path.AddArc($X + $Width - $diameter, $Y, $diameter, $diameter, 270, 90)
    $path.AddArc($X + $Width - $diameter, $Y + $Height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($X, $Y + $Height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()

    return $path
}

function Save-PngIconSetAsIco {
    param(
        [System.Collections.Generic.List[byte[]]]$ImageBytesList,
        [int[]]$Sizes,
        [string]$OutputPath
    )

    $directory = Split-Path -Path $OutputPath -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }

    $stream = New-Object System.IO.FileStream($OutputPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $writer = New-Object System.IO.BinaryWriter($stream)

    try {
        $count = $Sizes.Length
        $writer.Write([UInt16]0)
        $writer.Write([UInt16]1)
        $writer.Write([UInt16]$count)

        $offset = 6 + (16 * $count)
        for ($i = 0; $i -lt $count; $i++) {
            $size = $Sizes[$i]
            $bytes = $ImageBytesList[$i]
            $icoSizeByte = if ($size -ge 256) { [byte]0 } else { [byte]$size }

            $writer.Write($icoSizeByte)
            $writer.Write($icoSizeByte)
            $writer.Write([byte]0)
            $writer.Write([byte]0)
            $writer.Write([UInt16]1)
            $writer.Write([UInt16]32)
            $writer.Write([UInt32]$bytes.Length)
            $writer.Write([UInt32]$offset)

            $offset += $bytes.Length
        }

        for ($i = 0; $i -lt $count; $i++) {
            $writer.Write($ImageBytesList[$i])
        }
    }
    finally {
        $writer.Close()
        $stream.Close()
    }
}

$baseSize = 256
$bitmap = New-Object System.Drawing.Bitmap($baseSize, $baseSize, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

try {
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $graphics.Clear([System.Drawing.Color]::Transparent)

    $shadowPath = New-RoundedRectPath -X 16 -Y 18 -Width 224 -Height 224 -Radius 50
    $shadowBrush = New-Object System.Drawing.SolidBrush((New-Color -Hex "#08121A" -Alpha 45))
    $graphics.FillPath($shadowBrush, $shadowPath)

    $backgroundPath = New-RoundedRectPath -X 10 -Y 10 -Width 224 -Height 224 -Radius 48
    $backgroundBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Point(16, 10)),
        (New-Object System.Drawing.Point(224, 234)),
        (New-Color -Hex "#0D5E67"),
        (New-Color -Hex "#1CA37F")
    )
    $graphics.FillPath($backgroundBrush, $backgroundPath)

    $overlayBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Point(10, 20)),
        (New-Object System.Drawing.Point(210, 150)),
        (New-Color -Hex "#FFFFFF" -Alpha 56),
        (New-Color -Hex "#FFFFFF" -Alpha 0)
    )
    $graphics.FillEllipse($overlayBrush, 28, 20, 180, 100)

    $accentBrush = New-Object System.Drawing.SolidBrush((New-Color -Hex "#D8FFF4" -Alpha 34))
    $graphics.FillEllipse($accentBrush, 180, 32, 26, 26)
    $graphics.FillEllipse($accentBrush, 192, 62, 14, 14)

    $sheetShadowPath = New-RoundedRectPath -X 56 -Y 50 -Width 112 -Height 146 -Radius 24
    $sheetShadowBrush = New-Object System.Drawing.SolidBrush((New-Color -Hex "#04131A" -Alpha 28))
    $graphics.FillPath($sheetShadowBrush, $sheetShadowPath)

    $sheetPath = New-RoundedRectPath -X 50 -Y 44 -Width 112 -Height 146 -Radius 24
    $sheetBrush = New-Object System.Drawing.SolidBrush((New-Color -Hex "#F8FFFC"))
    $sheetBorderPen = New-Object System.Drawing.Pen((New-Color -Hex "#CDE9E4"), 2.0)
    $graphics.FillPath($sheetBrush, $sheetPath)
    $graphics.DrawPath($sheetBorderPen, $sheetPath)

    $headerBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Point(50, 44)),
        (New-Object System.Drawing.Point(162, 92)),
        (New-Color -Hex "#D6F8ED"),
        (New-Color -Hex "#ECFFFA")
    )
    $graphics.FillRectangle($headerBrush, 50, 44, 112, 34)

    $linePen = New-Object System.Drawing.Pen((New-Color -Hex "#7BCFB9"), 4.0)
    $linePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $linePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $graphics.DrawLine($linePen, 68, 64, 106, 64)
    $graphics.DrawLine($linePen, 68, 92, 144, 92)
    $graphics.DrawLine($linePen, 68, 112, 144, 112)

    $barBrushA = New-Object System.Drawing.SolidBrush((New-Color -Hex "#0F766E"))
    $barBrushB = New-Object System.Drawing.SolidBrush((New-Color -Hex "#22C55E"))
    $barBrushC = New-Object System.Drawing.SolidBrush((New-Color -Hex "#F59E0B"))
    $graphics.FillRectangle($barBrushA, 70, 134, 14, 22)
    $graphics.FillRectangle($barBrushB, 92, 124, 14, 32)
    $graphics.FillRectangle($barBrushC, 114, 116, 14, 40)

    $checkCircleBrush = New-Object System.Drawing.SolidBrush((New-Color -Hex "#12A57C"))
    $graphics.FillEllipse($checkCircleBrush, 60, 158, 40, 40)
    $checkPen = New-Object System.Drawing.Pen((New-Color -Hex "#FFFFFF"), 6.5)
    $checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $checkPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    $graphics.DrawLines($checkPen, @(
        (New-Object System.Drawing.Point(71, 178)),
        (New-Object System.Drawing.Point(79, 186)),
        (New-Object System.Drawing.Point(91, 170))
    ))

    $padShadowPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $padShadowPath.AddClosedCurve(@(
        (New-Object System.Drawing.PointF(108, 171)),
        (New-Object System.Drawing.PointF(98, 189)),
        (New-Object System.Drawing.PointF(95, 212)),
        (New-Object System.Drawing.PointF(109, 227)),
        (New-Object System.Drawing.PointF(132, 221)),
        (New-Object System.Drawing.PointF(146, 206)),
        (New-Object System.Drawing.PointF(179, 206)),
        (New-Object System.Drawing.PointF(196, 221)),
        (New-Object System.Drawing.PointF(217, 227)),
        (New-Object System.Drawing.PointF(231, 212)),
        (New-Object System.Drawing.PointF(228, 189)),
        (New-Object System.Drawing.PointF(218, 171)),
        (New-Object System.Drawing.PointF(196, 160)),
        (New-Object System.Drawing.PointF(130, 160))
    ), 0.42)
    $padShadowBrush = New-Object System.Drawing.SolidBrush((New-Color -Hex "#071219" -Alpha 48))
    $graphics.FillPath($padShadowBrush, $padShadowPath)

    $padPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $padPath.AddClosedCurve(@(
        (New-Object System.Drawing.PointF(102, 165)),
        (New-Object System.Drawing.PointF(92, 183)),
        (New-Object System.Drawing.PointF(89, 206)),
        (New-Object System.Drawing.PointF(103, 221)),
        (New-Object System.Drawing.PointF(126, 215)),
        (New-Object System.Drawing.PointF(140, 200)),
        (New-Object System.Drawing.PointF(173, 200)),
        (New-Object System.Drawing.PointF(190, 215)),
        (New-Object System.Drawing.PointF(211, 221)),
        (New-Object System.Drawing.PointF(225, 206)),
        (New-Object System.Drawing.PointF(222, 183)),
        (New-Object System.Drawing.PointF(212, 165)),
        (New-Object System.Drawing.PointF(190, 154)),
        (New-Object System.Drawing.PointF(124, 154))
    ), 0.42)
    $padBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Point(104, 154)),
        (New-Object System.Drawing.Point(220, 220)),
        (New-Color -Hex "#F59E0B"),
        (New-Color -Hex "#F97316")
    )
    $padBorderPen = New-Object System.Drawing.Pen((New-Color -Hex "#FFD188"), 2.0)
    $graphics.FillPath($padBrush, $padPath)
    $graphics.DrawPath($padBorderPen, $padPath)

    $padHighlightPen = New-Object System.Drawing.Pen((New-Color -Hex "#FFE6B0" -Alpha 160), 3.0)
    $padHighlightPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $padHighlightPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $graphics.DrawLine($padHighlightPen, 124, 163, 189, 163)

    $darkBrush = New-Object System.Drawing.SolidBrush((New-Color -Hex "#153142"))
    $graphics.FillRectangle($darkBrush, 124, 179, 10, 28)
    $graphics.FillRectangle($darkBrush, 115, 188, 28, 10)

    $buttonBlue = New-Object System.Drawing.SolidBrush((New-Color -Hex "#0F172A" -Alpha 150))
    $buttonCyan = New-Object System.Drawing.SolidBrush((New-Color -Hex "#DFFCF5"))
    $graphics.FillEllipse($buttonBlue, 182, 180, 14, 14)
    $graphics.FillEllipse($buttonBlue, 198, 168, 14, 14)
    $graphics.FillEllipse($buttonCyan, 185, 183, 8, 8)
    $graphics.FillEllipse($buttonCyan, 201, 171, 8, 8)

    $sparkBrush = New-Object System.Drawing.SolidBrush((New-Color -Hex "#FFFFFF" -Alpha 220))
    $graphics.FillEllipse($sparkBrush, 184, 54, 10, 10)

    $directory = Split-Path -Path $PreviewPath -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }

    $bitmap.Save($PreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)

    $sizes = @(16, 24, 32, 48, 64, 128, 256)
    $imageBytesList = New-Object 'System.Collections.Generic.List[byte[]]'

    foreach ($size in $sizes) {
        $resized = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g2 = [System.Drawing.Graphics]::FromImage($resized)
        try {
            $g2.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $g2.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g2.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $g2.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $g2.Clear([System.Drawing.Color]::Transparent)
            $g2.DrawImage($bitmap, 0, 0, $size, $size)

            $memory = New-Object System.IO.MemoryStream
            try {
                $resized.Save($memory, [System.Drawing.Imaging.ImageFormat]::Png)
                $imageBytesList.Add($memory.ToArray())
            }
            finally {
                $memory.Dispose()
            }
        }
        finally {
            $g2.Dispose()
            $resized.Dispose()
        }
    }

    Save-PngIconSetAsIco -ImageBytesList $imageBytesList -Sizes $sizes -OutputPath $IcoPath
}
finally {
    $graphics.Dispose()
    $bitmap.Dispose()
}

Write-Output ("Generated icon: {0}" -f $IcoPath)
Write-Output ("Generated preview: {0}" -f $PreviewPath)
