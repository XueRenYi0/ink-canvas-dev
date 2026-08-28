# 构建多尺寸 ICO：从 1024px 主图生成 16~256 六档 PNG 并打包
# 用法: powershell -File build_icon.ps1
Add-Type -AssemblyName System.Drawing

$src = "C:\Users\13555\AppData\Roaming\TRAE SOLO CN\ModularData\ai-agent\work-mode-projects\6a90fb57733fdc0cffa67dcd\Ink-Canvas-Dev\Resources\new_icon_1024.jpg"
$outIco = "C:\Users\13555\AppData\Roaming\TRAE SOLO CN\ModularData\ai-agent\work-mode-projects\6a90fb57733fdc0cffa67dcd\Ink-Canvas-Dev\Ink Canvas\Resources\InkCanvas.ico"

$img = [System.Drawing.Image]::FromFile($src)
Write-Output ("source: " + $img.Width + "x" + $img.Height)

# 中心裁方（保险：非方形时裁成正方形）
$side = [Math]::Min($img.Width, $img.Height)
$x0 = [int](($img.Width - $side) / 2)
$y0 = [int](($img.Height - $side) / 2)

$sizes = @(256, 128, 64, 48, 32, 16)
$pngStreams = @()

foreach ($s in $sizes) {
    $bmp = New-Object System.Drawing.Bitmap($s, $s)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    # 高质量缩小：DrawImage 带源矩形
    $g.DrawImage($img, (New-Object System.Drawing.Rectangle(0, 0, $s, $s)), (New-Object System.Drawing.Rectangle($x0, $y0, $side, $side)), [System.Drawing.GraphicsUnit]::Pixel)
    $g.Dispose()

    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $pngStreams += ,($ms.ToArray())
    Write-Output ("  png " + $s + "x" + $s + ": " + $ms.Length + " bytes")
}
$img.Dispose()

# 构造 ICO：ICONDIR + N 个 ICONDIRENTRY + PNG 数据
$count = $sizes.Count
$header = New-Object System.IO.MemoryStream
$bw = New-Object System.IO.BinaryWriter($header)
$bw.Write([UInt16]0)        # reserved
$bw.Write([UInt16]1)        # type = icon
$bw.Write([UInt16]$count)   # 图像数

$offset = 6 + 16 * $count   # 头部之后的第一个图像偏移
for ($i = 0; $i -lt $count; $i++) {
    $w = if ($sizes[$i] -ge 256) { 0 } else { $sizes[$i] }
    $bw.Write([byte]$w)             # width (256 用 0)
    $bw.Write([byte]$w)             # height
    $bw.Write([byte]0)              # palette
    $bw.Write([byte]0)              # reserved
    $bw.Write([UInt16]1)            # planes
    $bw.Write([UInt16]32)           # bpp
    $bw.Write([UInt32]$pngStreams[$i].Length)
    $bw.Write([UInt32]$offset)
    $offset += $pngStreams[$i].Length
}

$fs = [System.IO.File]::Create($outIco)
$bw.Flush()
$header.WriteTo($fs)
foreach ($png in $pngStreams) { $fs.Write($png, 0, $png.Length) }
$fs.Close()
$bw.Close()

$fi = Get-Item $outIco
Write-Output ("ICO written: " + $fi.Length + " bytes -> " + $outIco)
