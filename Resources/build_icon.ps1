# 构建多尺寸 ICO：1024 主图 → 圆角 + 白底透明度处理 → 多档 PNG 打包
# 用法:
#   对比预览:  powershell -File build_icon.ps1 -PreviewOnly
#   正式生成:  powershell -File build_icon.ps1 -TileAlpha 230            (半透白方案)
#              powershell -File build_icon.ps1 -TileGray 240              (浅灰方案)
#              powershell -File build_icon.ps1 -TileAlpha 255             (纯白方案，仅圆角)
# 参数:
#   -TileAlpha     白底不透明度 0-255（255=纯白不透明；越低越磨砂，黑笔始终不透明）
#   -TileGray      非零时改用浅灰不透明底，值为灰度 0-255（如 240=#F0F0F0），优先于 TileAlpha
#   -RadiusRatio   圆角半径比例（相对内容边长，0.18 ≈ 现代应用图标观感）
#   -ContentScale  内容安全区缩放（0.94 = 四周留 3% 边距，防止笔尖被圆角裁切）
param(
    [int]$TileAlpha = 230,
    [int]$TileGray = 0,
    [double]$RadiusRatio = 0.18,
    [double]$ContentScale = 0.94,
    [switch]$PreviewOnly
)
Add-Type -AssemblyName System.Drawing

# 像素级处理用内联 C#（PowerShell 逐像素太慢）
Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

public static class IconFx
{
    // 圆角矩形 SDF 覆盖率（1px 抗锯齿边缘）
    static double Cov(double x, double y, double cx, double cy, double half, double r)
    {
        double qx = Math.Abs(x - cx) - (half - r);
        double qy = Math.Abs(y - cy) - (half - r);
        double dx = Math.Max(qx, 0.0), dy = Math.Max(qy, 0.0);
        double d = Math.Sqrt(dx * dx + dy * dy) + Math.Min(Math.Max(qx, qy), 0.0) - r;
        double cov = 0.5 - d;
        if (cov < 0) cov = 0;
        if (cov > 1) cov = 1;
        return cov;
    }

    /// 处理单帧：亮度决定 alpha（黑笔=255 不透明，白底=tileAlpha）；
    /// tileGray>0 时亮像素改为不透明浅灰底。圆角外全透明。
    public static Bitmap Process(Bitmap src, int outSize, int tileAlpha, int tileGray, double radiusRatio, double contentScale)
    {
        Bitmap bmp = new Bitmap(outSize, outSize, PixelFormat.Format32bppArgb);
        int c = (int)Math.Round(outSize * contentScale);
        if (c > outSize) c = outSize;
        int off = (outSize - c) / 2;
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.DrawImage(src, new Rectangle(off, off, c, c));
        }

        BitmapData bd = bmp.LockBits(new Rectangle(0, 0, outSize, outSize),
            ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        int stride = bd.Stride;
        byte[] px = new byte[stride * outSize];
        Marshal.Copy(bd.Scan0, px, 0, px.Length);

        double half = c / 2.0;
        double r = radiusRatio * c;
        double cx = off + half, cy = off + half;

        for (int y = 0; y < outSize; y++)
        {
            for (int x = 0; x < outSize; x++)
            {
                int i = y * stride + x * 4;
                byte b = px[i], gg = px[i + 1], rr = px[i + 2];
                double L = (rr * 299.0 + gg * 587.0 + b * 114.0) / 1000.0;
                double a;
                if (tileGray > 0 && L >= 200)
                {
                    // 浅灰方案：亮背景像素统一改为不透明灰
                    px[i] = (byte)tileGray;
                    px[i + 1] = (byte)tileGray;
                    px[i + 2] = (byte)tileGray;
                    a = 255.0;
                }
                else
                {
                    // 透明度方案：黑笔保持 255，白底降到 tileAlpha，抗锯齿边缘平滑过渡
                    a = tileAlpha + (255.0 - L) * (255.0 - tileAlpha) / 255.0;
                    if (a > 255.0) a = 255.0;
                }
                a *= Cov(x + 0.5, y + 0.5, cx, cy, half, r);
                px[i + 3] = (byte)Math.Round(a);
            }
        }
        Marshal.Copy(px, 0, bd.Scan0, px.Length);
        bmp.UnlockBits(bd);
        return bmp;
    }
}
"@ -ReferencedAssemblies System.Drawing

$repo = Split-Path -Parent $PSScriptRoot
$src = Join-Path $PSScriptRoot "new_icon_1024.jpg"
$outIco = Join-Path $repo "Ink Canvas\Resources\InkCanvas.ico"

# 载入源图并中心裁方
$img = [System.Drawing.Image]::FromFile($src)
Write-Output ("source: " + $img.Width + "x" + $img.Height)
$side = [Math]::Min($img.Width, $img.Height)
$x0 = [int](($img.Width - $side) / 2)
$y0 = [int](($img.Height - $side) / 2)
$square = New-Object System.Drawing.Bitmap($side, $side)
$g = [System.Drawing.Graphics]::FromImage($square)
$g.DrawImage($img, (New-Object System.Drawing.Rectangle(0, 0, $side, $side)),
    (New-Object System.Drawing.Rectangle($x0, $y0, $side, $side)), [System.Drawing.GraphicsUnit]::Pixel)
$g.Dispose()
$img.Dispose()

if ($PreviewOnly) {
    # ===== 对比预览：上行深色底、下行浅色底，模拟任务栏环境 =====
    $variants = @(
        @{ Name = "A 纯白不透明"; Alpha = 255; Gray = 0 },
        @{ Name = "B 半透白 90%"; Alpha = 230; Gray = 0 },
        @{ Name = "C 浅灰 F0F0F0"; Alpha = 255; Gray = 240 }
    )
    $iconSize = 64
    $W = 352; $H = 300
    $prev = New-Object System.Drawing.Bitmap($W, $H)
    $g = [System.Drawing.Graphics]::FromImage($prev)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $fontTitle = New-Object System.Drawing.Font("Microsoft YaHei UI", 11, [System.Drawing.FontStyle]::Bold)
    $fontLabel = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)

    foreach ($row in @(0, 1)) {
        $dark = ($row -eq 0)
        # 行背景
        $bg = if ($dark) { [System.Drawing.Color]::FromArgb(31, 31, 31) } else { [System.Drawing.Color]::FromArgb(239, 239, 239) }
        $g.FillRectangle((New-Object System.Drawing.SolidBrush($bg)), 0, $row * 150, $W, 150)
        $fg = if ($dark) { [System.Drawing.Brushes]::White } else { [System.Drawing.Brushes]::Black }
        # 行标题
        $title = if ($dark) { "深色背景（模拟深色任务栏）" } else { "浅色背景（模拟浅色任务栏）" }
        $g.DrawString($title, $fontTitle, $fg, 12, 8 + $row * 150)
        # 三方案图标 + 标签
        for ($v = 0; $v -lt 3; $v++) {
            $icon = [IconFx]::Process($square, 256, $variants[$v].Alpha, $variants[$v].Gray, $RadiusRatio, $ContentScale)
            $ix = 24 + $v * 110
            $iy = 40 + $row * 150
            $g.DrawImage($icon, $ix, $iy, $iconSize, $iconSize)
            $icon.Dispose()
            $g.DrawString($variants[$v].Name, $fontLabel, $fg, $ix, $iy + $iconSize + 6)
        }
    }
    $g.Dispose()
    $prev.Save((Join-Path $PSScriptRoot "icon-preview.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $prev.Dispose()
    Write-Output ("preview written: " + (Join-Path $PSScriptRoot "icon-preview.png"))
}
else {
    # ===== 正式生成：1024 处理一次 → 六档缩小 → 打包 ICO =====
    $processed = [IconFx]::Process($square, 1024, $TileAlpha, $TileGray, $RadiusRatio, $ContentScale)
    $square.Dispose()

    $sizes = @(256, 128, 64, 48, 32, 16)
    $pngStreams = @()
    foreach ($s in $sizes) {
        $bmp = New-Object System.Drawing.Bitmap($s, $s)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.DrawImage($processed, (New-Object System.Drawing.Rectangle(0, 0, $s, $s)))
        $g.Dispose()
        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        $pngStreams += ,($ms.ToArray())
        Write-Output ("  png " + $s + "x" + $s + ": " + $ms.Length + " bytes")
    }
    $processed.Dispose()

    # 构造 ICO：ICONDIR + N 个 ICONDIRENTRY + PNG 数据
    $count = $sizes.Count
    $header = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($header)
    $bw.Write([UInt16]0)        # reserved
    $bw.Write([UInt16]1)        # type = icon
    $bw.Write([UInt16]$count)   # 图像数

    $offset = 6 + 16 * $count
    for ($i = 0; $i -lt $count; $i++) {
        $w = if ($sizes[$i] -ge 256) { 0 } else { $sizes[$i] }
        $bw.Write([byte]$w)
        $bw.Write([byte]$w)
        $bw.Write([byte]0)
        $bw.Write([byte]0)
        $bw.Write([UInt16]1)
        $bw.Write([UInt16]32)
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
    Write-Output "注意: exe 图标是编译时嵌入的，需重新编译程序才能生效"
}
