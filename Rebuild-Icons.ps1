Add-Type -AssemblyName System.Drawing

$assetsDir = Join-Path $PSScriptRoot "src\MagicVoice\Assets"
$sourcePath = Join-Path $assetsDir "App310.png"

# 1) Bounding box of opaque pixels in source
$src = New-Object System.Drawing.Bitmap($sourcePath)
$minX = $src.Width; $minY = $src.Height; $maxX = -1; $maxY = -1
for ($y = 0; $y -lt $src.Height; $y++) {
    for ($x = 0; $x -lt $src.Width; $x++) {
        $a = $src.GetPixel($x, $y).A
        if ($a -gt 8) {
            if ($x -lt $minX) { $minX = $x }
            if ($x -gt $maxX) { $maxX = $x }
            if ($y -lt $minY) { $minY = $y }
            if ($y -gt $maxY) { $maxY = $y }
        }
    }
}
$cropW = $maxX - $minX + 1
$cropH = $maxY - $minY + 1
$side = [Math]::Max($cropW, $cropH)
# Center on a square canvas (in case content isn't perfectly square)
$cx = ($minX + $maxX) / 2
$cy = ($minY + $maxY) / 2
$sx = [int][Math]::Round($cx - $side / 2)
$sy = [int][Math]::Round($cy - $side / 2)

$cropRect = New-Object System.Drawing.Rectangle($sx, $sy, $side, $side)
$crop = New-Object System.Drawing.Bitmap($side, $side, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$gC = [System.Drawing.Graphics]::FromImage($crop)
$gC.Clear([System.Drawing.Color]::Transparent)
$destRect = New-Object System.Drawing.Rectangle(0, 0, $side, $side)
$gC.DrawImage($src, $destRect, $cropRect, [System.Drawing.GraphicsUnit]::Pixel)
$gC.Dispose()
$src.Dispose()
Write-Output ("Cropped to {0}x{1} from origin ({2},{3})" -f $side, $side, $sx, $sy)

function Resize-Image($source, $size) {
    $bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.DrawImage($source, (New-Object System.Drawing.Rectangle(0, 0, $size, $size)))
    $g.Dispose()
    return $bmp
}

function Backup-File($path) {
    $bak = "$path.bak"
    if (-not (Test-Path $bak)) {
        Copy-Item -Path $path -Destination $bak
        Write-Output ("Backed up: {0}" -f $bak)
    }
}

# 2) Re-render PNG tiles
foreach ($pair in @(@(44, "App44.png"), @(150, "App150.png"), @(310, "App310.png"))) {
    $size = $pair[0]
    $name = $pair[1]
    $path = Join-Path $assetsDir $name
    Backup-File $path
    $resized = Resize-Image $crop $size
    $resized.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $resized.Dispose()
    Write-Output ("Wrote: {0} ({1}x{1})" -f $path, $size)
}

# 3) Build App.ico with frames 16, 24, 32, 48, 64, 128, 256 from cropped bitmap
$icoPath = Join-Path $assetsDir "App.ico"
Backup-File $icoPath

$frameSizes = 16, 24, 32, 48, 64, 128, 256
$frameBitmaps = @()
$framePngBytes = @()
foreach ($s in $frameSizes) {
    $b = Resize-Image $crop $s
    $ms = New-Object System.IO.MemoryStream
    $b.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $frameBitmaps += $b
    $framePngBytes += ,($ms.ToArray())
    $ms.Dispose()
}

$out = New-Object System.IO.MemoryStream
$bw = New-Object System.IO.BinaryWriter($out)
# ICONDIR
$bw.Write([UInt16]0)        # reserved
$bw.Write([UInt16]1)        # type = icon
$bw.Write([UInt16]$frameSizes.Count)

$dataOffset = 6 + 16 * $frameSizes.Count
for ($i = 0; $i -lt $frameSizes.Count; $i++) {
    $s = $frameSizes[$i]
    $bytes = $framePngBytes[$i]
    $wByte = if ($s -ge 256) { 0 } else { [byte]$s }
    $hByte = $wByte
    $bw.Write([byte]$wByte)         # width
    $bw.Write([byte]$hByte)         # height
    $bw.Write([byte]0)              # color count
    $bw.Write([byte]0)              # reserved
    $bw.Write([UInt16]1)            # color planes
    $bw.Write([UInt16]32)           # bits per pixel
    $bw.Write([UInt32]$bytes.Length)
    $bw.Write([UInt32]$dataOffset)
    $dataOffset += $bytes.Length
}
foreach ($bytes in $framePngBytes) { $bw.Write($bytes) }
$bw.Flush()
[System.IO.File]::WriteAllBytes($icoPath, $out.ToArray())
$bw.Dispose()
$out.Dispose()
foreach ($b in $frameBitmaps) { $b.Dispose() }
Write-Output ("Wrote: {0} ({1} frames)" -f $icoPath, $frameSizes.Count)

$crop.Dispose()
Write-Output "Done."
