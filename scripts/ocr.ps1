# 用 Windows 內建 OCR（zh-Hant-TW）辨識圖片，輸出每個字詞的位置（JSON，UTF-8）。
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File ocr.ps1 <圖片路徑> <輸出.json> [語言, 預設 zh-Hant-TW]
param([string]$ImagePath, [string]$OutJson, [string]$Lang = 'zh-Hant-TW')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]

$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
function Await($op, [Type]$type) { $t = $asTask.MakeGenericMethod($type).Invoke($null, @($op)); $t.Wait(); $t.Result }

$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync((Resolve-Path -LiteralPath $ImagePath).Path)) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new($Lang))
if ($null -eq $engine) { Write-Error "這台電腦沒有 $Lang 的 OCR 語言套件（Windows 設定 → 時間與語言 → 語言，新增「中文(繁體，台灣)」）"; exit 1 }
$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

$lines = foreach ($ln in $result.Lines) {
  ,@($ln.Words | ForEach-Object { @{ t = $_.Text; x = $_.BoundingRect.X; y = $_.BoundingRect.Y; w = $_.BoundingRect.Width; h = $_.BoundingRect.Height } })
}
$out = @{ width = $bitmap.PixelWidth; height = $bitmap.PixelHeight; lines = @($lines) }
[System.IO.File]::WriteAllText($OutJson, ($out | ConvertTo-Json -Depth 6 -Compress), (New-Object System.Text.UTF8Encoding $false))
$stream.Dispose()
