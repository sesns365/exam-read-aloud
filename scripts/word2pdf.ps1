# 用 Microsoft Word 把 .docx/.doc/.rtf/.odt 轉成 PDF（版面與 Word 完全相同）。
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File word2pdf.ps1 <輸入檔> <輸出.pdf>
param([string]$InPath, [string]$OutPdf)
$ErrorActionPreference = 'Stop'
$in = (Resolve-Path -LiteralPath $InPath).Path
$out = [System.IO.Path]::GetFullPath($OutPdf)
try {
  $word = New-Object -ComObject Word.Application
} catch {
  $soffice = @('C:\Program Files\LibreOffice\program\soffice.exe', 'C:\Program Files (x86)\LibreOffice\program\soffice.exe') | Where-Object { Test-Path $_ } | Select-Object -First 1
  if (-not $soffice) { Write-Error '找不到 Microsoft Word 或 LibreOffice，無法把 Word 轉成 PDF'; exit 1 }
  $dir = Split-Path $out
  & $soffice --headless --convert-to pdf --outdir $dir $in | Out-Null
  $gen = Join-Path $dir ([System.IO.Path]::GetFileNameWithoutExtension($in) + '.pdf')
  if ($gen -ne $out) { Move-Item -Force $gen $out }
  exit 0
}
try {
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($in, $false, $true)   # ConfirmConversions=false, ReadOnly=true
  $doc.ExportAsFixedFormat($out, 17)                 # 17 = wdExportFormatPDF
  $doc.Close($false)
} finally {
  # 只關閉這次開啟的 Word；若使用者原本就開著其他文件，不要關掉
  if ($word.Documents.Count -eq 0) { $word.Quit() }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
if (-not (Test-Path $out)) { Write-Error "Word 轉 PDF 失敗：$out"; exit 1 }
