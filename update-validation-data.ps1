param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot '驗測資料.xlsx'),
  [string]$DataPath = (Join-Path $PSScriptRoot 'validation-review-data-new.js')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipXml([IO.Compression.ZipArchive]$Zip, [string]$Name) {
  $entry = $Zip.GetEntry($Name)
  if (-not $entry) { throw "找不到 Excel 內容：$Name" }
  $reader = [IO.StreamReader]::new($entry.Open())
  try { return [xml]$reader.ReadToEnd() } finally { $reader.Close() }
}

function Get-SharedStrings([xml]$Xml) {
  $result = @()
  foreach ($si in $Xml.SelectNodes("//*[local-name()='si']")) {
    $parts = @()
    foreach ($t in $si.SelectNodes(".//*[local-name()='t']")) { $parts += $t.InnerText }
    $result += ($parts -join '')
  }
  return $result
}

function Convert-PythonDictToJson([string]$Text) {
  # Excel 中部分模組輸出為 Python dict repr，轉成可解析的 JSON。
  return ($Text -replace "'", '"' -replace '\bFalse\b', 'false' -replace '\bTrue\b', 'true' -replace '\bNone\b', 'null')
}

function Read-ModuleJsons([string]$Kind) {
  $zip = [IO.Compression.ZipFile]::OpenRead($ExcelPath)
  try {
    $strings = Get-SharedStrings (Read-ZipXml $zip 'xl/sharedStrings.xml')
    if ($Kind -eq '條件') {
      return @($strings | Where-Object { $_.Contains('EXPOSURE_COMPARISON') })
    }
    if ($Kind -eq '標的') {
      return @($strings | Where-Object { $_.Contains('REVENUE_EXPOSURE_SCORE') -and $_.Contains('SUBJECT_RISK_REPORT_TABLE') })
    }
    if ($Kind -eq '歷史') {
      return @($strings | Where-Object { $_.Contains('HISTORY_RESULT') -and $_.Contains('RECOMMENDED_REVIEW_CASES') })
    }
  } finally { $zip.Dispose() }
}

if (-not (Test-Path -LiteralPath $ExcelPath)) { throw "找不到 Excel：$ExcelPath" }
if (-not (Test-Path -LiteralPath $DataPath)) { throw "找不到資料檔：$DataPath" }

$raw = Get-Content -LiteralPath $DataPath -Raw -Encoding UTF8
$start = $raw.IndexOf('{')
$end = $raw.LastIndexOf('}')
if ($start -lt 0 -or $end -le $start) { throw 'validation-review-data-new.js 不是有效的資料格式' }
$data = $raw.Substring($start, $end - $start + 1) | ConvertFrom-Json

foreach ($kind in @('條件', '標的', '歷史')) {
  $modules = @()
  foreach ($case in $data.cases) {
    foreach ($module in $case.modules) {
      if ($module.name -eq $kind) { $modules += $module }
    }
  }
  $source = @(Read-ModuleJsons $kind)
  if ($source.Count -ne $modules.Count) {
    throw "$kind：Excel JSON $($source.Count) 筆，資料檔模組 $($modules.Count) 筆，數量不一致，已停止更新。"
  }
  for ($i = 0; $i -lt $modules.Count; $i++) {
    $text = Convert-PythonDictToJson $source[$i]
    $item = $text | ConvertFrom-Json
    foreach ($property in ($item.PSObject.Properties.Name)) {
      $modules[$i] | Add-Member -NotePropertyName $property -NotePropertyValue $item.$property -Force
    }
  }
  Write-Host "$kind：已更新 $($modules.Count) 個版次" -ForegroundColor Green
}

$output = 'window.validationReviewData = ' + ($data | ConvertTo-Json -Depth 40) + ';'
[IO.File]::WriteAllText($DataPath, $output, [Text.UTF8Encoding]::new($false))
Write-Host "完成：$DataPath" -ForegroundColor Cyan
