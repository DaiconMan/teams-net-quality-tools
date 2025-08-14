<# 
Generate-TeamsNet-Report.ps1
- Excel(COM)で CSV を取り込み、ピボット＆グラフを自動生成
- 既定の入力: %LOCALAPPDATA%\TeamsNet\teams_net_quality.csv / path_hop_quality.csv
- 出力: -Output で指定した .xlsx

保存推奨: UTF-8 (BOM) + CRLF
#>

param(
  [string]$InputDir = (Join-Path $env:LOCALAPPDATA "TeamsNet"),
  [Parameter(Mandatory=$true)][string]$Output,
  [switch]$Visible
)

# ===== Excel 定数 =====
# グラフ種別やピボット定数（バージョン差吸収のため数値で定義）
$xlDelimited           = 1
$xlDatabase            = 1
$xlYes                 = 1
$xlRowField            = 1
$xlColumnField         = 2
$xlPageField           = 3
$xlDataField           = 4
$xlAverage             = -4106
$xlLine                = 4
$xlColumnClustered     = 51
$xlLegendPositionBottom= -4107

# ===== 準備 =====
$ErrorActionPreference = "Stop"
$teamsCsv = Join-Path $InputDir "teams_net_quality.csv"
$hopsCsv  = Join-Path $InputDir "path_hop_quality.csv"

if(-not (Test-Path $teamsCsv)){
  throw "CSV が見つかりません: $teamsCsv"
}

# 親フォルダを作っておく
$outDir = Split-Path -Parent $Output
if($outDir -and -not (Test-Path $outDir)){ New-Item -ItemType Directory -Path $outDir | Out-Null }

# ===== Excel 起動 =====
try {
  $excel = New-Object -ComObject Excel.Application
} catch {
  throw "Excel の COM を起動できません。Microsoft Excel がインストールされている必要があります。"
}
$excel.Visible = [bool]$Visible
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

# 既存シート整理（Book1にデフォルト3枚ある想定）
while($wb.Worksheets.Count -gt 0){ $wb.Worksheets.Item(1).Delete() }

function New-Worksheet([string]$name) {
  $ws = $wb.Worksheets.Add()
  $ws.Name = $name
  return $ws
}

function Import-CsvToTable([string]$csvPath,[string]$sheetName,[string]$tableName){
  $ws = New-Worksheet $sheetName
  # QueryTable で UTF-8 として取り込み
  $qt = $ws.QueryTables.Add("TEXT;" + $csvPath, $ws.Range("A1"))
  $qt.TextFileParseType = $xlDelimited
  $qt.TextFileCommaDelimiter = $true
  $qt.TextFilePlatform = 65001   # UTF-8
  $qt.TextFileTrailingMinusNumbers = $true
  $qt.AdjustColumnWidth = $true
  $qt.Refresh() | Out-Null

  # テーブル化
  $used = $ws.UsedRange
  $lo = $ws.ListObjects.Add(1, $used, $null, $xlYes)  # xlSrcRange=1
  $lo.Name = $tableName
  $ws.Columns.AutoFit() | Out-Null
  return $lo
}

function Add-ValueField([object]$pt,[string]$fieldName,[string]$caption){
  try {
    $pf = $pt.PivotFields($fieldName)
    $df = $pt.AddDataField($pf, $caption, $xlAverage)
    $df.NumberFormat = "0.00"
  } catch {
    # フィールドが無い場合は無視
  }
}

function Add-Slicer([object]$pt,[string]$fieldName,[object]$ws,[double]$x,[double]$y){
  try {
    $sc = $wb.SlicerCaches.Add2($pt, $fieldName, $fieldName)
    $null = $sc.Slicers.Add($ws, $null, $fieldName, $fieldName, $x, $y, 180, 110)
  } catch {
    # 古いExcelやフィールド未存在でも無視して続行
  }
}

# ===== 取り込み =====
$loEnd = Import-CsvToTable -csvPath $teamsCsv -sheetName "EndToEnd" -tableName "tblEndToEnd"
if(Test-Path $hopsCsv){
  $loHop = Import-CsvToTable -csvPath $hopsCsv  -sheetName "HopsRaw"  -tableName "tblHopsRaw"
}

# ===== ピボット（EndToEnd） =====
$wsP1 = New-Worksheet "Pivot_EndToEnd"
$pc1 = $excel.PivotCaches().Create($xlDatabase, $loEnd.Range)
$pt1 = $pc1.CreatePivotTable($wsP1.Range("A3"), "pvtEnd")

# 行・ページ
$pfTime = $pt1.PivotFields("timestamp"); $pfTime.Orientation = $xlRowField; $pfTime.NumberFormat = "yyyy-mm-dd hh:mm:ss"
try { $pt1.PivotFields("host").Orientation = $xlPageField } catch {}
try { $pt1.PivotFields("conn_type").Orientation = $xlPageField } catch {}

# 値
Add-ValueField $pt1 "icmp_avg_ms"     "平均RTT(ms)"
Add-ValueField $pt1 "icmp_jitter_ms"  "平均ジッタ(ms)"
Add-ValueField $pt1 "loss_pct"        "平均ロス(%)"
Add-ValueField $pt1 "mos_estimate"    "平均MOS"

# グラフ
# 既存のチャートがないので新規作成
$ch1 = $wsP1.ChartObjects().Add(400, 20, 900, 380)
$ch1.Name = "chtEnd"
$ch1.Chart.SetSourceData($pt1.TableRange1)
$ch1.Chart.ChartType = $xlLine
$ch1.Chart.HasTitle = $true
$ch1.Chart.ChartTitle.Text = "End-to-End: RTT/Jitter/Loss/MOS（hostで絞り込み可）"
$ch1.Chart.Legend.Position = $xlLegendPositionBottom

# スライサー
Add-Slicer $pt1 "host" $wsP1 20 20

# ===== ピボット（Hops） =====
if($loHop){
  $wsP2 = New-Worksheet "Pivot_Hops"
  $pc2 = $excel.PivotCaches().Create($xlDatabase, $loHop.Range)
  $pt2 = $pc2.CreatePivotTable($wsP2.Range("A3"), "pvtHops")

  # ページ/行
  try { $pt2.PivotFields("target").Orientation = $xlPageField } catch {}
  try { $pt2.PivotFields("timestamp").Orientation = $xlPageField } catch {}
  try { $pt2.PivotFields("hop_index").Orientation = $xlRowField } catch {}

  # 値
  Add-ValueField $pt2 "icmp_jitter_ms" "平均ジッタ(ms)"
  Add-ValueField $pt2 "loss_pct"       "平均ロス(%)"

  # 最新タイムスタンプを選択（可能なら）
  try {
    $pfTs = $pt2.PivotFields("timestamp")
    $last = $null
    foreach($pi in $pfTs.PivotItems()){ $last = $pi }
    if($last){ $pfTs.ClearAllFilters(); $last.Visible = $true }
  } catch {}

  # グラフ
  $ch2 = $wsP2.ChartObjects().Add(400, 20, 900, 380)
  $ch2.Name = "chtHops"
  $ch2.Chart.SetSourceData($pt2.TableRange1)
  $ch2.Chart.ChartType = $xlColumnClustered
  $ch2.Chart.HasTitle = $true
  $ch2.Chart.ChartTitle.Text = "Hops: 平均ジッタ/ロス（target・timestampで絞り込み可）"
  $ch2.Chart.Legend.Position = $xlLegendPositionBottom

  # スライサー
  Add-Slicer $pt2 "target"    $wsP2 20 20
  Add-Slicer $pt2 "timestamp" $wsP2 20 140
}

# ===== 体裁 & 保存 =====
foreach($ws in @("EndToEnd","HopsRaw","Pivot_EndToEnd","Pivot_Hops")){
  if(($wb.Worksheets | % Name) -contains $ws){
    $wb.Worksheets($ws).Activate() | Out-Null
    try { $wb.ActiveSheet.Columns.AutoFit() | Out-Null } catch {}
  }
}

$wb.SaveAs($Output)
$wb.Close($true)
$excel.Quit()

# COM解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)   | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)| Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host "レポートを出力しました: $Output"