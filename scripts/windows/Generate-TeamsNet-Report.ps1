<# 
Generate-TeamsNet-Report.ps1  -- 完全版
- CSV（teams_net_quality.csv / path_hop_quality.csv）をExcelに取り込み
- ピボットテーブル＋スライサー＋グラフを自動生成
- クエリ/テーブル重なり、最後の1枚削除禁止 などのExcel仕様に対応済み

保存推奨: UTF-8 (BOM) + CRLF
#>

param(
  [string]$InputDir = (Join-Path $env:LOCALAPPDATA "TeamsNet"),
  [Parameter(Mandatory=$true)][string]$Output,
  [switch]$Visible
)

# ===== Excel 定数 =====
$xlDelimited            = 1
$xlDatabase             = 1
$xlYes                  = 1
$xlRowField             = 1
$xlColumnField          = 2
$xlPageField            = 3
$xlAverage              = -4106
$xlLine                 = 4
$xlColumnClustered      = 51
$xlLegendPositionBottom = -4107
$xlSrcRange             = 1
$xlInsertDeleteCells    = 2

# ===== 事前チェック =====
$ErrorActionPreference = "Stop"
$teamsCsv = Join-Path $InputDir "teams_net_quality.csv"
$hopsCsv  = Join-Path $InputDir "path_hop_quality.csv"

if(-not (Test-Path $teamsCsv)){
  throw "CSV が見つかりません: $teamsCsv"
}
if(-not (Test-Path $InputDir)){ throw "入力フォルダが存在しません: $InputDir" }

# 出力先フォルダ
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

# 既存シートの整理：最後の1枚は残す（Excelの仕様）
while($wb.Worksheets.Count -gt 1){ $wb.Worksheets.Item(1).Delete() }
$wb.Worksheets.Item(1).Name = "Scratch"

# ===== ヘルパー =====
function Get-Or-NewWorksheet([string]$name){
  # 既存なら徹底クリア、無ければ追加
  try {
    $ws = $wb.Worksheets.Item($name)
    # ピボット/テーブル/クエリ/チャート/内容をクリア
    try { foreach($pt in @($ws.PivotTables())){ $pt.TableRange2.Clear() } } catch {}
    try { foreach($lo in @($ws.ListObjects())){ $lo.Unlist() } } catch {}
    try { foreach($qt in @($ws.QueryTables())){ $qt.Delete() } } catch {}
    try { foreach($co in @($ws.ChartObjects())){ $co.Delete() } } catch {}
    $ws.Cells.Clear()
  } catch {
    $ws = $wb.Worksheets.Add()
    $ws.Name = $name
  }
  return $ws
}

# === CSV -> テーブル取り込み（クエリ重なり回避版） ===
function Import-CsvToTable([string]$csvPath,[string]$sheetName,[string]$tableName){
  $ws = Get-Or-NewWorksheet $sheetName

  # QueryTable で UTF-8 読み込み
  $qt = $ws.QueryTables.Add("TEXT;" + $csvPath, $ws.Range("A1"))
  $qt.TextFileParseType            = $xlDelimited
  $qt.TextFileCommaDelimiter       = $true
  $qt.TextFilePlatform             = 65001    # UTF-8
  $qt.TextFileTrailingMinusNumbers = $true
  $qt.AdjustColumnWidth            = $true
  $qt.RefreshStyle                 = $xlInsertDeleteCells
  $qt.Refresh() | Out-Null

  # 実データ範囲
  $rng = $qt.ResultRange
  if(-not $rng){ throw "CSV にデータが無いか、取り込みに失敗しました: $csvPath" }

  # クエリを削除（データは残す）→ その範囲でテーブル化
  $qt.Delete()
  $lo = $ws.ListObjects.Add($xlSrcRange, $rng, $null, $xlYes)
  $lo.Name = $tableName
  $ws.Columns.AutoFit() | Out-Null
  return $lo
}

function Add-ValueField([object]$pt,[string]$fieldName,[string]$caption){
  try {
    $pf = $pt.PivotFields($fieldName)
    $df = $pt.AddDataField($pf, $caption, $xlAverage)
    $df.NumberFormat = "0.00"
  } catch { } # フィールドが無い場合は無視
}

function Add-Slicer([object]$pt,[string]$fieldName,[object]$ws,[double]$x,[double]$y){
  try {
    # 古いExcelだと Add2 が無い場合あり → 失敗しても処理継続
    $sc = $wb.SlicerCaches.Add2($pt, $fieldName, $fieldName)
    $null = $sc.Slicers.Add($ws, $null, $fieldName, $fieldName, $x, $y, 180, 110)
  } catch { }
}

# ===== 取り込み =====
$loEnd = Import-CsvToTable -csvPath $teamsCsv -sheetName "EndToEnd" -tableName "tblEndToEnd"
$loHop = $null
if(Test-Path $hopsCsv){
  # 空ファイルのことがあるので Try
  try { $loHop = Import-CsvToTable -csvPath $hopsCsv -sheetName "HopsRaw" -tableName "tblHopsRaw" } catch { $loHop = $null }
}

# ===== ピボット：EndToEnd =====
$wsP1 = Get-Or-NewWorksheet "Pivot_EndToEnd"
$pc1 = $excel.PivotCaches().Create($xlDatabase, $loEnd.Range)
$pt1 = $pc1.CreatePivotTable($wsP1.Range("A3"), "pvtEnd")

# 行/ページフィールド
try { $pfTime = $pt1.PivotFields("timestamp"); $pfTime.Orientation = $xlRowField; $pfTime.NumberFormat = "yyyy-mm-dd hh:mm:ss" } catch {}
try { $pt1.PivotFields("host").Orientation = $xlPageField } catch {}
try { $pt1.PivotFields("conn_type").Orientation = $xlPageField } catch {}

# 値フィールド
Add-ValueField $pt1 "icmp_avg_ms"     "平均RTT(ms)"
Add-ValueField $pt1 "icmp_jitter_ms"  "平均ジッタ(ms)"
Add-ValueField $pt1 "loss_pct"        "平均ロス(%)"
Add-ValueField $pt1 "mos_estimate"    "平均MOS"

# グラフ
try {
  foreach($co in @($wsP1.ChartObjects())){ $co.Delete() }
} catch {}
$ch1 = $wsP1.ChartObjects().Add(400, 20, 900, 380)
$ch1.Name = "chtEnd"
$ch1.Chart.SetSourceData($pt1.TableRange1)
$ch1.Chart.ChartType = $xlLine
$ch1.Chart.HasTitle = $true
$ch1.Chart.ChartTitle.Text = "End-to-End: RTT/Jitter/Loss/MOS（hostで絞り込み可）"
$ch1.Chart.Legend.Position = $xlLegendPositionBottom

Add-Slicer $pt1 "host" $wsP1 20 20

# ===== ピボット：Hops =====
if($loHop){
  $wsP2 = Get-Or-NewWorksheet "Pivot_Hops"
  $pc2 = $excel.PivotCaches().Create($xlDatabase, $loHop.Range)
  $pt2 = $pc2.CreatePivotTable($wsP2.Range("A3"), "pvtHops")

  try { $pt2.PivotFields("target").Orientation    = $xlPageField } catch {}
  try { $pt2.PivotFields("timestamp").Orientation = $xlPageField } catch {}
  try { $pt2.PivotFields("hop_index").Orientation = $xlRowField  } catch {}

  Add-ValueField $pt2 "icmp_jitter_ms" "平均ジッタ(ms)"
  Add-ValueField $pt2 "loss_pct"       "平均ロス(%)"

  # 最新タイムスタンプのみ選択（可能なら）
  try {
    $pfTs = $pt2.PivotFields("timestamp")
    $last = $null
    foreach($pi in $pfTs.PivotItems()){ $last = $pi }
    if($last){ $pfTs.ClearAllFilters(); $last.Visible = $true }
  } catch {}

  try { foreach($co in @($wsP2.ChartObjects())){ $co.Delete() } } catch {}
  $ch2 = $wsP2.ChartObjects().Add(400, 20, 900, 380)
  $ch2.Name = "chtHops"
  $ch2.Chart.SetSourceData($pt2.TableRange1)
  $ch2.Chart.ChartType = $xlColumnClustered
  $ch2.Chart.HasTitle = $true
  $ch2.Chart.ChartTitle.Text = "Hops: 平均ジッタ/ロス（target・timestampで絞り込み可）"
  $ch2.Chart.Legend.Position = $xlLegendPositionBottom

  Add-Slicer $pt2 "target"    $wsP2 20 20
  Add-Slicer $pt2 "timestamp" $wsP2 20 140
}

# ===== 体裁 & 保存 =====
foreach($n in @("EndToEnd","HopsRaw","Pivot_EndToEnd","Pivot_Hops")){
  try {
    $ws = $wb.Worksheets.Item($n)
    $ws.Columns.AutoFit() | Out-Null
  } catch {}
}

# プレースホルダー削除（最後の1枚禁止ルールに配慮）
if(($wb.Worksheets.Count -gt 1) -and (($wb.Worksheets | ForEach-Object Name) -contains "Scratch")){
  $wb.Worksheets("Scratch").Delete()
}

$wb.SaveAs($Output)
$wb.Close($true)
$excel.Quit()

# COM解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)    | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host "レポートを出力しました: $Output"
