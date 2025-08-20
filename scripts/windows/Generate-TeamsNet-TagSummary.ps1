<#
.SYNOPSIS
  Tag(=測定場所)ごとの品質サマリを作成し、Excel/CSVで出力する（PowerShell 5.1対応）

.PARAMETER CsvPath
  マージ済みの入力CSVパス（tags/probe列を含む想定）

.PARAMETER Output
  出力フォルダ（既定: .\Output）

.PARAMETER TagColumn
  Tag 列の明示指定（既定: 自動検出: probe|tag|location）

.PARAMETER TargetColumn
  Target 列の明示指定（既定: 自動検出: target|host|hostname|dst|endpoint|server）
  ※互換のため -HostColumn エイリアスも受理します（内部では使いません）

.PARAMETER LatencyColumn
  レイテンシ列の明示指定（既定: 自動検出: rtt_ms|rtt|latency_ms|latency|avg_ms|response_ms）

.PARAMETER ThresholdMs
  遅延のしきい値（既定: 100）

.OUTPUTS
  Output\TagSummary.xlsx, Output\TagSummary.csv

.NOTES
  - Excel COM を使用（Office がない環境では CSV のみ出力）
  - 列名は大小区別せず照合
  - PowerShell 5.1 互換
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$CsvPath,

  [string]$Output = ".\Output",

  [string]$TagColumn,

  [Alias('HostColumn')]
  [string]$TargetColumn,

  [string]$LatencyColumn,

  [int]$ThresholdMs = 100
)

function Write-Info($msg){ Write-Host "[TagSummary] $msg" }

function Test-File([string]$Path){
  if(-not (Test-Path -LiteralPath $Path)){
    throw "File not found: $Path"
  }
}

function Find-FirstPresentColumn($Headers, [string[]]$Candidates){
  foreach($c in $Candidates){
    if($Headers -contains $c){ return $c }
  }
  return $null
}

function Detect-Columns($Headers){
  # 小文字でそろえる
  $h = @($Headers | ForEach-Object { ($_ -as [string]).ToLowerInvariant() })

  $tag = Find-FirstPresentColumn $h @('probe','tag','location')
  $tgt = Find-FirstPresentColumn $h @('target','host','hostname','dst','endpoint','server')
  $lat = Find-FirstPresentColumn $h @('rtt_ms','rtt','latency_ms','latency','avg_ms','response_ms')

  @{ Tag=$tag; Target=$tgt; Latency=$lat }
}

function Percentile([double[]]$values, [double]$p){
  if(-not $values -or $values.Count -eq 0){ return $null }
  $arr = @($values | Sort-Object)
  $n = $arr.Count
  if($n -eq 1){ return $arr[0] }
  $rank = [math]::Ceiling($p * $n)
  $idx = [math]::Min([math]::Max($rank-1,0), $n-1)
  return $arr[$idx]
}

function Ensure-Dir($path){
  if(-not (Test-Path -LiteralPath $path)){ New-Item -ItemType Directory -Path $path | Out-Null }
}

function Sanitize-WorksheetName([string]$name){
  if([string]::IsNullOrWhiteSpace($name)){ return "_(blank)" }
  $n = $name -replace '[\\\/\?\*\[\]:]', '_'
  if($n.Length -gt 31){ return $n.Substring(0,31) }
  return $n
}

try{
  Test-File $CsvPath
  Ensure-Dir $Output

  Write-Info "Loading CSV: $CsvPath"
  $rows = Import-Csv -LiteralPath $CsvPath
  if(-not $rows){ throw "CSV rows are empty." }

  # ヘッダ取得（小文字化）
  $headersLower = @($rows[0].PSObject.Properties.Name | ForEach-Object { ($_ -as [string]).ToLowerInvariant() })

  # 列自動検出または指定優先
  $det = Detect-Columns $headersLower
  $tagCol = if($TagColumn){ $TagColumn } else { $det.Tag }
  $tgtCol = if($TargetColumn){ $TargetColumn } else { $det.Target }
  $latCol = if($LatencyColumn){ $LatencyColumn } else { $det.Latency }

  if(-not $tagCol){ throw "Tag 列が見つかりません。-TagColumn で指定するか、CSV に probe|tag|location のいずれかの列を含めてください。" }
  if(-not $latCol){ throw "レイテンシ列が見つかりません。-LatencyColumn で指定するか、CSV に rtt_ms|rtt|latency_ms|latency|avg_ms|response_ms のいずれかの列を含めてください。" }

  Write-Info ("Using columns -> Tag:'{0}', Target:'{1}', Latency:'{2}'" -f $tagCol,$tgtCol,$latCol)

  # 正規化して投影
  $proj = foreach($r in $rows){
    $tag = ($r.$tagCol) -as [string]
    $tgt = if($tgtCol){ ($r.$tgtCol) -as [string] } else { $null }

    # 数値化（'123ms' 形式にも一部対応）
    $raw = ($r.$latCol) -as [string]
    if($raw -match '([-+]?\d+(\.\d+)?)'){ $lat = [double]$matches[1] } else { $lat = $null }

    [pscustomobject]@{
      Tag = $tag
      Target = if($tgt){ $tgt } else { "_(all)" }
      LatencyMs = $lat
    }
  }

  # 欠損除外
  $proj = $proj | Where-Object { $_.LatencyMs -ne $null }
  if(-not $proj){ throw "有効なレイテンシ行がありません（数値に変換できませんでした）。" }

  # 集計（Tag, Target）
  Write-Info "Aggregating by Tag/Target..."
  $grouped = $proj | Group-Object Tag, Target

  $outRows = foreach($g in $grouped){
    $tag = $g.Group[0].Tag
    $tgt = $g.Group[0].Target
    $vals = @($g.Group | ForEach-Object { $_.LatencyMs })
    $n = $vals.Count
    $avg = [math]::Round(($vals | Measure-Object -Average).Average, 1)
    $p95 = [math]::Round((Percentile $vals 0.95), 1)
    $over = @($vals | Where-Object { $_ -ge $ThresholdMs }).Count
    $overPct = if($n -gt 0){ [math]::Round(100.0 * $over / $n, 1) } else { $null }

    [pscustomobject]@{
      Tag = if([string]::IsNullOrWhiteSpace($tag)){"_(blank)"}else{$tag}
      Target = $tgt
      Samples = $n
      AvgMs = $avg
      P95Ms = $p95
      OverThresholdCount = $over
      OverThresholdPct = $overPct
      ThresholdMs = $ThresholdMs
    }
  } | Sort-Object Tag, Target

  # CSV 出力
  $csvOut = Join-Path $Output "TagSummary.csv"
  Write-Info "Writing CSV: $csvOut"
  $outRows | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $csvOut

  # Excel 出力（Office 環境のみ）
  $xlsxOut = Join-Path $Output "TagSummary.xlsx"
  $excel = $null
  try{
    $excel = New-Object -ComObject Excel.Application
  }catch{
    Write-Info "Excel COM を作成できませんでした。CSV のみ出力します。"
  }

  if($excel){
    try{
      $excel.Visible = $false
      $wb = $excel.Workbooks.Add()
      # Summary シート
      $ws = $wb.Worksheets.Item(1)
      $ws.Name = "Summary"
      $headers = @('Tag','Target','Samples','AvgMs','P95Ms','OverThresholdCount','OverThresholdPct','ThresholdMs')
      for($i=0;$i -lt $headers.Count;$i++){
        $ws.Cells.Item(1, $i+1) = $headers[$i]
      }
      $row = 2
      foreach($r in $outRows){
        $ws.Cells.Item($row,1) = $r.Tag
        $ws.Cells.Item($row,2) = $r.Target
        $ws.Cells.Item($row,3) = $r.Samples
        $ws.Cells.Item($row,4) = $r.AvgMs
        $ws.Cells.Item($row,5) = $r.P95Ms
        $ws.Cells.Item($row,6) = $r.OverThresholdCount
        $ws.Cells.Item($row,7) = $r.OverThresholdPct
        $ws.Cells.Item($row,8) = $r.ThresholdMs
        $row++
      }
      # Tag ごとに別シート
      $byTag = $outRows | Group-Object Tag
      foreach($tg in $byTag){
        $name = Sanitize-WorksheetName $tg.Name
        $sheet = $wb.Worksheets.Add()
        $null = $sheet.Move($ws)   # 先頭へ
        $sheet.Name = $name
        for($i=0;$i -lt $headers.Count;$i++){
          $sheet.Cells.Item(1, $i+1) = $headers[$i]
        }
        $ridx = 2
        foreach($r in ($tg.Group | Sort-Object Target)){
          $sheet.Cells.Item($ridx,1) = $r.Tag
          $sheet.Cells.Item($ridx,2) = $r.Target
          $sheet.Cells.Item($ridx,3) = $r.Samples
          $sheet.Cells.Item($ridx,4) = $r.AvgMs
          $sheet.Cells.Item($ridx,5) = $r.P95Ms
          $sheet.Cells.Item($ridx,6) = $r.OverThresholdCount
          $sheet.Cells.Item($ridx,7) = $r.OverThresholdPct
          $sheet.Cells.Item($ridx,8) = $r.ThresholdMs
          $ridx++
        }
      }

      # 体裁
      for($i=1; $i -le $wb.Worksheets.Count; $i++){
        $s = $wb.Worksheets.Item($i)
        $s.Rows.Item(1).Font.Bold = $true
        $null = $s.Columns.AutoFit()
      }

      Write-Info "Saving Excel: $xlsxOut"
      $wb.SaveAs($xlsxOut)
      $wb.Close($true)
    }finally{
      $excel.Quit() | Out-Null
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    }
  }

  Write-Info "Done."
}catch{
  Write-Error $_.Exception.Message
  exit 1
}
