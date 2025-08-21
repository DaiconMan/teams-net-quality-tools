<#
  Generate-NetQuality-HTMLReport.ps1 (PS 5.1 Compatible)
  - 入力CSV
    * teams_net_quality.csv（UTF-8 BOM）
    * targets.csv（UTF-8） … ヘッダー: role,key,label
    * node_roles.csv（UTF-8, 任意） … ヘッダー: ip_of_host,role,label,segment
    * floors.csv（UTF-8） … ヘッダー: bssid,area,floor,tag,(任意: ap / ap_name)
  - 主要仕様
    * フィルタ採用条件: host / hop_ip / target のいずれかが targets.key に一致した行のみ採用
    * 表示キー: 採用された行は常に teams_net_quality.csv の host を採用（host空欄時のみフォールバック）
    * role/label: host→target→hop の順で targets から付与、なければ node_roles で補完
    * SAAS 集計: 採用role=SAAS の行は http_head_ms、それ以外は icmp_avg_ms を使用
    * area/floor/tag: teams.bssid ↔ floors.bssid 一致のみで付与
    * AP表示: teams.ap_name → floors.(ap|ap_name) → BSSIDラベル
    * HTML UI:
        - 行クリックで詳細折りたたみ＋時系列ミニグラフ
        - 「最悪時間帯」列: 0–23時の平均で最悪な時間帯を求め、その時間帯内の最悪値も併記（例 15時台 (183 ms)）
        - ヘッダークリックでソート（トグル）
        - P95ヘッダーに解説ツールチップ
  - ログ: -EnableCompareLog で候補/採用/ドロップ/メトリクスを出力
  - 注意: 三項演算子( ?: )不使用 / $Host未使用 / UTF-8 BOM 出力 / OneDrive・日本語パス対応
#>

[CmdletBinding()]
param(
  [string]$QualityCsv = ".\teams_net_quality.csv",
  [string]$TargetsCsv = ".\targets.csv",
  [string]$NodeRoleCsv = ".\node_roles.csv",
  [string]$FloorsCsv = ".\floors.csv",
  [string]$OutHtml = ".\NetQuality-Report.html",

  # 比較ログ（任意）
  [switch]$EnableCompareLog,
  [string]$LogFile = ".\NetQuality-MatchLog.txt",
  [int]$MaxCompareLogLines = 200
)

function Parse-Double { param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $t = ($s -replace '[^0-9\.\-]', '')
  $val = 0.0
  if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$val)) { return [double]$val }
  return $null
}
function Get-Median { param([double[]]$arr)
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  if ($n % 2 -eq 1) { return [double]$s[[int][math]::Floor($n/2)] }
  $a = [double]$s[$n/2 - 1]; $b = [double]$s[$n/2]
  return ($a + $b) / 2.0
}
function Get-Percentile { param([double[]]$arr, [double]$p)
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  $k = ($n - 1) * $p
  $f = [math]::Floor($k); $c = [math]::Ceiling($k)
  if ($f -eq $c) { return [double]$s[[int]$k] }
  [double]$sf = $s[[int]$f]; [double]$sc = $s[[int]$c]
  return $sf + ($k - $f) * ($sc - $sf)
}
function Safe-Lower { param([string]$s) if ($null -eq $s) { return "" } return ($s.ToString().Trim().ToLower()) }
function Normalize-Bssid { param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $h = ($s -replace '[^0-9a-fA-F]', '').ToLower()
  if ($h.Length -ge 12) { return $h.Substring(0,12) }
  return $h
}
function Import-CsvUtf8 { param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { return @() }
  $lines = Get-Content -LiteralPath $Path -Encoding UTF8
  if ($lines -is [string]) { $lines = @($lines) }
  if ($lines.Count -eq 0) { return @() }
  return $lines | ConvertFrom-Csv
}
function Is-BlankOrUnknown { param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $true }
  $v = $s.Trim().ToLower()
  if ($v -eq "unknown" -or $v -eq "(unknown)" -or $v -eq "n/a" -or $v -eq "-") { return $true }
  return $false
}
function Format-BssidLabel { param([string]$hex12)
  if ([string]::IsNullOrWhiteSpace($hex12)) { return "(AP unknown)" }
  $h = ($hex12 -replace '[^0-9a-f]', '').ToLower()
  if ($h.Length -lt 12) { return "BSSID:" + $h }
  $parts = @(); for($i=0;$i -lt 12;$i+=2){ $parts += $h.Substring($i,2) }
  return ("BSSID: " + ($parts -join ":"))
}

# ===== ロガー =====
$script:LogWriter = $null
$script:CompareLogLines = 0
function Open-Logger { param([string]$Path)
  try {
    $full = [System.IO.Path]::GetFullPath($Path)
    $dir  = [System.IO.Path]::GetDirectoryName($full)
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
      if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    }
    $enc = New-Object System.Text.UTF8Encoding($true)
    $sw = New-Object System.IO.StreamWriter($full, $false, $enc)
    $script:LogWriter = $sw
    $script:LogWriter.WriteLine(("[{0}] Start logging" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))); $script:LogWriter.Flush()
    Write-Output ("ログ開始: {0}" -f $full)
  } catch { Write-Warning "ログファイルを開けませんでした。" }
}
function Close-Logger {
  if ($null -ne $script:LogWriter) {
    try { $script:LogWriter.WriteLine(("[{0}] End logging" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))); $script:LogWriter.Flush(); $script:LogWriter.Close() } catch {}
    $script:LogWriter = $null
  }
}
function Log-Line { param([string]$message)
  if (-not $EnableCompareLog) { return }
  if ($null -eq $script:LogWriter) { return }
  $script:LogWriter.WriteLine(("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss.fff"), $message))
  $script:CompareLogLines++
  if (($script:CompareLogLines % 50) -eq 0) { $script:LogWriter.Flush() }
}

# ===== CSV 読み込み =====
if (-not (Test-Path -LiteralPath $QualityCsv)) { Write-Error "QualityCsv が見つかりません: $QualityCsv"; exit 1 }
if (-not (Test-Path -LiteralPath $TargetsCsv)) { Write-Error "TargetsCsv が見つかりません: $TargetsCsv"; exit 1 }
if (-not (Test-Path -LiteralPath $FloorsCsv))  { Write-Error "FloorsCsv が見つかりません: $FloorsCsv"; exit 1 }

if ($EnableCompareLog) { Open-Logger -Path $LogFile }

$teams   = Import-CsvUtf8 -Path $QualityCsv
$targets = Import-CsvUtf8 -Path $TargetsCsv
$roles   = @()
if (Test-Path -LiteralPath $NodeRoleCsv) { $roles = Import-CsvUtf8 -Path $NodeRoleCsv }
$floors  = Import-CsvUtf8 -Path $FloorsCsv

Log-Line ("読み込み: teams={0}行, targets={1}行, roles={2}行, floors={3}行" -f $teams.Count,$targets.Count,$roles.Count,$floors.Count)

# ===== targets辞書 =====
$targetSet     = @{}; $roleByKey = @{}; $labelByKey = @{}
foreach($t in $targets){
  $k = Safe-Lower $t.key; if ([string]::IsNullOrWhiteSpace($k)) { continue }
  $targetSet[$k] = $true
  if (-not [string]::IsNullOrWhiteSpace($t.role))  { $roleByKey[$k]  = $t.role }
  if (-not [string]::IsNullOrWhiteSpace($t.label)) { $labelByKey[$k] = $t.label }
}
# node_roles（補助）
$roleByNode=@{}; $labelByNode=@{}; $segmentByNode=@{}
foreach($r in $roles){
  $nk = Safe-Lower $r.ip_of_host; if ([string]::IsNullOrWhiteSpace($nk)) { continue }
  if (-not [string]::IsNullOrWhiteSpace($r.role))    { $roleByNode[$nk]   = $r.role }
  if (-not [string]::IsNullOrWhiteSpace($r.label))   { $labelByNode[$nk]  = $r.label }
  if (-not [string]::IsNullOrWhiteSpace($r.segment)) { $segmentByNode[$nk]= $r.segment }
}
# floors（BSSID完全一致 + 任意AP名）
$areaByBssid=@{}; $floorByBssid=@{}; $tagByBssid=@{}; $apByBssid=@{}
foreach($f in $floors){
  $b = Normalize-Bssid $f.bssid
  if ([string]::IsNullOrWhiteSpace($b)) { continue }
  $areaByBssid[$b]  = $f.area
  if (-not [string]::IsNullOrWhiteSpace($f.floor)) { $floorByBssid[$b] = $f.floor }
  if (-not [string]::IsNullOrWhiteSpace($f.tag))   { $tagByBssid[$b]   = $f.tag }
  $apCandidate = $null
  if ($f.PSObject.Properties.Name -contains 'ap') { $apCandidate = $f.ap }
  if ($f.PSObject.Properties.Name -contains 'ap_name' -and [string]::IsNullOrWhiteSpace($apCandidate)) { $apCandidate = $f.ap_name }
  if (-not (Is-BlankOrUnknown $apCandidate)) { $apByBssid[$b] = $apCandidate }
}

# ===== 正規化 & マッチング（targets一致ならhost採用）=====
$qual = @()
$cntAdoptHost = 0
$cntMatchHost = 0; $cntMatchHop = 0; $cntMatchTarget = 0; $cntDropped = 0
$cntFloorHitExact = 0; $cntFloorMissEmpty = 0; $cntFloorMissNotFound = 0
$cntApFromTeams=0; $cntApFromFloors=0; $cntApFromBssid=0; $cntApUnknown=0
$lineBudget = $MaxCompareLogLines; if ($lineBudget -lt 0) { $lineBudget = 0 }

foreach($q in $teams){
  $hopKey = Safe-Lower $q.hop_ip
  $tgtKey = Safe-Lower $q.target
  $hstKey = Safe-Lower $q.host

  $inTgt_hst = $false; if (-not [string]::IsNullOrWhiteSpace($hstKey)) { if ($targetSet.ContainsKey($hstKey)) { $inTgt_hst = $true } }
  $inTgt_hop = $false; if (-not [string]::IsNullOrWhiteSpace($hopKey)) { if ($targetSet.ContainsKey($hopKey)) { $inTgt_hop = $true } }
  $inTgt_tgt = $false; if (-not [string]::IsNullOrWhiteSpace($tgtKey)) { if ($targetSet.ContainsKey($tgtKey)) { $inTgt_tgt = $true } }

  # 候補ロールの取得
  $roleHst = $null; if ($roleByKey.ContainsKey($hstKey)) { $roleHst = $roleByKey[$hstKey] }
  $roleHop = $null; if ($roleByKey.ContainsKey($hopKey)) { $roleHop = $roleByKey[$hopKey] }
  $roleTgt = $null; if ($roleByKey.ContainsKey($tgtKey)) { $roleTgt = $roleByKey[$tgtKey] }

  if ($EnableCompareLog -and $lineBudget -gt 0) {
    Log-Line ("cmp-candidates: host={0} in={1} role={2} | hop={3} in={4} role={5} | target={6} in={7} role={8}" -f `
      $hstKey,$inTgt_hst,($(if ($null -ne $roleHst){$roleHst}else{""})), `
      $hopKey,$inTgt_hop,($(if ($null -ne $roleHop){$roleHop}else{""})), `
      $tgtKey,$inTgt_tgt,($(if ($null -ne $roleTgt){$roleTgt}else{""})))
    $lineBudget = $lineBudget - 1
  }

  $anyMatch = $inTgt_hst -or $inTgt_hop -or $inTgt_tgt
  if (-not $anyMatch) {
    $cntDropped++
    if ($EnableCompareLog -and $lineBudget -gt 0) {
      Log-Line ("cmp-drop: no match (host={0}, hop={1}, target={2})" -f $hstKey,$hopKey,$tgtKey)
      $lineBudget = $lineBudget - 1
    }
    continue
  }

  # 表示キーは host を最優先
  $displayKey = $hstKey; $pickSrc = "host"
  if ([string]::IsNullOrWhiteSpace($displayKey)) {
    if ($inTgt_hop) { $displayKey = $hopKey; $pickSrc = "hop(fallback)" }
    elseif ($inTgt_tgt) { $displayKey = $tgtKey; $pickSrc = "target(fallback)" }
    else { $displayKey = $hopKey; $pickSrc = "fallback(empty-host)" }
  }

  # 採用role/label/segment: host→target→hop の順
  $matchRole=$null; $matchLabel=$null; $segmentVal=""
  if ($roleByKey.ContainsKey($hstKey)) {
    $matchRole  = $roleByKey[$hstKey]
    if ($labelByKey.ContainsKey($hstKey)) { $matchLabel = $labelByKey[$hstKey] }
    if ($segmentByNode.ContainsKey($hstKey)) { $segmentVal = $segmentByNode[$hstKey] }
    $cntMatchHost++
  } elseif ($roleByKey.ContainsKey($tgtKey)) {
    $matchRole  = $roleByKey[$tgtKey]
    if ($labelByKey.ContainsKey($tgtKey)) { $matchLabel = $labelByKey[$tgtKey] }
    if ($segmentByNode.ContainsKey($tgtKey)) { $segmentVal = $segmentByNode[$tgtKey] }
    $cntMatchTarget++
  } elseif ($roleByKey.ContainsKey($hopKey)) {
    $matchRole  = $roleByKey[$hopKey]
    if ($labelByKey.ContainsKey($hopKey)) { $matchLabel = $labelByKey[$hopKey] }
    if ($segmentByNode.ContainsKey($hopKey)) { $segmentVal = $segmentByNode[$hopKey] }
    $cntMatchHop++
  } else {
    if ($segmentByNode.ContainsKey($displayKey)) { $segmentVal = $segmentByNode[$displayKey] }
  }
  if ([string]::IsNullOrWhiteSpace($matchRole))  { $matchRole  = "Uncategorized" }
  if ([string]::IsNullOrWhiteSpace($matchLabel)) { $matchLabel = $displayKey }

  $cntAdoptHost++
  if ($EnableCompareLog -and $lineBudget -gt 0) {
    Log-Line ("cmp-pick: displayKey={0} src={1} role={2} label={3}" -f $displayKey,$pickSrc,$matchRole,$matchLabel)
    $lineBudget = $lineBudget - 1
  }

  # floors: BSSID → area/floor/tag + AP
  $bNorm = Normalize-Bssid $q.bssid
  $areaVal = "Unknown"; $floorVal = $null; $tagVal = $null
  $bssidStatus = ""
  if ($null -eq $bNorm) { $cntFloorMissEmpty++; $bssidStatus="bssid=EMPTY" }
  else {
    if ($areaByBssid.ContainsKey($bNorm)) {
      $areaVal = $areaByBssid[$bNorm]
      if ($floorByBssid.ContainsKey($bNorm)) { $floorVal = $floorByBssid[$bNorm] }
      if ($tagByBssid.ContainsKey($bNorm))   { $tagVal   = $tagByBssid[$bNorm] }
      $cntFloorHitExact++; $bssidStatus="bssid=HIT"
    } else { $cntFloorMissNotFound++; $bssidStatus="bssid=NOTFOUND" }
  }

  # AP名解決
  $apFromTeams = $q.ap_name
  $apFromFloors = $null; if ($null -ne $bNorm -and $apByBssid.ContainsKey($bNorm)) { $apFromFloors = $apByBssid[$bNorm] }
  $apLabel = $null
  if (-not (Is-BlankOrUnknown $apFromTeams)) { $apLabel = $apFromTeams; $cntApFromTeams++ }
  elseif (-not (Is-BlankOrUnknown $apFromFloors)) { $apLabel = $apFromFloors; $cntApFromFloors++ }
  elseif ($null -ne $bNorm) { $apLabel = Format-BssidLabel $bNorm; $cntApFromBssid++ }
  else { $apLabel = "(AP unknown)"; $cntApUnknown++ }
  $apKey = $null; if ($null -ne $bNorm) { $apKey = $bNorm } else { $apKey = Safe-Lower $apLabel }

  # 数値系：採用roleが SAAS なら http_head_ms、他は icmp_avg_ms
  $rtt  = Parse-Double $q.icmp_avg_ms
  $http = Parse-Double $q.http_head_ms
  $roleNorm = $matchRole.ToString().Trim().ToUpper()
  $metric = $null; $metricKind = "RTT"
  if ($roleNorm -eq "SAAS") { if ($null -ne $http) { $metric = $http; $metricKind = "HTTP" } else { $metric = $rtt; $metricKind = "RTT" } }
  else { $metric = $rtt; $metricKind = "RTT" }
  $jit  = Parse-Double $q.icmp_jitter_ms
  $loss = Parse-Double $q.loss_pct
  $mos  = Parse-Double $q.mos_estimate
  if ($null -eq $mos -and $null -ne $rtt -and $null -ne $loss) { $mos = [math]::Round((4.5 - 0.0004*[double]$rtt - 0.1*[double]$loss),2) }

  if ($EnableCompareLog -and $lineBudget -gt 0) {
    $lbssid = if ($null -eq $bNorm) { "(empty)" } else { $bNorm }
    Log-Line ("cmp-metric: key={0} role={1} kind={2} value={3} bssid={4} area={5} ap={6}" -f $displayKey,$roleNorm,$metricKind,$metric,$lbssid,$areaVal,$apLabel)
    $lineBudget = $lineBudget - 1
  }

  $obj = [PSCustomObject]@{
    timestamp = $q.timestamp
    target    = $displayKey
    role      = $matchRole
    label     = $matchLabel
    segment   = $segmentVal
    metric_ms = $metric
    metric_kind = $metricKind
    rtt_ms    = $rtt
    jitter_ms = $jit
    loss_pct  = $loss
    mos       = $mos
    ssid      = $q.ssid
    bssid     = $bNorm
    ap_key    = $apKey
    ap_label  = $apLabel
    area      = $areaVal
    floor     = $floorVal
    ap_tag    = $tagVal
  }
  $qual += $obj
}

# ===== 明細（グラフ用） =====
$detailMap = @{}
foreach($o in $qual){
  $gk = ('{0}|{1}|{2}|{3}|{4}' -f $o.area, $o.ap_key, $o.target, $o.role, $o.segment)
  if (-not $detailMap.ContainsKey($gk)) { $detailMap[$gk] = New-Object System.Collections.ArrayList }
  if ($null -ne $o.metric_ms -and -not [string]::IsNullOrWhiteSpace($o.timestamp)) {
    [void]$detailMap[$gk].Add([PSCustomObject]@{ ts=$o.timestamp; v=[double]$o.metric_ms })
  }
}

# ===== 集計 =====
$summaryRows = @()
$groups = $qual | Group-Object -Property area, ap_key, target, role, segment
foreach($g in $groups){
  $vals = @($g.Group | Where-Object { $_.metric_ms -ne $null } | ForEach-Object {[double]$_.metric_ms})
  $jits = @($g.Group | Where-Object { $_.jitter_ms -ne $null } | ForEach-Object {[double]$_.jitter_ms})
  $loss = @($g.Group | Where-Object { $_.loss_pct  -ne $null } | ForEach-Object {[double]$_.loss_pct})
  $mosv = @($g.Group | Where-Object { $_.mos       -ne $null } | ForEach-Object {[double]$_.mos})

  $med = $null; if($vals.Count -gt 0){ $med = [math]::Round((Get-Median $vals),1) }
  $p95 = $null; if($vals.Count -gt 0){ $p95 = [math]::Round((Get-Percentile $vals 0.95),1) }
  $jit = $null; if($jits.Count -gt 0){ $jit = [math]::Round((Get-Median $jits),1) }
  $los = $null; if($loss.Count -gt 0){ $los = [math]::Round(($loss | Measure-Object -Average | Select-Object -ExpandProperty Average),2) }
  $mos = $null; if($mosv.Count -gt 0){ $mos = [math]::Round((Get-Median $mosv),2) }

  $first = $g.Group[0]
  $gkey  = ('{0}|{1}|{2}|{3}|{4}' -f $first.area, $first.ap_key, $first.target, $first.role, $first.segment)

  $summaryRows += [PSCustomObject]@{
    gkey     = $gkey
    area     = $first.area
    ap_key   = $first.ap_key
    ap_label = $first.ap_label
    target   = $first.target
    role     = $first.role
    segment  = $first.segment
    count    = $g.Count
    resp_med = $med
    resp_p95 = $p95
    jit_med  = $jit
    loss_avg = $los
    mos_med  = $mos
  }
}

# ===== JSON 生成 =====
$summaryJson = $summaryRows | ConvertTo-Json -Depth 6
$detailsOrdered = [ordered]@{}
foreach($k in $detailMap.Keys){ $detailsOrdered[$k] = $detailMap[$k] }
$detailsJson = ($detailsOrdered | ConvertTo-Json -Depth 6)

# ===== HTML =====
$htmlTemplate = @'
<!doctype html>
<html lang="ja"><head>
<meta charset="utf-8" />
<title>NetQuality Report (クリック展開グラフ＋ソート)</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
  body { font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Hiragino Kaku Gothic ProN","Noto Sans JP",sans-serif; margin: 16px; }
  h1 { font-size: 20px; margin: 0 0 12px; }
  .filters { display:flex; gap:8px; flex-wrap: wrap; margin: 8px 0 16px; }
  select, input { padding:6px; border:1px solid #ccc; border-radius: 8px; }
  table { border-collapse: collapse; width: 100%; margin: 8px 0 24px; }
  th, td { border-bottom: 1px solid #eee; padding: 8px; text-align: left; }
  th { background: #fafafa; position: sticky; top:0; z-index: 1; cursor: pointer; user-select:none; }
  tr.detail td { background:#fafcff; padding:12px 8px; }
  .mono { font-family: ui-monospace, Menlo, Consolas, "Liberation Mono", monospace; }
  .rt-ok{background:#e7f7e7;} .rt-warn{background:#fff5e0;} .rt-bad{background:#fdecec;} .loss-bad{background:#fdecec;}
  .muted{color:#777;}
  .hint{color:#666; font-size:12px;}
  .canvaswrap{width:100%; height:180px;}
  .meta { font-size:12px; color:#444; margin:4px 0 0; }
  .sort-ind{margin-left:6px; opacity:.6;}
</style>
</head><body>
  <h1>ネットワーク品質レポート（targets一致ならhost採用／クリック展開＆ヘッダーソート）</h1>

  <div class="filters">
    <select id="areaSel"><option value="">(すべてのエリア)</option></select>
    <select id="apSel"><option value="">(すべてのAP)</option></select>
    <input id="targetSearch" placeholder="対象(部分一致)" />
    <select id="roleSel"><option value="">(すべての役割)</option></select>
    <select id="segSel"><option value="">(すべてのセグメント)</option></select>
  </div>

  <table id="sumTbl">
    <thead><tr>
      <th data-sort="area">エリア<span class="sort-ind"></span></th>
      <th data-sort="ap_label">AP<span class="sort-ind"></span></th>
      <th data-sort="target">対象<span class="sort-ind"></span></th>
      <th data-sort="role">役割<span class="sort-ind"></span></th>
      <th data-sort="segment">セグメント<span class="sort-ind"></span></th>
      <th data-sort="count">試行数<span class="sort-ind"></span></th>
      <th data-sort="resp_med">応答中央値(ms)<span class="sort-ind"></span></th>
      <th data-sort="resp_p95" title="P95: 応答遅延の95パーセンタイル。全データを小さい順に並べ、上位5%を除いた境界値。ピーク遅延の傾向把握に有効。">P95(ms)<span class="sort-ind"></span></th>
      <th data-sort="jit_med">ジッタ中央値<span class="sort-ind"></span></th>
      <th data-sort="loss_avg">損失率(平均)<span class="sort-ind"></span></th>
      <th data-sort="mos_med">MOS(中央値)<span class="sort-ind"></span></th>
      <th data-sort="worst_value">最悪時間帯<span class="sort-ind"></span></th>
    </tr></thead><tbody></tbody>
  </table>

<script>
var summaryRows = __SUMMARY_JSON__;
var detailsMap  = __DETAILS_JSON__; // { gkey: [ {ts:"...", v:number}, ... ] }

function uniq(a){var o=[],i;for(i=0;i<a.length;i++){var x=a[i];if(x&&o.indexOf(x)===-1)o.push(x);}o.sort();return o;}
function fillSel(el,opts){for(var i=0;i<opts.length;i++){var op=document.createElement('option');op.textContent=opts[i];op.value=opts[i];el.appendChild(op);}}

var areaSel=document.getElementById('areaSel'),
    apSel=document.getElementById('apSel'),
    roleSel=document.getElementById('roleSel'),
    segSel=document.getElementById('segSel'),
    qInput=document.getElementById('targetSearch');

fillSel(areaSel,uniq(summaryRows.map(function(r){return r.area;})));
fillSel(apSel,uniq(summaryRows.map(function(r){return r.ap_label;}).filter(function(x){return !!x;})));
fillSel(roleSel,uniq(summaryRows.map(function(r){return r.role;})));
fillSel(segSel,uniq(summaryRows.map(function(r){return r.segment;})));

[areaSel,apSel,roleSel,segSel,qInput].forEach(function(el){el.addEventListener('input',render);el.addEventListener('change',render);});

function colorResp(td,v){if(v==null)return;if(v<50)td.classList.add('rt-ok');else if(v<100)td.classList.add('rt-warn');else td.classList.add('rt-bad');}
function colorLoss(td,v){if(v!=null&&v>3)td.classList.add('loss-bad');}
function hourString(h){ return (h<10?("0"+h):h) + "時台"; }

var worstCache={};
function worstHourAndValue(points,gkey){
  if(worstCache[gkey]) return worstCache[gkey];
  var res={label:"", hour:-1, value:null};
  if(!points||points.length===0){ worstCache[gkey]=res; return res; }

  var buckets=new Array(24), counts=new Array(24);
  for(var i=0;i<24;i++){buckets[i]=0;counts[i]=0;}
  for(var i=0;i<points.length;i++){
    var t=new Date(points[i].ts); if(isNaN(t)) continue;
    var h=t.getHours(), v=points[i].v; if(v==null) continue;
    buckets[h]+=v; counts[h]++;
  }
  var worstAvg=-1, wh=-1;
  for(var h=0;h<24;h++){
    if(counts[h]===0) continue;
    var avg=buckets[h]/counts[h];
    if(avg>worstAvg){ worstAvg=avg; wh=h; }
  }
  if(wh<0){ worstCache[gkey]=res; return res; }

  // その時間帯に属するデータの中で最悪値
  var maxInHour=null;
  for(var i=0;i<points.length;i++){
    var t=new Date(points[i].ts); if(isNaN(t)) continue;
    if(t.getHours()!==wh) continue;
    var v=points[i].v; if(v==null) continue;
    if(maxInHour==null || v>maxInHour) maxInHour=v;
  }
  res.hour=wh; res.value=maxInHour;
  res.label = hourString(wh) + (maxInHour!=null?(" ("+Math.round(maxInHour)+" ms)"):"");
  worstCache[gkey]=res; return res;
}

var sortKey=null, sortAsc=true;
function setSortIndicator(){
  var ths=document.querySelectorAll('thead th');
  for(var i=0;i<ths.length;i++){
    var th=ths[i]; var span=th.querySelector('.sort-ind');
    if(!span) continue;
    var k=th.getAttribute('data-sort');
    if(k && k===sortKey){ span.textContent = sortAsc ? "▲" : "▼"; }
    else { span.textContent=""; }
  }
}

var tableBody=document.querySelector('#sumTbl tbody');

function drawLineChart(canvas, points){
  var ctx=canvas.getContext('2d');
  var W=canvas.width, H=canvas.height;
  ctx.clearRect(0,0,W,H);
  if(!points || points.length===0){ ctx.fillText("データなし",10,14); return; }

  var minT=Infinity, maxT=-Infinity, minV=Infinity, maxV=-Infinity;
  for(var i=0;i<points.length;i++){
    var t = new Date(points[i].ts).getTime(); if(isNaN(t)) continue;
    var v = points[i].v; if(v==null) continue;
    if(t<minT)minT=t; if(t>maxT)maxT=t;
    if(v<minV)minV=v; if(v>maxV)maxV=v;
  }
  if(!isFinite(minT)||!isFinite(maxT)||minT===maxT){ minT=Date.now()-60000; maxT=Date.now(); }
  if(!isFinite(minV)||!isFinite(maxV)||minV===maxV){ minV=0; maxV=(isFinite(points[0].v)?points[0].v:100); if(maxV<=0)maxV=100; }

  var padL=40, padR=10, padT=10, padB=22;
  function x(t){ return padL + ( (t-minT)/(maxT-minT) ) * (W-padL-padR); }
  function y(v){ return H-padB - ( (v-minV)/(maxV-minV) ) * (H-padT-padB); }

  ctx.beginPath(); ctx.moveTo(padL, padT); ctx.lineTo(padL, H-padB); ctx.lineTo(W-padR, H-padB); ctx.stroke();

  ctx.font="12px sans-serif"; ctx.textAlign="right"; ctx.textBaseline="middle";
  var ticks=4; for(var i=0;i<=ticks;i++){ var vv=minV+(maxV-minV)*i/ticks; var yy=y(vv);
    ctx.fillText(vv.toFixed(0), padL-6, yy); ctx.beginPath(); ctx.moveTo(padL,yy); ctx.lineTo(W-padR,yy); ctx.strokeStyle="rgba(0,0,0,0.06)"; ctx.stroke(); ctx.strokeStyle="black";
  }

  ctx.beginPath();
  var first=true, worstV=-Infinity, worstX=0, worstY=0;
  for(var i=0;i<points.length;i++){
    var tt=new Date(points[i].ts).getTime(); if(isNaN(tt)) continue;
    var vv=points[i].v; if(vv==null) continue;
    var xx=x(tt), yy=y(vv);
    if(first){ ctx.moveTo(xx,yy); first=false; } else { ctx.lineTo(xx,yy); }
    if(vv>worstV){ worstV=vv; worstX=xx; worstY=yy; }
  }
  ctx.stroke();

  ctx.beginPath(); ctx.arc(worstX,worstY,3,0,6.283); ctx.fill();

  ctx.textAlign="left"; ctx.textBaseline="top";
  var minD=new Date(minT), maxD=new Date(maxT);
  ctx.fillText(minD.toLocaleString(), padL, H-20);
  ctx.textAlign="right";
  ctx.fillText(maxD.toLocaleString(), W-10, H-20);
}

function render(){
  var area=areaSel.value||"",ap=apSel.value||"",role=roleSel.value||"",seg=segSel.value||"",q=(qInput.value||"").toLowerCase();

  // フィルタ
  var rows=summaryRows.filter(function(r){
    if(area && r.area!==area) return false;
    if(ap && (r.ap_label||"")!==ap) return false;
    if(role && r.role!==role) return false;
    if(seg && r.segment!==seg) return false;
    if(q && String(r.target||"").toLowerCase().indexOf(q)===-1) return false;
    return true;
  });

  // ソート
  rows.sort(function(a,b){
    var ka, kb;
    if(sortKey==="worst_value"){
      var wa=worstHourAndValue(detailsMap[a.gkey]||[], a.gkey).value;
      var wb=worstHourAndValue(detailsMap[b.gkey]||[], b.gkey).value;
      ka=(wa==null?-Infinity:wa); kb=(wb==null?-Infinity:wb);
    }else if(sortKey){
      ka=a[sortKey]; kb=b[sortKey];
      // 数値は数値で比較
      var numKeys={"count":1,"resp_med":1,"resp_p95":1,"jit_med":1,"loss_avg":1,"mos_med":1};
      if(numKeys[sortKey]){ ka=(ka==null?-Infinity:ka); kb=(kb==null?-Infinity:kb); }
      else { ka=(ka||""); kb=(kb||""); }
    }else{
      // デフォルト：エリア→AP→応答中央値(昇順)
      var keyA=(a.area||"")+"|"+(a.ap_label||""); var keyB=(b.area||"")+"|"+(b.ap_label||"");
      if(keyA<keyB) return -1; if(keyA>keyB) return 1;
      var ra=(a.resp_med!=null)?a.resp_med:1e12, rb=(b.resp_med!=null)?b.resp_med:1e12;
      return ra-rb;
    }
    var cmp = (ka<kb)?-1:((ka>kb)?1:0);
    return sortAsc?cmp:-cmp;
  });

  // 描画
  tableBody.innerHTML='';
  for(var i=0;i<rows.length;i++){
    var r=rows[i];
    var pts = detailsMap[r.gkey] || [];
    var wres = worstHourAndValue(pts, r.gkey);

    var tr=document.createElement('tr'); tr.className="main"; tr.dataset.gkey=r.gkey;
    function td(t){var e=document.createElement('td'); e.textContent=(t==null?"":t); return e;}
    tr.appendChild(td(r.area));
    tr.appendChild(td(r.ap_label||""));
    var ttd=td(r.target); ttd.classList.add('mono'); tr.appendChild(ttd);
    tr.appendChild(td(r.role||""));
    tr.appendChild(td(r.segment||""));
    tr.appendChild(td(r.count));
    var rt=td(r.resp_med); colorResp(rt,r.resp_med); tr.appendChild(rt);
    var rp=td(r.resp_p95); colorResp(rp,r.resp_p95); tr.appendChild(rp);
    tr.appendChild(td(r.jit_med));
    var ls=td(r.loss_avg); colorLoss(ls,r.loss_avg); tr.appendChild(ls);
    tr.appendChild(td(r.mos_med));
    tr.appendChild(td(wres.label));

    // 詳細行
    var dtr=document.createElement('tr'); dtr.className="detail"; dtr.style.display="none";
    var tdwrap=document.createElement('td'); tdwrap.colSpan=12;
    var div=document.createElement('div'); div.className="canvaswrap";
    var canvas=document.createElement('canvas'); canvas.width=tdwrap.clientWidth||800; canvas.height=180;
    div.appendChild(canvas);
    var meta=document.createElement('div'); meta.className="meta";
    meta.textContent="クリックで開閉／縦軸=ms（role=SAAS はHTTPヘッダ遅延、それ以外はICMP RTT）";
    tdwrap.appendChild(div); tdwrap.appendChild(meta); dtr.appendChild(tdwrap);

    tr.addEventListener('click', function(dtrRef,cvRef,points){
      return function(){
        if(dtrRef.style.display==="none"){ dtrRef.style.display="table-row"; drawLineChart(cvRef, points); }
        else { dtrRef.style.display="none"; }
      };
    }(dtr,canvas,pts));

    tableBody.appendChild(tr); tableBody.appendChild(dtr);
  }
  setSortIndicator();
}
render();

// ヘッダークリックでソート
var ths=document.querySelectorAll('thead th[data-sort]');
for(var i=0;i<ths.length;i++){
  ths[i].addEventListener('click',function(){
    var k=this.getAttribute('data-sort');
    if(sortKey===k){ sortAsc=!sortAsc; } else { sortKey=k; sortAsc=true; }
    render();
  });
}
</script>
</body></html>
'@

$html = $htmlTemplate.Replace('__SUMMARY_JSON__', $summaryJson).Replace('__DETAILS_JSON__', $detailsJson)

try {
  $fullPath = [System.IO.Path]::GetFullPath($OutHtml)
  $dir = [System.IO.Path]::GetDirectoryName($fullPath)
  if (-not [string]::IsNullOrWhiteSpace($dir)) {
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
  }
  $enc = New-Object System.Text.UTF8Encoding($true) # BOM付き
  [System.IO.File]::WriteAllText($fullPath, $html, $enc)
  Write-Output ("HTMLレポートを出力しました(UTF-8 BOM): {0}" -f $fullPath)
} catch {
  $html | Out-File -FilePath $OutHtml
  Write-Warning "WriteAllText に失敗したため Out-File で出力しました。"
}

# 参考メトリクス
$total = $teams.Count
$after = $qual.Count
$mappedArea = ($qual | Where-Object { $_.area -ne "Unknown" }).Count
if ($total -gt 0) {
  $pctUsed = [math]::Round(100.0 * $after / $total, 1)
  $pctArea = 0.0
  if ($after -gt 0) { $pctArea = [math]::Round(100.0 * $mappedArea / $after, 1) }
  $summary = ("teams行数: {0}, 採用行数: {1} ({2}%) | area付与率: {3}% ({4}/{5}) | match(host/target/hop) host={6} target={7} hop={8} | host採用行={9} | floors: HIT={10} / EMPTY={11} / NOTFOUND={12} | AP解決: teams={13} floors={14} bssid={15} unknown={16}" -f `
    $total,$after,$pctUsed,$pctArea,$mappedArea,$after,$cntMatchHost,$cntMatchTarget,$cntMatchHop,$cntAdoptHost,$cntFloorHitExact,$cntFloorMissEmpty,$cntFloorMissNotFound,$cntApFromTeams,$cntApFromFloors,$cntApFromBssid,$cntApUnknown)
  Write-Output $summary
  Log-Line $summary
}

Close-Logger
