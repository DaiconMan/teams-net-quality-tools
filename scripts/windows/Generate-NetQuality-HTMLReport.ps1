<#
  Generate-NetQuality-HTMLReport.ps1 (PS 5.1 Compatible)
  - 入力CSV（文字コード）
    * teams_net_quality.csv（UTF-8 BOM） … 指定ヘッダー:
      timestamp,target,hop_index,hop_ip,icmp_avg_ms,icmp_jitter_ms,loss_pct,notes,conn_type,ssid,bssid,signal_pct,ap_name,roamed,roam_from,roam_to,host,dns_ms,tcp_443_ms,http_head_ms,mos_estimate,probe,machine,user,tz_offset,source_file
    * targets.csv（UTF-8） … ヘッダー: role,key,label
      - role: L2 / L3 / SAAS / RTR_WAN / RTR_LAN
      - key : FQDN or IP（teams_net_quality.csv の target と一致させる）
      - label: 機器名やサービス名
    * node_roles.csv（UTF-8） … ヘッダー: ip_of_host,role,label,segment
    * floors.csv（UTF-8） … ヘッダー: bssid,area,floor,tag
  - 処理ポリシー
    * 宛先（target）は必ず targets.csv の key に含まれるもの**のみ**採用（それ以外は除外）
    * 役割/ラベルは targets.csv を最優先。無い場合のみ node_roles.csv を参照
    * floors.csv の bssid で area / floor / tag を付与（BSSID正規化・接頭辞「*」不要、完全一致のみ）
  - 出力: HTML単一ファイル（UTF-8 BOM, 外部ライブラリ不要）
  - 注意:
    * 三項演算子( ?: )は不使用
    * PowerShellの $Host は未使用
    * OneDrive/日本語/スペースを含むパス考慮（出力先ディレクトリ自動作成）
#>

[CmdletBinding()]
param(
  [string]$QualityCsv = ".\teams_net_quality.csv",
  [string]$TargetsCsv = ".\targets.csv",
  [string]$NodeRoleCsv = ".\node_roles.csv",
  [string]$FloorsCsv = ".\floors.csv",
  [string]$OutHtml = ".\NetQuality-Report.html"
)

# ===== 共通ヘルパ =====
function Parse-Double {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $t = ($s -replace '[^0-9\.\-]', '')
  $val = 0.0
  if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$val)) {
    return [double]$val
  }
  return $null
}

function Get-Median {[CmdletBinding()] param([double[]]$arr)
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  if ($n % 2 -eq 1) { return [double]$s[[int][math]::Floor($n/2)] }
  $a = [double]$s[$n/2 - 1]; $b = [double]$s[$n/2]
  return ($a + $b) / 2.0
}

function Get-Percentile {[CmdletBinding()] param([double[]]$arr, [double]$p)
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  $k = ($n - 1) * $p
  $f = [math]::Floor($k); $c = [math]::Ceiling($k)
  if ($f -eq $c) { return [double]$s[[int]$k] }
  [double]$sf = $s[[int]$f]; [double]$sc = $s[[int]$c]
  return $sf + ($k - $f) * ($sc - $sf)
}

function Safe-Lower { param([string]$s)
  if ($null -eq $s) { return "" }
  return ($s.ToString().Trim().ToLower())
}

function Normalize-Bssid { param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $h = ($s -replace '[^0-9a-fA-F]', '').ToLower()
  if ($h.Length -ge 12) { return $h.Substring(0,12) }
  return $h
}

function Import-CsvUtf8 {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { return @() }
  # PS5.1対策: Import-Csv に -Encoding なし。UTF-8(BOM有無)は Get-Content -Encoding UTF8 + ConvertFrom-Csv で統一。
  $lines = Get-Content -LiteralPath $Path -Encoding UTF8
  if ($lines -is [string]) { $lines = @($lines) }
  if ($lines.Count -eq 0) { return @() }
  return $lines | ConvertFrom-Csv
}

# ===== CSV 読み込み =====
if (-not (Test-Path -LiteralPath $QualityCsv)) { Write-Error "QualityCsv が見つかりません: $QualityCsv"; exit 1 }
if (-not (Test-Path -LiteralPath $TargetsCsv)) { Write-Error "TargetsCsv が見つかりません: $TargetsCsv"; exit 1 }
if (-not (Test-Path -LiteralPath $FloorsCsv))  { Write-Error "FloorsCsv が見つかりません: $FloorsCsv"; exit 1 }

$teams  = Import-CsvUtf8 -Path $QualityCsv
$targets= Import-CsvUtf8 -Path $TargetsCsv
$roles  = @()
if (Test-Path -LiteralPath $NodeRoleCsv) { $roles = Import-CsvUtf8 -Path $NodeRoleCsv }
$floors = Import-CsvUtf8 -Path $FloorsCsv

# ===== targets.csv を基準にフィルタ＆役割・ラベル優先付与 =====
# 形式: role,key,label
$targetSet     = @{}   # key(lower) => $true
$roleByTarget  = @{}   # key(lower) => role(大文字のまま)
$labelByTarget = @{}   # key(lower) => label
foreach($t in $targets){
  $key = Safe-Lower $t.key
  if ([string]::IsNullOrWhiteSpace($key)) { continue }
  if (-not $targetSet.ContainsKey($key)) { $targetSet[$key] = $true }
  if (-not [string]::IsNullOrWhiteSpace($t.role))  { $roleByTarget[$key]  = $t.role }
  if (-not [string]::IsNullOrWhiteSpace($t.label)) { $labelByTarget[$key] = $t.label }
}

# node_roles.csv（ip_of_host,role,label,segment）… targetsに無いキーの補助情報としてのみ使用
$roleByNode   = @{}
$labelByNode  = @{}
$segmentByNode= @{}
foreach($r in $roles){
  $k = Safe-Lower $r.ip_of_host
  if ([string]::IsNullOrWhiteSpace($k)) { continue }
  if (-not [string]::IsNullOrWhiteSpace($r.role))   { $roleByNode[$k]   = $r.role }
  if (-not [string]::IsNullOrWhiteSpace($r.label))  { $labelByNode[$k]  = $r.label }
  if (-not [string]::IsNullOrWhiteSpace($r.segment)){ $segmentByNode[$k]= $r.segment }
}

# floors.csv（bssid,area,floor,tag）→ BSSID完全一致辞書
$areaByBssid  = @{}
$floorByBssid = @{}
$tagByBssid   = @{}
foreach($f in $floors){
  $b = Normalize-Bssid $f.bssid
  if ([string]::IsNullOrWhiteSpace($b)) { continue }
  if (-not $areaByBssid.ContainsKey($b))  { $areaByBssid[$b]  = $f.area }
  if (-not [string]::IsNullOrWhiteSpace($f.floor)) { $floorByBssid[$b] = $f.floor }
  if (-not [string]::IsNullOrWhiteSpace($f.tag))   { $tagByBssid[$b]   = $f.tag }
}

# ===== teams_net_quality.csv を targets に含まれる宛先だけに絞る =====
$filtered = @()
$droppedNotInTargets = 0
foreach($row in $teams){
  $tgt = Safe-Lower $row.target
  if ([string]::IsNullOrWhiteSpace($tgt)) { $droppedNotInTargets++; continue }
  if (-not $targetSet.ContainsKey($tgt)) { $droppedNotInTargets++; continue }
  $filtered += $row
}
Write-Output ("targets.csvに含まれない宛先を除外: {0}件" -f $droppedNotInTargets)

# ===== 正規化 & マッピング =====
$qual = @()
$emptyBssid = 0; $emptyAp = 0
foreach($q in $filtered){
  $bNorm = Normalize-Bssid $q.bssid
  if ($null -eq $bNorm) { $emptyBssid++ }
  $apName = $q.ap_name
  if ([string]::IsNullOrWhiteSpace($apName)) { $emptyAp++ }

  # floors の area/floor/tag
  $areaVal = "Unknown"; $floorVal = $null; $tagVal = $null
  if ($bNorm -and $areaByBssid.ContainsKey($bNorm)) { $areaVal = $areaByBssid[$bNorm] }
  if ($bNorm -and $floorByBssid.ContainsKey($bNorm)) { $floorVal = $floorByBssid[$bNorm] }
  if ($bNorm -and $tagByBssid.ContainsKey($bNorm))   { $tagVal   = $tagByBssid[$bNorm] }

  # role/label/segment の優先順位: targets.csv ＞ node_roles.csv ＞ 既定
  $key = Safe-Lower $q.target
  $roleVal = "Uncategorized"
  if ($roleByTarget.ContainsKey($key)) { $roleVal = $roleByTarget[$key] }
  elseif ($roleByNode.ContainsKey($key)) { $roleVal = $roleByNode[$key] }

  $labelVal = $q.target
  if ($labelByTarget.ContainsKey($key)) { $labelVal = $labelByTarget[$key] }
  elseif ($labelByNode.ContainsKey($key)) { $labelVal = $labelByNode[$key] }

  $segVal = ""
  if ($segmentByNode.ContainsKey($key)) { $segVal = $segmentByNode[$key] }

  # 数値系
  $rtt  = Parse-Double $q.icmp_avg_ms
  $jit  = Parse-Double $q.icmp_jitter_ms
  $loss = Parse-Double $q.loss_pct
  $mos  = Parse-Double $q.mos_estimate
  if ($null -eq $mos -and $null -ne $rtt -and $null -ne $loss) {
    $mos = [math]::Round((4.5 - 0.0004*[double]$rtt - 0.1*[double]$loss),2)
  }

  $obj = [PSCustomObject]@{
    timestamp = $q.timestamp
    target    = $q.target
    rtt_ms    = $rtt
    jitter_ms = $jit
    loss_pct  = $loss
    mos       = $mos
    ssid      = $q.ssid
    bssid     = $bNorm
    ap_name   = $apName
    area      = $areaVal
    floor     = $floorVal
    ap_tag    = $tagVal
    role      = $roleVal
    label     = $labelVal
    segment   = $segVal
  }
  $qual += $obj
}

# ===== 集計（エリア / AP / 対象 / 役割）=====
$summaryRows = @()
$groups = $qual | Group-Object -Property area, ap_name, target, role, segment
foreach($g in $groups){
  $rtts = @($g.Group | Where-Object {$_.rtt_ms -ne $null}    | ForEach-Object {[double]$_.rtt_ms})
  $jits = @($g.Group | Where-Object {$_.jitter_ms -ne $null} | ForEach-Object {[double]$_.jitter_ms})
  $loss = @($g.Group | Where-Object {$_.loss_pct -ne $null}  | ForEach-Object {[double]$_.loss_pct})
  $mosv = @($g.Group | Where-Object {$_.mos -ne $null}       | ForEach-Object {[double]$_.mos})

  $rtt_med = $null; if($rtts.Count -gt 0){ $rtt_med = [math]::Round((Get-Median $rtts),1) }
  $rtt_p95 = $null; if($rtts.Count -gt 0){ $rtt_p95 = [math]::Round((Get-Percentile $rtts 0.95),1) }
  $jit_med = $null; if($jits.Count -gt 0){ $jit_med = [math]::Round((Get-Median $jits),1) }
  $loss_avg= $null; if($loss.Count -gt 0){ $loss_avg = [math]::Round(($loss | Measure-Object -Average | Select-Object -ExpandProperty Average),2) }
  $mos_med = $null; if($mosv.Count -gt 0){ $mos_med = [math]::Round((Get-Median $mosv),2) }

  $summaryRows += [PSCustomObject]@{
    area     = $g.Group[0].area
    ap_name  = $g.Group[0].ap_name
    target   = $g.Group[0].target
    role     = $g.Group[0].role
    segment  = $g.Group[0].segment
    count    = $g.Count
    rtt_med  = $rtt_med
    rtt_p95  = $rtt_p95
    jit_med  = $jit_med
    loss_avg = $loss_avg
    mos_med  = $mos_med
  }
}

# ===== HTML テンプレ生成（@' … '@ + JSON置換）=====
$summaryJson = $summaryRows | ConvertTo-Json -Depth 5

$htmlTemplate = @'
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<title>NetQuality Report (Targets-Filtered)</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
  body { font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Hiragino Kaku Gothic ProN","Noto Sans JP",sans-serif; margin: 16px; }
  h1 { font-size: 20px; margin: 0 0 12px; }
  .filters { display:flex; gap:8px; flex-wrap: wrap; margin: 8px 0 16px; }
  select, input, label { padding:6px; border:1px solid #ccc; border-radius: 8px; }
  label.chk { border:none; padding:0 6px 0 0; }
  table { border-collapse: collapse; width: 100%; margin: 8px 0 24px; }
  th, td { border-bottom: 1px solid #eee; padding: 8px; text-align: left; }
  th { background: #fafafa; position: sticky; top:0; z-index: 1; }
  .mono { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace; }
  .hint { color:#666; font-size:12px; }
  .legend { font-size:12px; color:#444; margin:4px 0 12px; }
  .legend span { margin-right:12px; padding:2px 6px; border-radius:6px; }
  .rt-ok   { background:#e7f7e7; }
  .rt-warn { background:#fff5e0; }
  .rt-bad  { background:#fdecec; }
  .loss-bad{ background:#fdecec; }
  .muted { color:#777; }
</style>
</head>
<body>
  <h1>ネットワーク品質レポート（targets.csv で宛先を絞り込み）</h1>
  <div class="hint">
    このレポートは <strong>targets.csv の key に含まれる target のみ</strong>を集計しています。floors.csv の BSSID から area/floor/tag を付与しています。
  </div>

  <div class="filters">
    <select id="areaSel"><option value="">(すべてのエリア)</option></select>
    <select id="apSel"><option value="">(すべてのAP)</option></select>
    <input id="targetSearch" placeholder="対象(部分一致)" />
    <select id="roleSel"><option value="">(すべての役割)</option></select>
    <select id="segSel"><option value="">(すべてのセグメント)</option></select>
  </div>

  <div class="legend">
    <span class="rt-ok">RTT &lt; 50ms</span>
    <span class="rt-warn">50–100ms</span>
    <span class="rt-bad">≥ 100ms</span>
    <span class="loss-bad">損失 &gt; 3%</span>
  </div>

  <table id="sumTbl">
    <thead>
      <tr>
        <th>エリア</th>
        <th>AP</th>
        <th>対象</th>
        <th>役割</th>
        <th>セグメント</th>
        <th>試行数</th>
        <th>RTT中央値</th>
        <th>RTT P95</th>
        <th>ジッタ中央値</th>
        <th>損失率(平均)</th>
        <th>MOS(中央値)</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>

<script>
  var summaryRows = __SUMMARY_JSON__;

  function uniq(vals){
    var out = [];
    for (var i=0;i<vals.length;i++){
      var x = vals[i];
      if (x && out.indexOf(x) === -1) { out.push(x); }
    }
    out.sort();
    return out;
  }
  function fillSel(el, opts){
    for (var i=0;i<opts.length;i++){
      var op = document.createElement('option');
      op.textContent = opts[i]; op.value = opts[i];
      el.appendChild(op);
    }
  }
  var areaSel = document.getElementById('areaSel');
  var apSel   = document.getElementById('apSel');
  var roleSel = document.getElementById('roleSel');
  var segSel  = document.getElementById('segSel');
  var qInput  = document.getElementById('targetSearch');

  fillSel(areaSel, uniq(summaryRows.map(function(r){return r.area;})));
  fillSel(apSel,   uniq(summaryRows.map(function(r){return r.ap_name;}).filter(function(x){return !!x;})));
  fillSel(roleSel, uniq(summaryRows.map(function(r){return r.role;})));
  fillSel(segSel,  uniq(summaryRows.map(function(r){return r.segment;})));

  function addInputHandler(el){ el.addEventListener('input', render); el.addEventListener('change', render); }
  addInputHandler(areaSel); addInputHandler(apSel); addInputHandler(roleSel); addInputHandler(segSel); addInputHandler(qInput);

  function colorRtt(td, val){
    if (val == null) return;
    if (val < 50) td.classList.add('rt-ok');
    else if (val < 100) td.classList.add('rt-warn');
    else td.classList.add('rt-bad');
  }
  function colorLoss(td, val){ if (val != null && val > 3) td.classList.add('loss-bad'); }

  function render(){
    var area = areaSel.value || "";
    var ap   = apSel.value || "";
    var role = roleSel.value || "";
    var seg  = segSel.value || "";
    var q    = (qInput.value || "").toLowerCase();

    var tbody = document.querySelector('#sumTbl tbody');
    tbody.innerHTML = '';

    var rows = summaryRows.slice().sort(function(a,b){
      var ka = (a.area||"") + "|" + (a.ap_name||"");
      var kb = (b.area||"") + "|" + (b.ap_name||"");
      if (ka < kb) return -1;
      if (ka > kb) return 1;
      var ra = -1; if (a.rtt_med != null) { ra = -a.rtt_med; }
      var rb = -1; if (b.rtt_med != null) { rb = -b.rtt_med; }
      return ra - rb;
    });

    for (var i=0;i<rows.length;i++){
      var r = rows[i];
      if (area && r.area !== area) continue;
      if (ap && (r.ap_name||"") !== ap) continue;
      if (role && r.role !== role) continue;
      if (seg && r.segment !== seg) continue;
      if (q && String(r.target||"").toLowerCase().indexOf(q) === -1) continue;

      var tr = document.createElement('tr');
      function td(t){ var e=document.createElement('td'); e.textContent = (t==null?"":t); return e; }

      tr.appendChild(td(r.area));
      tr.appendChild(td(r.ap_name||""));

      var ttd = td(r.target); ttd.classList.add('mono'); tr.appendChild(ttd);

      tr.appendChild(td(r.role||""));
      tr.appendChild(td(r.segment||""));
      tr.appendChild(td(r.count));

      var rttm = td(r.rtt_med); colorRtt(rttm, r.rtt_med); tr.appendChild(rttm);
      var rttp = td(r.rtt_p95); colorRtt(rttp, r.rtt_p95); tr.appendChild(rttp);
      tr.appendChild(td(r.jit_med));
      var loss = td(r.loss_avg); colorLoss(loss, r.loss_avg); tr.appendChild(loss);
      tr.appendChild(td(r.mos_med));

      tbody.appendChild(tr);
    }
  }

  render();
</script>
</body>
</html>
'@

$html = $htmlTemplate.Replace('__SUMMARY_JSON__', $summaryJson)

# ===== 出力 (UTF-8 BOM) =====
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

# ===== 参考メトリクス（数値のみ。実データは表示しません）=====
$total = $teams.Count
$after = $qual.Count
$mappedArea = ($qual | Where-Object { $_.area -ne "Unknown" }).Count
if ($total -gt 0) {
  $pctUsed = [math]::Round(100.0 * $after / $total, 1)
  $pctArea = 0.0
  if ($after -gt 0) { $pctArea = [math]::Round(100.0 * $mappedArea / $after, 1) }
  Write-Output ("teams行数: {0}, フィルタ後: {1} ({2}%)" -f $total,$after,$pctUsed)
  Write-Output ("area付与率（Unknown除外）: {0}% ({1}/{2})" -f $pctArea,$mappedArea,$after)
  Write-Output ("bssid空欄: {0} / ap_name空欄: {1}" -f $emptyBssid,$emptyAp)
}
