<#
  Generate-NetQuality-HTMLReport.ps1 (PS 5.1 Compatible)
  - 入力: teams_net_quality.csv（Measure-NetQuality-WithHops.ps1 の品質CSV）
  - 補正: floor.csv (BSSID→エリア/階/tag), node_roles.csv (IP/FQDN→役割/ラベル/セグメント)
  - 出力: HTML 単一ファイル（外部ライブラリ不要）
  - 仕様: path_hop_quality.csv は解析しない（ZscalerによりInternet/SaaS向けpingは信頼しない想定）
  - 既定で SaaS / Internet は一覧から除外（HTML内のトグルで含め可能）
  - 注意: PowerShellの $Host は未使用。CSVの列名 host は target として扱う。

  使い方例:
    powershell -ExecutionPolicy Bypass -File .\Generate-NetQuality-HTMLReport.ps1 `
      -QualityCsv "$env:LOCALAPPDATA\TeamsNet\teams_net_quality.csv" `
      -BssidFloorCsv ".\floor.csv" -NodeRoleCsv ".\node_roles.csv" `
      -OutHtml ".\NetQuality-Report.html"
#>

[CmdletBinding()]
param(
  [string]$QualityCsv    = (Join-Path $env:LOCALAPPDATA "TeamsNet\teams_net_quality.csv"),
  [string]$BssidFloorCsv = ".\floor.csv",
  [string]$NodeRoleCsv   = ".\node_roles.csv",
  [string]$OutHtml       = ".\NetQuality-Report.html"
)

# ===== ヘルパ =====
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

function Get-Median {
  param([double[]]$arr)
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  if ($n % 2 -eq 1) { return [double]$s[[int][math]::Floor($n/2)] }
  else {
    $a = [double]$s[$n/2 - 1]; $b = [double]$s[$n/2]
    return ($a + $b) / 2.0
  }
}

function Get-Percentile {
  param([double[]]$arr, [double]$p) # 0.95 等
  if (-not $arr -or $arr.Count -eq 0) { return $null }
  $s = $arr | Sort-Object
  $n = $s.Count
  $k = ($n - 1) * $p
  $f = [math]::Floor($k); $c = [math]::Ceiling($k)
  if ($f -eq $c) { return [double]$s[[int]$k] }
  [double]$sf = $s[[int]$f]; [double]$sc = $s[[int]$c]
  return $sf + ($k - $f) * ($sc - $sf)
}

function Safe-Lower {
  param([string]$s)
  if ($null -eq $s) { return "" }
  return ($s.ToString().Trim().ToLower())
}

# ===== CSV 読み込み =====
if (-not (Test-Path $QualityCsv)) { Write-Error "QualityCsv が見つかりません: $QualityCsv"; exit 1 }
$qualRaw = Import-Csv -Path $QualityCsv

$floorMap = @()
if (Test-Path $BssidFloorCsv) { $floorMap = Import-Csv -Path $BssidFloorCsv }

$nodeRoles = @()
if (Test-Path $NodeRoleCsv) { $nodeRoles = Import-Csv -Path $NodeRoleCsv }

# BSSID→Area/Floor/Tag 辞書
$areaByBssid  = @{}
$floorByBssid = @{}
$tagByBssid   = @{}
foreach($r in $floorMap){
  $b = Safe-Lower $r.bssid
  if (-not $areaByBssid.ContainsKey($b))  { $areaByBssid[$b]  = $r.area }
  if ($r.PSObject.Properties.Name -contains 'floor') { $floorByBssid[$b] = $r.floor }
  if ($r.PSObject.Properties.Name -contains 'tag')   { $tagByBssid[$b]   = $r.tag }
}

# ノード（IP/FQDN）→役割/ラベル/セグメント
$roleByNode    = @{}
$labelByNode   = @{}
$segmentByNode = @{}
foreach($r in $nodeRoles){
  $k = Safe-Lower $r.ip_or_host
  $roleByNode[$k] = $r.role
  if ($r.PSObject.Properties.Name -contains 'label')   { $labelByNode[$k]   = $r.label }
  if ($r.PSObject.Properties.Name -contains 'segment') { $segmentByNode[$k] = $r.segment }
}

# ===== 正規化（teams_net_quality）=====
# 想定列: timestamp, host(対象), icmp_avg_ms, icmp_jitter_ms, loss_pct, mos, ssid, bssid, ap_name
$qual = @()
foreach($q in $qualRaw){
  $bssidNorm = Safe-Lower $q.bssid
  $targetTxt = $q.host

  # area/floor/tag を辞書から取得（PS5.1のため三項演算子は使わない）
  $areaVal = "Unknown"
  if ($areaByBssid.ContainsKey($bssidNorm)) { $areaVal = $areaByBssid[$bssidNorm] }

  $floorVal = $null
  if ($floorByBssid.ContainsKey($bssidNorm)) { $floorVal = $floorByBssid[$bssidNorm] }

  $tagVal = $null
  if ($tagByBssid.ContainsKey($bssidNorm)) { $tagVal = $tagByBssid[$bssidNorm] }

  $obj = [PSCustomObject]@{
    timestamp = $q.timestamp
    target    = $targetTxt
    rtt_ms    = Parse-Double $q.icmp_avg_ms
    jitter_ms = Parse-Double $q.icmp_jitter_ms
    loss_pct  = Parse-Double $q.loss_pct
    mos       = Parse-Double $q.mos
    ssid      = $q.ssid
    bssid     = $bssidNorm
    ap_name   = $q.ap_name
    area      = $areaVal
    floor     = $floorVal
    ap_tag    = $tagVal
  }

  # MOS 未計算時は簡易推定
  if ($null -eq $obj.mos -and $null -ne $obj.rtt_ms -and $null -ne $obj.loss_pct) {
    $obj.mos = [math]::Round((4.5 - 0.0004*[double]$obj.rtt_ms - 0.1*[double]$obj.loss_pct),2)
  }

  # 役割・ラベル・セグメント付与（三項演算子は使わず if で代入）
  $nk = Safe-Lower $obj.target

  $roleVal = "Uncategorized"
  if ($roleByNode.ContainsKey($nk)) { $roleVal = $roleByNode[$nk] }
  $obj | Add-Member -NotePropertyName role -NotePropertyValue $roleVal

  $labelVal = $obj.target
  if ($labelByNode.ContainsKey($nk)) { $labelVal = $labelByNode[$nk] }
  $obj | Add-Member -NotePropertyName label -NotePropertyValue $labelVal

  $segVal = ""
  if ($segmentByNode.ContainsKey($nk)) { $segVal = $segmentByNode[$nk] }
  $obj | Add-Member -NotePropertyName segment -NotePropertyValue $segVal

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

# ===== HTML 生成 =====
$summaryJson = $summaryRows | ConvertTo-Json -Depth 5

$html = @"
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<title>NetQuality Report (No-Hop / Internal Focus)</title>
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
  .role-badge { display:inline-block; padding: 2px 8px; border-radius: 999px; background:#eee; font-size: 12px; }
  .muted { color:#777; }
</style>
</head>
<body>
  <h1>ネットワーク品質レポート（Hop解析なし／内部重視）</h1>
  <div class="hint">
    このレポートは <strong>path_hop_quality.csv を使用しません</strong>。Zscaler経由でのInternet/SaaS向けpingは信頼しない想定のため、既定で表示から除外します。
  </div>

  <div class="filters">
    <select id="areaSel"><option value="">(すべてのエリア)</option></select>
    <select id="apSel"><option value="">(すべてのAP)</option></select>
    <input id="targetSearch" placeholder="対象(部分一致)" />
    <select id="roleSel"><option value="">(すべての役割)</option></select>
    <select id="segSel"><option value="">(すべてのセグメント)</option></select>
    <label class="chk"><input type="checkbox" id="includeExternal"> 外部(SaaS/Internet)を含める</label>
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
  // ====== データ埋め込み ======
  var summaryRows = $summaryJson;

  // ====== UI初期化 ======
  function uniq(vals){
    var out = [];
    for (var i=0;i<vals.length;i++){
      var x = vals[i];
      if (x) {
        if (out.indexOf(x) === -1) { out.push(x); }
      }
    }
    out.sort();
    return out;
  }
  function fillSel(el, opts){
    for (var i=0;i<opts.length;i++){
      var op = document.createElement('option');
      op.textContent = opts[i];
      op.value = opts[i];
      el.appendChild(op);
    }
  }
  var areaSel = document.getElementById('areaSel');
  var apSel   = document.getElementById('apSel');
  var roleSel = document.getElementById('roleSel');
  var segSel  = document.getElementById('segSel');
  var qInput  = document.getElementById('targetSearch');
  var includeExternal = document.getElementById('includeExternal');

  fillSel(areaSel, uniq(summaryRows.map(function(r){return r.area;})));
  fillSel(apSel,   uniq(summaryRows.map(function(r){return r.ap_name;}).filter(function(x){return !!x;})));
  fillSel(roleSel, uniq(summaryRows.map(function(r){return r.role;})));
  fillSel(segSel,  uniq(summaryRows.map(function(r){return r.segment;})));

  function addInputHandler(el){
    el.addEventListener('input', render);
    el.addEventListener('change', render);
  }
  addInputHandler(areaSel); addInputHandler(apSel); addInputHandler(roleSel);
  addInputHandler(segSel);  addInputHandler(qInput); addInputHandler(includeExternal);

  function colorRtt(td, val){
    if (val == null) return;
    if (val < 50) td.classList.add('rt-ok');
    else if (val < 100) td.classList.add('rt-warn');
    else td.classList.add('rt-bad');
  }
  function colorLoss(td, val){
    if (val == null) return;
    if (val > 3) td.classList.add('loss-bad');
  }
  function isExternalRole(role){
    if (!role) return false;
    var r = String(role).toLowerCase();
    if (r === 'saas') return true;
    if (r === 'internet') return true;
    return false;
  }

  function render(){
    var area = areaSel.value || "";
    var ap   = apSel.value || "";
    var role = roleSel.value || "";
    var seg  = segSel.value || "";
    var q    = (qInput.value || "").toLowerCase();
    var showExt = includeExternal.checked;

    var tbody = document.querySelector('#sumTbl tbody');
    tbody.innerHTML = '';

    // sort: area -> ap -> worst rtt_med desc （三項演算子は使わない）
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
      if (!showExt && isExternalRole(r.role)) continue;
      if (q && String(r.target).toLowerCase().indexOf(q) === -1) continue;

      var tr = document.createElement('tr');
      function td(t){ var e=document.createElement('td'); e.textContent = (t==null?"":t); return e; }

      tr.appendChild(td(r.area));
      tr.appendChild(td(r.ap_name||""));

      var ttd = td(r.target); ttd.classList.add('mono'); tr.appendChild(ttd);

      var rtd = document.createElement('td');
      rtd.textContent = (r.role||"");
      if (isExternalRole(r.role)) { rtd.classList.add('muted'); }
      tr.appendChild(rtd);

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

  // 既定で外部は除外
  includeExternal.checked = false;
  render();
</script>
</body>
</html>
"@

# 出力
$html | Out-File -FilePath $OutHtml -Encoding utf8
Write-Output "HTMLレポートを出力しました: $OutHtml"
