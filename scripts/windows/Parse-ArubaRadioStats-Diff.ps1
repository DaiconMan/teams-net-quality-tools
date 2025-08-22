<# 
.SYNOPSIS
  Aruba "show ap debug radio-stats" の2スナップショットから差分/秒を計算しCSV/HTML出力。
.DESCRIPTION
  - PowerShell 5.1 対応。三項演算子未使用。予約語 Host 不使用。
  - OneDrive/日本語/スペースを考慮（Join-Path / -LiteralPath 使用）。Cドライブ固定参照なし。
  - Before/After それぞれのテキスト出力を読み取り、以下を抽出:
      * Rx retry frames / RX CRC Errors / RX PLCP Errors（累積）
      * Channel Changes / TX Power Changes（累積）
      * Channel busy 1s / 4s / 64s (％)   ※瞬時値 → After 側を採用
      * Ch/Tx/Rx Time perct @ beacon intvl（末尾値）
      * CCA percentage of our bss / other bss / interference
  - 差分/秒、変更回数/時間を計算し CSV 出力。必要に応じて HTML レポートも出力。
  - 経過秒は "output time" を優先し自動算出（解析失敗時はファイル更新時刻差→既定900秒）。
.PARAMETER BeforeFile
  先に取得した radio-stats のテキストファイル。
.PARAMETER AfterFile
  後で取得した radio-stats のテキストファイル。
.PARAMETER DurationSec
  経過秒を明示的に上書きしたい場合のみ指定（通常は省略）。
.PARAMETER OutputCsv
  出力CSVのパス。未指定時は AfterFile と同一フォルダに自動命名。
.PARAMETER OutputHtml
  出力HTMLのパス。指定時は CSV に加えて HTML も生成。未指定時は作成しない。
.PARAMETER Title
  HTMLのタイトル文字列（未指定時は自動）。
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$BeforeFile,
  [Parameter(Mandatory=$true)][string]$AfterFile,
  [int]$DurationSec,
  [string]$OutputCsv,
  [string]$OutputHtml,
  [string]$Title
)

#--- 文字列HTMLエスケープ（System.Web未依存）
function HtmlEscape {
  param([string]$s)
  if ($null -eq $s) { return '' }
  $r = $s.Replace('&','&amp;')
  $r = $r.Replace('<','&lt;')
  $r = $r.Replace('>','&gt;')
  $r = $r.Replace('"','&quot;')
  $r = $r.Replace("'",'&#39;')
  return $r
}

#--- 共通: 文字列から最後の数値(整数/小数)を抜く
function Get-LastNumber {
  param([string]$Line)
  if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
  $matches = [regex]::Matches($Line, '(-?\d+(?:\.\d+)?)')
  if ($matches.Count -gt 0) { return [double]$matches[$matches.Count-1].Value }
  return $null
}

#--- 1行に複数の%が並ぶケース（例: 1s/4s/64s）用に、名前ごとに抽出
function TryExtractPercentTriplet {
  param([string]$Line, [ref]$Busy1s, [ref]$Busy4s, [ref]$Busy64s)
  $ok = $false
  $m1 = [regex]::Match($Line, '1s[^0-9\-]*(-?\d+(?:\.\d+)?)')
  if ($m1.Success) { $Busy1s.Value = [double]$m1.Groups[1].Value; $ok = $true }
  $m4 = [regex]::Match($Line, '4s[^0-9\-]*(-?\d+(?:\.\d+)?)')
  if ($m4.Success) { $Busy4s.Value = [double]$m4.Groups[1].Value; $ok = $true }
  $m64 = [regex]::Match($Line, '64s[^0-9\-]*(-?\d+(?:\.\d+)?)')
  if ($m64.Success){ $Busy64s.Value = [double]$m64.Groups[1].Value; $ok = $true }
  return $ok
}

#--- "output time" を各種表記から抽出
function Extract-OutputTime {
  param([string[]]$Lines)

  $dt = $null

  # 候補行
  $candidates = @()
  foreach ($raw in $Lines) {
    if ($raw -match '(?i)(output\s*time|出力(時刻|時間|日時)|生成時刻)') {
      $candidates += $raw
    }
  }
  if ($candidates.Count -eq 0) { return $null }

  # 1) Unix epoch(秒)
  foreach ($line in $candidates) {
    $m = [regex]::Match($line, '(?<!\d)(\d{10})(?:\.\d+)?(?!\d)')
    if ($m.Success) {
      try {
        $sec = [double]$m.Groups[1].Value
        $epoch = [DateTime]'1970-01-01 00:00:00'
        return $epoch.AddSeconds($sec)
      } catch {}
    }
  }

  # 2) ISO系
  foreach ($line in $candidates) {
    $m = [regex]::Match($line, '(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})')
    if ($m.Success) {
      try { return [DateTime]::Parse($m.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture) } catch {}
    }
  }

  # 3) スラッシュ
  foreach ($line in $candidates) {
    $m = [regex]::Match($line, '((\d{1,4})/(\d{1,2})/(\d{1,4})\s+\d{1,2}:\d{2}:\d{2})')
    if ($m.Success) {
      try { return [DateTime]::Parse($m.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture) } catch {}
      try { return [DateTime]::Parse($m.Groups[1].Value, [System.Globalization.CultureInfo]::GetCultureInfo('ja-JP')) } catch {}
      try { return [DateTime]::Parse($m.Groups[1].Value, [System.Globalization.CultureInfo]::GetCultureInfo('en-US')) } catch {}
    }
  }

  # 4) 英語月名
  foreach ($line in $candidates) {
    $m = [regex]::Match($line, '([A-Za-z]{3}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}(?:\s+\d{4})?)')
    if ($m.Success) {
      $s = $m.Groups[1].Value
      if ($s -notmatch '\s\d{4}$') {
        $year = (Get-Date).Year
        $s = "$s $year"
      }
      try { return [DateTime]::Parse($s, [System.Globalization.CultureInfo]::GetCultureInfo('en-US')) } catch {}
    }
  }

  return $null
}

#--- ファイル全体をパース（AP/Radio単位の辞書 + OutputTime を返す）
function Parse-RadioStatsFile {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
  $lines = Get-Content -LiteralPath $Path -Encoding UTF8

  $result = @{}
  $ap = ''
  $radio = ''
  $outTime = Extract-OutputTime -Lines $lines

  foreach ($raw in $lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    # AP名とRadio番号
    $m = [regex]::Match($line, '(?i)ap[-\s_]*name\s+([A-Za-z0-9_\-\.:]+)')
    if ($m.Success) { $ap = $m.Groups[1].Value }

    $mr = [regex]::Match($line, '(?i)\bradio\b[^0-9]*([01])')
    if ($mr.Success) { $radio = $mr.Groups[1].Value }

    $mh = [regex]::Match($line, '(?i)AP\s+([^\s]+).*?Radio[^0-9]*([01])')
    if ($mh.Success) {
      $ap = $mh.Groups[1].Value
      $radio = $mh.Groups[2].Value
    }

    # キー
    $key = ''
    if (-not [string]::IsNullOrWhiteSpace($ap)) {
      if (-not [string]::IsNullOrWhiteSpace($radio)) { $key = "$ap|$radio" } else { $key = "$ap|?" }
    } else {
      if (-not [string]::IsNullOrWhiteSpace($radio)) { $key = "Unknown|$radio" } else { $key = "Unknown|?" }
    }
    if (-not $result.ContainsKey($key)) {
      $obj = New-Object psobject -Property @{
        AP = $ap; Radio = $radio;
        RxRetry = $null; RxCRC = $null; RxPLCP = $null;
        ChannelChanges = $null; TxPowerChanges = $null;
        Busy1s = $null; Busy4s = $null; Busy64s = $null;
        BusyBeacon = $null; TxBeacon = $null; RxBeacon = $null;
        CCA_Our = $null; CCA_Other = $null; CCA_Interference = $null
      }
      $result[$key] = $obj
    }
    $cur = $result[$key]

    # 累積カウンタ
    if ($line -match '(?i)\bRx\s*retry\s*frames\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxRetry = [double]$v } }
    elseif ($line -match '(?i)\bRX?\s*CRC\b.*\bError') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxCRC = [double]$v } }
    elseif ($line -match '(?i)\bRX?\s*PLCP\b.*\bError') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxPLCP = [double]$v } }
    elseif ($line -match '(?i)\bChannel\s*Changes\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.ChannelChanges = [double]$v } }
    elseif ($line -match '(?i)\bTX\s*Power\s*Changes\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.TxPowerChanges = [double]$v } }

    # Busy 1s/4s/64s
    if ($line -match '(?i)\bChannel\s*busy\b') {
      $b1=$null;$b4=$null;$b64=$null
      $ok = TryExtractPercentTriplet -Line $line -Busy1s ([ref]$b1) -Busy4s ([ref]$b4) -Busy64s ([ref]$b64)
      if ($ok) {
        if ($b1 -ne $null) { $cur.Busy1s = [double]$b1 }
        if ($b4 -ne $null) { $cur.Busy4s = [double]$b4 }
        if ($b64 -ne $null){ $cur.Busy64s = [double]$b64 }
      }
    }

    # @ beacon interval
    if ($line -match '(?i)\bCh\s*Busy\s*perct\s*@\s*beacon') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.BusyBeacon = [double]$v } }
    elseif ($line -match '(?i)\bTx\s*Time\s*perct\s*@\s*beacon') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.TxBeacon = [double]$v } }
    elseif ($line -match '(?i)\bRx\s*Time\s*perct\s*@\s*beacon') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxBeacon = [double]$v } }

    # CCA breakdown
    if ($line -match '(?i)\bCCA\b.*\bour\b.*\bbss\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.CCA_Our = [double]$v } }
    elseif ($line -match '(?i)\bCCA\b.*\bother\b.*\bbss\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.CCA_Other = [double]$v } }
    elseif ($line -match '(?i)\bCCA\b.*\binterference\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.CCA_Interference = [double]$v } }
  }

  $ret = New-Object psobject -Property @{ Data = $result; OutputTime = $outTime }
  return $ret
}

#--- 差分の安全計算（負値→0）
function Diff-NonNegative { param([double]$After,[double]$Before)
  if ($After -eq $null -or $Before -eq $null) { return $null }
  $d = $After - $Before
  if ($d -lt 0) { return 0.0 }
  return $d
}

#--- 入口
$beforeObj = Parse-RadioStatsFile -Path $BeforeFile
$afterObj  = Parse-RadioStatsFile -Path $AfterFile

$before = $beforeObj.Data
$after  = $afterObj.Data
$beforeTime = $beforeObj.OutputTime
$afterTime  = $afterObj.OutputTime

# 経過秒の決定
if ($DurationSec -le 0) {
  $sec = 0
  if ($beforeTime -ne $null -and $afterTime -ne $null) {
    try { $sec = [int][Math]::Abs(($afterTime - $beforeTime).TotalSeconds) } catch { $sec = 0 }
  }
  if ($sec -le 0) {
    try {
      $t1 = [System.IO.File]::GetLastWriteTime($BeforeFile)
      $t2 = [System.IO.File]::GetLastWriteTime($AfterFile)
      $sec = [int]([Math]::Abs(($t2 - $t1).TotalSeconds))
    } catch { $sec = 0 }
  }
  if ($sec -le 0) { $sec = 900 }
  $DurationSec = $sec
}

# 出力先
if ([string]::IsNullOrWhiteSpace($OutputCsv)) {
  $outDir = Split-Path -LiteralPath $AfterFile -Parent
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $OutputCsv = Join-Path -LiteralPath $outDir -ChildPath ("aruba_radio_stats_diff_{0}.csv" -f $ts)
}

# ヘッダ
$header = @(
  'AP','Radio','DurationSec',
  'RxRetry_per_s','RxCRC_per_s','RxPLCP_per_s',
  'ChannelChanges_per_h','TxPowerChanges_per_h',
  'Busy1s_pct','Busy4s_pct','Busy64s_pct',
  'BusyBeacon_pct','TxBeacon_pct','RxBeacon_pct',
  'CCA_Our_pct','CCA_Other_pct','CCA_Interference_pct'
) -join ','

Set-Content -LiteralPath $OutputCsv -Value $header -Encoding UTF8

# HTML用に行データを保持
$rows = @()

# キーの和集合
$keys = New-Object System.Collections.Generic.HashSet[string]
foreach ($k in $before.Keys) { [void]$keys.Add($k) }
foreach ($k in $after.Keys)  { [void]$keys.Add($k) }

foreach ($k in $keys) {
  $b = $null; $a = $null
  if ($before.ContainsKey($k)) { $b = $before[$k] }
  if ($after.ContainsKey($k))  { $a = $after[$k] }

  $ap = ''; $radio = ''
  if ($a -ne $null) { $ap = $a.AP; $radio = $a.Radio }
  if ([string]::IsNullOrWhiteSpace($ap) -and $b -ne $null) { $ap = $b.AP }
  if ([string]::IsNullOrWhiteSpace($radio) -and $b -ne $null) { $radio = $b.Radio }

  # 差分
  $dRetry = Diff-NonNegative $a.RxRetry $b.RxRetry
  $dCRC   = Diff-NonNegative $a.RxCRC   $b.RxCRC
  $dPLCP  = Diff-NonNegative $a.RxPLCP  $b.RxPLCP
  $dChg   = Diff-NonNegative $a.ChannelChanges $b.ChannelChanges
  $dTxPw  = Diff-NonNegative $a.TxPowerChanges $b.TxPowerChanges

  $retry_ps = $null; $crc_ps = $null; $plcp_ps = $null; $chg_ph = $null; $txp_ph = $null
  if ($dRetry -ne $null) { $retry_ps = [Math]::Round($dRetry / $DurationSec, 6) }
  if ($dCRC   -ne $null) { $crc_ps   = [Math]::Round($dCRC   / $DurationSec, 6) }
  if ($dPLCP  -ne $null) { $plcp_ps  = [Math]::Round($dPLCP  / $DurationSec, 6) }
  if ($dChg   -ne $null) { $chg_ph   = [Math]::Round(($dChg   * 3600.0) / $DurationSec, 6) }
  if ($dTxPw  -ne $null) { $txp_ph   = [Math]::Round(($dTxPw  * 3600.0) / $DurationSec, 6) }

  # 瞬時系は After を優先、無ければ Before
  function Pick-AfterFirst { param($afterV,$beforeV)
    if ($afterV -ne $null) { return $afterV }
    return $beforeV
  }

  $busy1s = Pick-AfterFirst $a.Busy1s $b.Busy1s
  $busy4s = Pick-AfterFirst $a.Busy4s $b.Busy4s
  $busy64 = Pick-AfterFirst $a.Busy64s $b.Busy64s
  $busyB  = Pick-AfterFirst $a.BusyBeacon $b.BusyBeacon
  $txB    = Pick-AfterFirst $a.TxBeacon   $b.TxBeacon
  $rxB    = Pick-AfterFirst $a.RxBeacon   $b.RxBeacon
  $ccaO   = Pick-AfterFirst $a.CCA_Our $b.CCA_Our
  $ccaOt  = Pick-AfterFirst $a.CCA_Other $b.CCA_Other
  $ccaI   = Pick-AfterFirst $a.CCA_Interference $b.CCA_Interference

  # CSV行
  $vals = @(
    $ap, $radio, $DurationSec,
    $retry_ps, $crc_ps, $plcp_ps,
    $chg_ph, $txp_ph,
    $busy1s, $busy4s, $busy64,
    $busyB, $txB, $rxB,
    $ccaO, $ccaOt, $ccaI
  ) | ForEach-Object { if ($_ -eq $null) { '' } else { $_.ToString() } }

  $escaped = @()
  foreach ($v in $vals) {
    if ($v -match '[,"]') { $escaped += ('"{0}"' -f ($v -replace '"','""')) } else { $escaped += $v }
  }
  $line = ($escaped -join ',')
  Add-Content -LiteralPath $OutputCsv -Value $line -Encoding UTF8

  # HTML用
  $row = New-Object psobject -Property @{
    AP=$ap; Radio=$radio; DurationSec=$DurationSec;
    RxRetry_per_s=$retry_ps; RxCRC_per_s=$crc_ps; RxPLCP_per_s=$plcp_ps;
    ChannelChanges_per_h=$chg_ph; TxPowerChanges_per_h=$txp_ph;
    Busy1s_pct=$busy1s; Busy4s_pct=$busy4s; Busy64s_pct=$busy64;
    BusyBeacon_pct=$busyB; TxBeacon_pct=$txB; RxBeacon_pct=$rxB;
    CCA_Our_pct=$ccaO; CCA_Other_pct=$ccaOt; CCA_Interference_pct=$ccaI
  }
  $rows += $row
}

# ---- HTML 出力（任意） ----
if (-not [string]::IsNullOrWhiteSpace($OutputHtml)) {
  $outDir = Split-Path -LiteralPath $OutputHtml -Parent
  if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

  $titleText = $Title
  if ([string]::IsNullOrWhiteSpace($titleText)) {
    $bt = ''; $at = ''
    if ($beforeTime -ne $null) { $bt = $beforeTime.ToString('yyyy-MM-dd HH:mm:ss') }
    if ($afterTime  -ne $null) { $at = $afterTime.ToString('yyyy-MM-dd HH:mm:ss') }
    $titleText = "Aruba Radio Stats Diff ($bt → $at)"
  }

  # テーブルHTML組立
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine('<!DOCTYPE html>')
  [void]$sb.AppendLine('<meta charset="UTF-8">')
  [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
  [void]$sb.AppendLine("<title>{0}</title>" -f (HtmlEscape $titleText))
  [void]$sb.AppendLine('<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,"Noto Sans","Hiragino Kaku Gothic ProN","Yu Gothic",sans-serif;margin:16px}
h1{font-size:20px;margin:0 0 8px}
.small{color:#555;font-size:12px;margin-bottom:12px}
table{border-collapse:collapse;width:100%;table-layout:auto}
th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
th{background:#f7f7f7;position:sticky;top:0;cursor:pointer}
tr:nth-child(even){background:#fafafa}
.bad{background:#ffecec}
.note{font-size:12px;color:#333;margin-top:8px}
input[type="search"]{padding:6px 8px;width:280px;max-width:60%}
.kpi{display:inline-block;margin-right:16px;font-size:12px}
.tag{display:inline-block;border:1px solid #ddd;border-radius:3px;padding:2px 6px;margin-right:6px;background:#fafafa}
</style>')

  [void]$sb.AppendLine("<h1>{0}</h1>" -f (HtmlEscape $titleText))
  $btStr = ''; $atStr = ''
  if ($beforeTime -ne $null) { $btStr = $beforeTime.ToString('yyyy-MM-dd HH:mm:ss') }
  if ($afterTime  -ne $null) { $atStr = $afterTime.ToString('yyyy-MM-dd HH:mm:ss') }
  [void]$sb.AppendLine('<div class="small">Before: ' + (HtmlEscape $btStr) + ' / After: ' + (HtmlEscape $atStr) + ' / DurationSec: ' + $DurationSec + '</div>')
  [void]$sb.AppendLine('<div class="small"><span class="kpi">Busy = 空中占有（%）</span><span class="kpi">Retry/CRC/PLCP = 受信品質の悪化指標（/s）</span><span class="kpi">Channel/TX Power Changes = ARMの変更頻度（/h）</span></div>')
  [void]$sb.AppendLine('<div style="margin:10px 0"><input id="flt" type="search" placeholder="フィルタ（AP/数値を部分一致で抽出）..." oninput="filterTable()"></div>')

  # テーブルヘッダ
  $cols = @(
    'AP','Radio','DurationSec',
    'RxRetry_per_s','RxCRC_per_s','RxPLCP_per_s',
    'ChannelChanges_per_h','TxPowerChanges_per_h',
    'Busy1s_pct','Busy4s_pct','Busy64s_pct',
    'BusyBeacon_pct','TxBeacon_pct','RxBeacon_pct',
    'CCA_Our_pct','CCA_Other_pct','CCA_Interference_pct'
  )

  [void]$sb.AppendLine('<table id="tbl"><thead><tr>')
  foreach ($c in $cols) {
    [void]$sb.AppendLine("<th data-col=""$c"">$c</th>")
  }
  [void]$sb.AppendLine('</tr></thead><tbody>')

  foreach ($r in $rows) {
    [void]$sb.AppendLine('<tr>')
    foreach ($c in $cols) {
      $v = $r.PSObject.Properties[$c].Value
      $text = ''
      if ($null -ne $v) {
        if ($v -is [double] -or $v -is [single]) { $text = ([string]([Math]::Round([double]$v,6))) }
        else { $text = [string]$v }
      }
      $attr = ''
      if ($null -ne $v) {
        if ($v -is [double] -or $v -is [single] -or $v -is [int]) { $attr = ' data-val="'+([string]$v)+'"' }
      }
      [void]$sb.AppendLine('<td'+$attr+'>'+ (HtmlEscape $text) +'</td>')
    }
    [void]$sb.AppendLine('</tr>')
  }

  [void]$sb.AppendLine('</tbody></table>')
  [void]$sb.AppendLine('<div class="note">ヘッダーをクリックでソート。検索ボックスで部分一致フィルタ（複数列対象）。</div>')

  # ソート＆フィルタ JS（三項演算子不使用）
  [void]$sb.AppendLine('<script>
(function(){
  var lastCol=-1, asc=true;
  var tbl=document.getElementById("tbl");
  var ths=tbl.tHead.rows[0].cells;
  for(var i=0;i<ths.length;i++){
    (function(idx){
      ths[idx].addEventListener("click", function(){
        if(lastCol===idx){ asc=!asc; } else { lastCol=idx; asc=true; }
        sortBy(idx, asc);
      });
    })(i);
  }
  function getVal(td){
    var dv=td.getAttribute("data-val");
    if(dv!==null){ var n=parseFloat(dv); if(!isNaN(n)){ return {n:n, s:td.textContent}; } }
    return {n:null, s:td.textContent.toLowerCase()};
  }
  function cmp(a,b,ascFlag){
    if(a.n!==null && b.n!==null){
      if(a.n<b.n) return ascFlag? -1:1;
      if(a.n>b.n) return ascFlag? 1:-1;
      return 0;
    }
    if(a.s<b.s) return ascFlag? -1:1;
    if(a.s>b.s) return ascFlag? 1:-1;
    return 0;
  }
  function sortBy(col,ascFlag){
    var tbody=tbl.tBodies[0];
    var rows=[].slice.call(tbody.rows);
    rows.sort(function(r1,r2){
      var a=getVal(r1.cells[col]); var b=getVal(r2.cells[col]);
      return cmp(a,b,ascFlag);
    });
    for(var i=0;i<rows.length;i++){ tbody.appendChild(rows[i]); }
  }
  window.filterTable=function(){
    var q=document.getElementById("flt").value.toLowerCase();
    var trs=tbl.tBodies[0].rows;
    for(var i=0;i<trs.length;i++){
      var show=false;
      var tds=trs[i].cells;
      for(var j=0;j<tds.length;j++){
        var t=tds[j].textContent.toLowerCase();
        if(t.indexOf(q)>=0){ show=true; break; }
      }
      trs[i].style.display = show? "":"none";
    }
  };
})();
</script>')

  $html = $sb.ToString()
  Set-Content -LiteralPath $OutputHtml -Value $html -Encoding UTF8
}

Write-Output ("CSV : {0}" -f $OutputCsv)
if (-not [string]::IsNullOrWhiteSpace($OutputHtml)) { Write-Output ("HTML: {0}" -f $OutputHtml) }
exit 0