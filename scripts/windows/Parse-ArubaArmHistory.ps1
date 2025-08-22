<#
.SYNOPSIS
  Aruba "show ap arm history ap-name <AP名>" のテキスト出力(複数可)から、
  変更種別(チャネル/TxPower)×理由の件数を集計しCSV/HTML出力。
.DESCRIPTION
  - PS 5.1 対応。OneDrive/日本語パス考慮。Cドライブ非依存。三項演算子未使用。
  - 緩い検出で以下を抽出:
     * AP名: "ap-name <name>" または "AP <name>" 系
     * 変更種別: "channel ... change(d)" / "(tx|transmit) power ... change(d)"
     * 理由: "reason: <text>" / "reason=<text>" / "(reason <text>)" / "[reason: <text>]" 等
     * タイムスタンプらしき部分（あれば）
  - 2つのCSVを出力:
     1) summary: AP, ChangeType, Reason, Count
     2) events : AP, Timestamp, ChangeType, Reason, RawLine
  - -OutputHtml を指定した場合、上記2表を1つのHTMLにまとめて出力。
.PARAMETER InputFiles
  ARM履歴のテキストファイル（ワイルドカード可）。複数指定可。
.PARAMETER OutputDir
  CSV出力フォルダ。未指定なら最初の入力と同じ場所。
.PARAMETER OutputHtml
  HTML出力のパス。指定時のみ作成（CSVに加えて生成）。
.PARAMETER Title
  HTMLタイトル文字列。
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string[]]$InputFiles,
  [string]$OutputDir,
  [string]$OutputHtml,
  [string]$Title
)

function Ensure-Dir { param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

# HTMLエスケープ
function HtmlEscape {
  param([string]$s)
  if ($null -eq $s) { return '' }
  $r = $s.Replace('&','&amp;'); $r = $r.Replace('<','&lt;'); $r = $r.Replace('>','&gt;')
  $r = $r.Replace('"','&quot;'); $r = $r.Replace("'",'&#39;'); return $r
}

# AP名抽出
function Extract-APName { param([string]$line)
  $m = [regex]::Match($line, '(?i)ap[-\s_]*name\s+([A-Za-z0-9_\-\.:]+)'); if ($m.Success) { return $m.Groups[1].Value }
  $m2 = [regex]::Match($line, '(?i)\bAP\s+([^\s\]]+)'); if ($m2.Success) { return $m2.Groups[1].Value }
  return $null
}

function Detect-ChangeType { param([string]$line)
  if ($line -match '(?i)\bchannel\b.*\bchang') { return 'Channel' }
  if ($line -match '(?i)\b(tx|transmit)\b.*\bpower\b.*\bchang') { return 'TxPower' }
  return $null
}

function Extract-Reason { param([string]$line)
  $m = [regex]::Match($line, '(?i)reason\s*[:=]\s*([^\]\)]+)'); if ($m.Success) { return ($m.Groups[1].Value.Trim()) }
  $m2 = [regex]::Match($line, '(?i)[$begin:math:text$\\[]\\s*reason\\s+([^\\]$end:math:text$]+)[\)\]]'); if ($m2.Success) { return ($m2.Groups[1].Value.Trim()) }
  $candidates = @('interference','noise','load','dfs','coverage','utilization','error','roaming')
  foreach ($c in $candidates) { if ($line -match ("(?i)\b{0}\b" -f [regex]::Escape($c))) { return $c } }
  return 'Unknown'
}

function Extract-Timestamp { param([string]$line)
  $m = [regex]::Match($line, '((\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{2,4}|[A-Za-z]{3}\s+\d{1,2})\s+\d{1,2}:\d{2}:\d{2})')
  if ($m.Success) { return $m.Groups[1].Value }
  return ''
}

# 入力展開
$inputs = @()
foreach ($p in $InputFiles) {
  $matches = Get-ChildItem -File -LiteralPath $p -ErrorAction SilentlyContinue
  if (-not $matches) {
    $dir = Split-Path -LiteralPath $p -Parent; if ([string]::IsNullOrWhiteSpace($dir)) { $dir='.' }
    $name = Split-Path -LiteralPath $p -Leaf
    $matches = Get-ChildItem -File -Path (Join-Path -LiteralPath $dir -ChildPath $name) -ErrorAction SilentlyContinue
  }
  if ($matches) { $inputs += $matches }
}
if ($inputs.Count -eq 0) { throw "No input files found." }

if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = Split-Path -LiteralPath $inputs[0].FullName -Parent }
Ensure-Dir -Path $OutputDir

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$summaryCsv = Join-Path -LiteralPath $OutputDir -ChildPath ("arm_history_summary_{0}.csv" -f $ts)
$eventsCsv  = Join-Path -LiteralPath $OutputDir -ChildPath ("arm_history_events_{0}.csv"  -f $ts)

Set-Content -LiteralPath $summaryCsv -Value 'AP,ChangeType,Reason,Count' -Encoding UTF8
Set-Content -LiteralPath $eventsCsv  -Value 'AP,Timestamp,ChangeType,Reason,RawLine' -Encoding UTF8

$counts = @{}
$events = @()

foreach ($f in $inputs) {
  $ap = ''
  $lines = Get-Content -LiteralPath $f.FullName -Encoding UTF8
  foreach ($raw in $lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    $apFound = Extract-APName $line
    if ($apFound) { $ap = $apFound }

    $chg = Detect-ChangeType $line
    if ($chg) {
      $reason = Extract-Reason $line
      $tsStr  = Extract-Timestamp $line

      $useAP = $ap; if ([string]::IsNullOrWhiteSpace($useAP)) { $useAP = 'Unknown' }

      $vals = @($useAP,$tsStr,$chg,$reason,$line) | ForEach-Object {
        $v = $_; if ($v -eq $null) { $v = '' }
        if ($v -match '[,"]') { '"{0}"' -f ($v -replace '"','""') } else { $v }
      }
      Add-Content -LiteralPath $eventsCsv -Value ($vals -join ',') -Encoding UTF8

      $events += New-Object psobject -Property @{ AP=$useAP; Timestamp=$tsStr; ChangeType=$chg; Reason=$reason; RawLine=$line }

      $k = "{0}|{1}|{2}" -f $useAP,$chg,$reason
      if ($counts.ContainsKey($k)) { $counts[$k] = $counts[$k] + 1 } else { $counts[$k] = 1 }
    }
  }
}

# summary CSV
$keys = $counts.Keys | Sort-Object
foreach ($k in $keys) {
  $parts = $k -split '\|',3
  $ap = $parts[0]; $chg = $parts[1]; $reason = $parts[2]
  $cnt = $counts[$k]
  $vals = @($ap,$chg,$reason,$cnt) | ForEach-Object {
    if ($_ -match '[,"]') { '"{0}"' -f ($_ -replace '"','""') } else { $_ }
  }
  Add-Content -LiteralPath $summaryCsv -Value ($vals -join ',') -Encoding UTF8
}

Write-Output ("Summary CSV: {0}" -f $summaryCsv)
Write-Output ("Events  CSV: {0}" -f $eventsCsv)

# ---- HTML 出力（任意） ----
if (-not [string]::IsNullOrWhiteSpace($OutputHtml)) {
  $outDir = Split-Path -LiteralPath $OutputHtml -Parent
  if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

  $titleText = $Title
  if ([string]::IsNullOrWhiteSpace($titleText)) { $titleText = "Aruba ARM History Summary" }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine('<!DOCTYPE html>')
  [void]$sb.AppendLine('<meta charset="UTF-8">')
  [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
  [void]$sb.AppendLine("<title>{0}</title>" -f (HtmlEscape $titleText))
  [void]$sb.AppendLine('<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,"Noto Sans","Hiragino Kaku Gothic ProN","Yu Gothic",sans-serif;margin:16px}
h1{font-size:20px;margin:0 0 8px}
h2{font-size:16px;margin:16px 0 8px}
.small{color:#555;font-size:12px;margin-bottom:12px}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
th{background:#f7f7f7;position:sticky;top:0;cursor:pointer}
tr:nth-child(even){background:#fafafa}
input[type="search"]{padding:6px 8px;width:280px;max-width:60%}
.tag{display:inline-block;border:1px solid #ddd;border-radius:3px;padding:2px 6px;margin-right:6px;background:#fafafa}
</style>')
  [void]$sb.AppendLine("<h1>{0}</h1>" -f (HtmlEscape $titleText))
  [void]$sb.AppendLine('<div class="small">ヘッダークリックでソート／検索でフィルタ</div>')
  [void]$sb.AppendLine('<div style="margin:10px 0"><input id="flt" type="search" placeholder="フィルタ（AP/理由/種別/時刻を部分一致）..." oninput="filterTables()"></div>')

  # Summary テーブル
  [void]$sb.AppendLine('<h2>Summary（理由別集計）</h2>')
  [void]$sb.AppendLine('<table id="sum"><thead><tr><th>AP</th><th>ChangeType</th><th>Reason</th><th>Count</th></tr></thead><tbody>')
  foreach ($k in $keys) {
    $parts = $k -split '\|',3
    $ap = $parts[0]; $chg = $parts[1]; $reason = $parts[2]; $cnt = $counts[$k]
    [void]$sb.AppendLine('<tr><td>'+ (HtmlEscape $ap) +'</td><td>'+ (HtmlEscape $chg) +'</td><td>'+ (HtmlEscape $reason) +'</td><td data-val="'+$cnt+'">'+$cnt+'</td></tr>')
  }
  [void]$sb.AppendLine('</tbody></table>')

  # Events テーブル
  [void]$sb.AppendLine('<h2>Events（明細）</h2>')
  [void]$sb.AppendLine('<table id="evt"><thead><tr><th>AP</th><th>Timestamp</th><th>ChangeType</th><th>Reason</th><th>RawLine</th></tr></thead><tbody>')
  foreach ($e in $events) {
    $ap = HtmlEscape $e.AP
    $ts2 = HtmlEscape $e.Timestamp
    $chg = HtmlEscape $e.ChangeType
    $rsn = HtmlEscape $e.Reason
    $raw = HtmlEscape $e.RawLine
    [void]$sb.AppendLine('<tr><td>'+ $ap +'</td><td>'+ $ts2 +'</td><td>'+ $chg +'</td><td>'+ $rsn +'</td><td>'+ $raw +'</td></tr>')
  }
  [void]$sb.AppendLine('</tbody></table>')

  # 共通のソート＆フィルタ JS
  [void]$sb.AppendLine('<script>
(function(){
  function attachSort(tblId){
    var tbl=document.getElementById(tblId);
    if(!tbl || !tbl.tHead) return;
    var lastCol=-1, asc=true;
    var ths=tbl.tHead.rows[0].cells;
    for(var i=0;i<ths.length;i++){
      (function(idx){
        ths[idx].addEventListener("click", function(){
          if(lastCol===idx){ asc=!asc; } else { lastCol=idx; asc=true; }
          sortBy(tbl, idx, asc);
        });
      })(i);
    }
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
  function sortBy(tbl,col,ascFlag){
    var tbody=tbl.tBodies[0];
    var rows=[].slice.call(tbody.rows);
    rows.sort(function(r1,r2){
      var a=getVal(r1.cells[col]); var b=getVal(r2.cells[col]);
      return cmp(a,b,ascFlag);
    });
    for(var i=0;i<rows.length;i++){ tbody.appendChild(rows[i]); }
  }
  window.filterTables=function(){
    var q=document.getElementById("flt").value.toLowerCase();
    var ids=["sum","evt"];
    for(var x=0;x<ids.length;x++){
      var tbl=document.getElementById(ids[x]); if(!tbl) continue;
      var trs=tbl.tBodies[0].rows;
      for(var i=0;i<trs.length;i++){
        var show=false, tds=trs[i].cells;
        for(var j=0;j<tds.length;j++){
          var t=tds[j].textContent.toLowerCase();
          if(t.indexOf(q)>=0){ show=true; break; }
        }
        trs[i].style.display = show? "":"none";
      }
    }
  };
  attachSort("sum"); attachSort("evt");
})();
</script>')
  $html = $sb.ToString()
  Set-Content -LiteralPath $OutputHtml -Value $html -Encoding UTF8
  Write-Output ("HTML: {0}" -f $OutputHtml)
}

exit 0