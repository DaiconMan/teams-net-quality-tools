<# 
.SYNOPSIS
  Aruba "show ap debug radio-stats" の2スナップショットから差分/秒を計算しCSV/HTML出力。
  HTMLには簡易の原因判定（Diagnosis）を表示。

.DESCRIPTION
  - PowerShell 5.1 対応。三項演算子未使用。予約語 Host 不使用。
  - OneDrive/日本語/スペース対応（Split-Path/Join-Pathの扱いを安全化）。Cドライブ固定参照なし。
  - Before/After のテキスト出力から次を抽出:
      * Rx retry frames / RX CRC Errors / RX PLCP Errors（累積 → 差分/秒）
      * Channel Changes / TX Power Changes（累積 → /時間）
      * Channel busy 1s / 4s / 64s（％, After優先）
      * Ch/Tx/Rx Time perct @ beacon intvl（After優先）
      * CCA percentage of our bss / other bss / interference（After優先）
  - 経過秒は "output time" を優先し自動算出（失敗時はファイル更新時刻差→既定900秒）。
  - HTML出力時は診断列（Diagnosis）を追加。行の背景色で注意度も表示。
  - 解析時の誤検知対策：
      * 「コマンド行（show ap debug radio-stats 0）」を AP/Radio と誤認しないよう正規表現を厳格化
      * メトリクス行を検出した時のみオブジェクトを生成（空行の出力抑止）

.PARAMETER BeforeFile
  先に取得した radio-stats のテキストファイル。
.PARAMETER AfterFile
  後で取得した radio-stats のテキストファイル。
.PARAMETER DurationSec
  経過秒を明示的に上書きしたい場合のみ指定（通常は省略）。
.PARAMETER OutputCsv
  出力CSVのパス。未指定時は AfterFile と同一フォルダ（不可ならカレント）に自動命名。
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

# -------------------- ユーザ調整しやすい診断しきい値 --------------------
# （必要ならここだけ書き換え）
$TH_BusyHigh   = 60     # Busy64s がこの%を超えると「高占有」
$TH_BusyWarn   = 40     # Busy64s がこの%を超えると「注意」
$TH_ExcessCCA  = 15     # BusyBeacon - (Tx+Rx) がこのpt超だと CCA影響大と判断
$TH_OtherHigh  = 20     # CCA_Other % がこのpt超で同チャネル影響大
$TH_InterfHigh = 10     # CCA_Interference % がこのpt超で非Wi-Fi干渉疑い
$TH_RetryHigh  = 5      # RxRetry_per_s がこの値超で再送多い
$TH_CRCHigh    = 100    # RxCRC_per_s がこの値超でCRC多い
$TH_PLCPHigh   = 50     # RxPLCP_per_s がこの値超でPLCP多い
$TH_ChgHigh    = 3      # ChannelChanges_per_h 高い目安
$TH_TxPHigh    = 10     # TxPowerChanges_per_h 高い目安

# -------------------- 安全なパス操作 --------------------
function Get-ParentOrCwd {
  param([string]$PathLike)
  $dir = $null
  if (-not [string]::IsNullOrWhiteSpace($PathLike)) {
    if (Test-Path -LiteralPath $PathLike) {
      try { $dir = Split-Path -LiteralPath $PathLike -Parent } catch { $dir = $null }
    } else {
      try { $dir = Split-Path -Path $PathLike -Parent } catch { $dir = $null }
    }
  }
  if ([string]::IsNullOrWhiteSpace($dir)) {
    try { return (Get-Location).Path } catch { return "." }
  }
  return $dir
}

# -------------------- ユーティリティ --------------------
function HtmlEscape {
  param([string]$s)
  if ($null -eq $s) { return '' }
  $r = $s.Replace('&','&amp;'); $r = $r.Replace('<','&lt;'); $r = $r.Replace('>','&gt;')
  $r = $r.Replace('"','&quot;'); $r = $r.Replace("'",'&#39;'); return $r
}

function Get-LastNumber {
  param([string]$Line)
  if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
  $matches = [regex]::Matches($Line, '(-?\d+(?:\.\d+)?)')
  if ($matches.Count -gt 0) { return [double]$matches[$matches.Count-1].Value }
  return $null
}

function TryExtractPercentTriplet {
  param([string]$Line, [ref]$Busy1s, [ref]$Busy4s, [ref]$Busy64s)
  $ok = $false
  $m1 = [regex]::Match($Line, '\b1s\b[^0-9\-]*(-?\d+(?:\.\d+)?)')
  if ($m1.Success) { $Busy1s.Value = [double]$m1.Groups[1].Value; $ok = $true }
  $m4 = [regex]::Match($Line, '\b4s\b[^0-9\-]*(-?\d+(?:\.\d+)?)')
  if ($m4.Success) { $Busy4s.Value = [double]$m4.Groups[1].Value; $ok = $true }
  $m64 = [regex]::Match($Line, '\b64s\b[^0-9\-]*(-?\d+(?:\.\d+)?)')
  if ($m64.Success){ $Busy64s.Value = [double]$m64.Groups[1].Value; $ok = $true }
  return $ok
}

function Extract-OutputTime {
  param([string[]]$Lines)
  $candidates = @()
  foreach ($raw in $Lines) {
    if ($raw -match '(?i)(output\s*time|出力(時刻|時間|日時)|生成時刻)') { $candidates += $raw }
  }
  if ($candidates.Count -eq 0) { return $null }

  foreach ($line in $candidates) {
    $m = [regex]::Match($line, '(?<!\d)(\d{10})(?:\.\d+)?(?!\d)')
    if ($m.Success) {
      try { return ([DateTime]'1970-01-01').AddSeconds([double]$m.Groups[1].Value) } catch {}
    }
  }
  foreach ($line in $candidates) {
    $m = [regex]::Match($line, '(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})')
    if ($m.Success) { try { return [DateTime]::Parse($m.Groups[1].Value,[System.Globalization.CultureInfo]::InvariantCulture) } catch {} }
  }
  foreach ($line in $candidates) {
    $m = [regex]::Match($line, '((\d{1,4})/(\d{1,2})/(\d{1,4})\s+\d{1,2}:\d{2}:\d{2})')
    if ($m.Success) {
      try { return [DateTime]::Parse($m.Groups[1].Value,[System.Globalization.CultureInfo]::InvariantCulture) } catch {}
      try { return [DateTime]::Parse($m.Groups[1].Value,[System.Globalization.CultureInfo]::GetCultureInfo('ja-JP')) } catch {}
      try { return [DateTime]::Parse($m.Groups[1].Value,[System.Globalization.CultureInfo]::GetCultureInfo('en-US')) } catch {}
    }
  }
  foreach ($line in $candidates) {
    $m = [regex]::Match($line, '([A-Za-z]{3}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}(?:\s+\d{4})?)')
    if ($m.Success) {
      $s = $m.Groups[1].Value
      if ($s -notmatch '\s\d{4}$') { $s = "$s " + (Get-Date).Year }
      try { return [DateTime]::Parse($s,[System.Globalization.CultureInfo]::GetCultureInfo('en-US')) } catch {}
    }
  }
  return $null
}

# ---- 診断（タグ配列と重大度クラスを返す） ----
function Build-Diagnosis {
  param(
    [double]$Busy64,[double]$BusyB,[double]$TxB,[double]$RxB,
    [double]$CCAO,[double]$CCAI,[double]$Retry,[double]$CRC,[double]$PLCP,
    [double]$ChgPH,[double]$TxPPH
  )
  $tags = New-Object System.Collections.Generic.List[string]
  $severity = 'sev-ok'  # 'sev-warn' / 'sev-crit'

  # 1) 同チャネル過密（他BSS優勢）
  $excess = $null
  if ($BusyB -ne $null -and $TxB -ne $null -and $RxB -ne $null) { $excess = $BusyB - ($TxB + $RxB) }
  if ($excess -ne $null -and $CCAO -ne $null) {
    if ($excess -ge $TH_ExcessCCA -and $CCAO -ge $TH_OtherHigh) {
      $tags.Add("Co-channel (Other=" + [int]$CCAO + "%, ΔCCA=" + [int]$excess + "pt)")
      $severity = 'sev-warn'
    }
  }
  if ($Busy64 -ne $null -and $Busy64 -ge $TH_BusyHigh) {
    $tags.Add("High Busy (" + [int]$Busy64 + "%)")
    if ($severity -eq 'sev-warn') { $severity = 'sev-crit' } else { $severity = 'sev-warn' }
  }

  # 2) 非Wi-Fi干渉（PLCP優勢＋Interference）
  if ($CCAI -ne $null -and $CCAI -ge $TH_InterfHigh -and $PLCP -ne $null -and $PLCP -ge $TH_PLCPHigh) {
    $tags.Add("Non-Wi-Fi interference (Interf=" + [int]$CCAI + "%, PLCP/s=" + [int]$PLCP + ")")
    if ($severity -eq 'sev-warn') { $severity = 'sev-crit' } else { $severity = 'sev-warn' }
  }

  # 3) 品質（SNR/隠れ端末）
  if ($Busy64 -ne $null -and $Busy64 -le $TH_BusyWarn) {
    if ( ($Retry -ne $null -and $Retry -ge $TH_RetryHigh) -or ($CRC -ne $null -and $CRC -ge $TH_CRCHigh) ) {
      $tags.Add("Quality/SNR or hidden-node (Retry/CRC high)")
      if ($severity -eq 'sev-ok') { $severity = 'sev-warn' }
    }
  }

  # 4) ARMフラップ
  if (($ChgPH -ne $null -and $ChgPH -ge $TH_ChgHigh) -or ($TxPPH -ne $null -and $TxPPH -ge $TH_TxPHigh)) {
    $tags.Add("ARM flapping (Chg/TxP)")
    if ($severity -eq 'sev-ok') { $severity = 'sev-warn' }
  }

  if ($tags.Count -eq 0) { $tags.Add("Looks normal") }
  return ,@($tags,$severity)
}

function Diff-NonNegative { param([double]$After,[double]$Before)
  if ($After -eq $null -or $Before -eq $null) { return $null }
  $d = $After - $Before
  if ($d -lt 0) { return 0.0 }
  return $d
}

# ---- ファイルパース（誤認識抑止版） ----
function Parse-RadioStatsFile {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
  $lines = Get-Content -LiteralPath $Path -Encoding UTF8

  $outTime = Extract-OutputTime -Lines $lines

  $result = @{}
  $ap = ''
  $radio = ''

  foreach ($raw in $lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    # --- AP名とRadio番号の検出を厳格化 ---
    # ap-name <AP> のみ素直に採用
    $mApName = [regex]::Match($line, '(?i)\bap[-\s_]*name\s+([A-Za-z0-9_\-\.:]+)\b')
    if ($mApName.Success) { $ap = $mApName.Groups[1].Value }

    # "AP <name> ... Radio 0/1" の行のみ（先頭"AP "から始まる行に限定、"radio-stats"は除外）
    $mHdr = [regex]::Match($line, '^(?i)\s*AP\s+([^\s]+).*?\bRadio\s+([01])\b')
    if ($mHdr.Success) {
      $ap    = $mHdr.Groups[1].Value
      $radio = $mHdr.Groups[2].Value
    }

    # "radio-stats 0" を拾わないよう、"radio <digit>" のみ許可
    $mRadio = [regex]::Match($line, '(?i)\bradio\s+([01])\b')
    if ($mRadio.Success) { $radio = $mRadio.Groups[1].Value }

    # --- この行がメトリクスかどうか判定（メトリクス行のみオブジェクト生成） ---
    $isMetric = $false
    if ($line -match '(?i)\bRx\s*retry\s*frames\b') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bRX?\s*CRC\b.*\bError') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bRX?\s*PLCP\b.*\bError') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bChannel\s*Changes\b') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bTX\s*Power\s*Changes\b') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bChannel\s*busy\b') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\b(Ch|Tx|Rx)\s*Time\s*perct\s*@\s*beacon') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bCCA\b.*\b(bss|interference)\b') { $isMetric = $true }

    if (-not $isMetric) { continue }

    # キー（AP/Radioの無い行は Unknown に）
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

    # --- 値の抽出 ---
    if ($line -match '(?i)\bRx\s*retry\s*frames\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxRetry = [double]$v } }
    elseif ($line -match '(?i)\bRX?\s*CRC\b.*\bError') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxCRC = [double]$v } }
    elseif ($line -match '(?i)\bRX?\s*PLCP\b.*\bError') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxPLCP = [double]$v } }
    elseif ($line -match '(?i)\bChannel\s*Changes\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.ChannelChanges = [double]$v } }
    elseif ($line -match '(?i)\bTX\s*Power\s*Changes\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.TxPowerChanges = [double]$v } }

    if ($line -match '(?i)\bChannel\s*busy\b') {
      $b1=$null;$b4=$null;$b64=$null
      $ok = TryExtractPercentTriplet -Line $line -Busy1s ([ref]$b1) -Busy4s ([ref]$b4) -Busy64s ([ref]$b64)
      if ($ok) {
        if ($b1 -ne $null) { $cur.Busy1s  = [double]$b1 }
        if ($b4 -ne $null) { $cur.Busy4s  = [double]$b4 }
        if ($b64 -ne $null){ $cur.Busy64s = [double]$b64 }
      }
    }

    if     ($line -match '(?i)\bCh\s*Busy\s*perct\s*@\s*beacon') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.BusyBeacon = [double]$v } }
    elseif ($line -match '(?i)\bTx\s*Time\s*perct\s*@\s*beacon') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.TxBeacon   = [double]$v } }
    elseif ($line -match '(?i)\bRx\s*Time\s*perct\s*@\s*beacon') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxBeacon   = [double]$v } }

    if     ($line -match '(?i)\bCCA\b.*\bour\b.*\bbss\b')       { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.CCA_Our            = [double]$v } }
    elseif ($line -match '(?i)\bCCA\b.*\bother\b.*\bbss\b')     { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.CCA_Other          = [double]$v } }
    elseif ($line -match '(?i)\bCCA\b.*\binterference\b')       { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.CCA_Interference   = [double]$v } }
  }

  $ret = New-Object psobject -Property @{ Data = $result; OutputTime = $outTime }
  return $ret
}

# -------------------- メイン処理 --------------------
$beforeObj = Parse-RadioStatsFile -Path $BeforeFile
$afterObj  = Parse-RadioStatsFile -Path $AfterFile

$before = $beforeObj.Data
$after  = $afterObj.Data
$beforeTime = $beforeObj.OutputTime
$afterTime  = $afterObj.OutputTime

# 経過秒の決定（引数 > output time > ファイル時刻 > 既定900）
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

# CSV出力先
if ([string]::IsNullOrWhiteSpace($OutputCsv)) {
  $outDir = Get-ParentOrCwd -PathLike $AfterFile
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $OutputCsv = Join-Path -Path $outDir -ChildPath ("aruba_radio_stats_diff_{0}.csv" -f $ts)
}

# CSVヘッダ（Diagnosis列を追加）
$header = @(
  'AP','Radio','DurationSec',
  'RxRetry_per_s','RxCRC_per_s','RxPLCP_per_s',
  'ChannelChanges_per_h','TxPowerChanges_per_h',
  'Busy1s_pct','Busy4s_pct','Busy64s_pct',
  'BusyBeacon_pct','TxBeacon_pct','RxBeacon_pct',
  'CCA_Our_pct','CCA_Other_pct','CCA_Interference_pct',
  'Diagnosis'
) -join ','
Set-Content -LiteralPath $OutputCsv -Value $header -Encoding UTF8

# 出力行（HTML用も保持）
$rows = @()

# キーの和集合
$keys = New-Object System.Collections.Generic.HashSet[string]
foreach ($k in $before.Keys) { [void]$keys.Add($k) }
foreach ($k in $after.Keys)  { [void]$keys.Add($k) }

foreach ($k in $keys) {
  $b = $null; $a = $null
  if ($before.ContainsKey($k)) { $b = $before[$k] }
  if ($after.ContainsKey($k))  { $a = $after[$k] }

  # AP/Radio
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

  # --- 空行抑止：何も値が無い行はスキップ（DurationSec以外で） ---
  $hasAny = $false
  $valsToCheck = @($retry_ps,$crc_ps,$plcp_ps,$chg_ph,$txp_ph,$busy1s,$busy4s,$busy64,$busyB,$txB,$rxB,$ccaO,$ccaOt,$ccaI)
  foreach ($vv in $valsToCheck) { if ($vv -ne $null -and $vv -ne '') { $hasAny = $true } }
  if (-not $hasAny) { continue }

  # 診断生成
  $diag = Build-Diagnosis -Busy64 $busy64 -BusyB $busyB -TxB $txB -RxB $rxB `
                          -CCAO $ccaOt -CCAI $ccaI -Retry $retry_ps -CRC $crc_ps -PLCP $plcp_ps `
                          -ChgPH $chg_ph -TxPPH $txp_ph
  $tags = $diag[0]; $sev = $diag[1]
  $diagText = ($tags -join '; ')

  # CSV出力
  $vals = @(
    $ap, $radio, $DurationSec,
    $retry_ps, $crc_ps, $plcp_ps,
    $chg_ph, $txp_ph,
    $busy1s, $busy4s, $busy64,
    $busyB, $txB, $rxB,
    $ccaO, $ccaOt, $ccaI,
    $diagText
  ) | ForEach-Object { if ($_ -eq $null) { '' } else { $_.ToString() } }

  $escaped = @()
  foreach ($v in $vals) {
    if ($v -match '[,"]') { $escaped += ('"{0}"' -f ($v -replace '"','""')) } else { $escaped += $v }
  }
  Add-Content -LiteralPath $OutputCsv -Value ($escaped -join ',') -Encoding UTF8

  # HTML用格納
  $row = New-Object psobject -Property @{
    AP=$ap; Radio=$radio; DurationSec=$DurationSec;
    RxRetry_per_s=$retry_ps; RxCRC_per_s=$crc_ps; RxPLCP_per_s=$plcp_ps;
    ChannelChanges_per_h=$chg_ph; TxPowerChanges_per_h=$txp_ph;
    Busy1s_pct=$busy1s; Busy4s_pct=$busy4s; Busy64s_pct=$busy64;
    BusyBeacon_pct=$busyB; TxBeacon_pct=$txB; RxBeacon_pct=$rxB;
    CCA_Our_pct=$ccaO; CCA_Other_pct=$ccaOt; CCA_Interference_pct=$ccaI;
    Diagnosis=$tags; Severity=$sev
  }
  $rows += $row
}

Write-Output ("CSV : {0}" -f $OutputCsv)

# ---- HTML出力 ----
if (-not [string]::IsNullOrWhiteSpace($OutputHtml)) {
  $outDir = Get-ParentOrCwd -PathLike $OutputHtml
  if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

  $titleText = $Title
  if ([string]::IsNullOrWhiteSpace($titleText)) {
    $bt = ''; $at = ''
    if ($beforeTime -ne $null) { $bt = $beforeTime.ToString('yyyy-MM-dd HH:mm:ss') }
    if ($afterTime  -ne $null) { $at = $afterTime.ToString('yyyy-MM-dd HH:mm:ss') }
    $titleText = "Aruba Radio Stats Diff ($bt → $at)"
  }

  $cols = @(
    'AP','Radio','DurationSec',
    'RxRetry_per_s','RxCRC_per_s','RxPLCP_per_s',
    'ChannelChanges_per_h','TxPowerChanges_per_h',
    'Busy1s_pct','Busy4s_pct','Busy64s_pct',
    'BusyBeacon_pct','TxBeacon_pct','RxBeacon_pct',
    'CCA_Our_pct','CCA_Other_pct','CCA_Interference_pct',
    'Diagnosis'
  )

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine('<!DOCTYPE html>')
  [void]$sb.AppendLine('<meta charset="UTF-8">')
  [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
  [void]$sb.AppendLine("<title>{0}</title>" -f (HtmlEscape $titleText))
  [void]$sb.AppendLine('<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,"Noto Sans","Hiragino Kaku Gothic ProN","Yu Gothic",sans-serif;margin:16px}
h1{font-size:20px;margin:0 0 8px}
.small{color:#555;font-size:12px;margin-bottom:12px}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
th{background:#f7f7f7;position:sticky;top:0;cursor:pointer}
tr:nth-child(even){background:#fafafa}
input[type="search"]{padding:6px 8px;width:280px;max-width:60%}
.tag{display:inline-block;border:1px solid #ddd;border-radius:3px;padding:2px 6px;margin:2px;background:#fafafa}
.sev-ok  { }
.sev-warn{ background:#fff7e6 }
.sev-crit{ background:#ffecec }
.kpi{display:inline-block;margin-right:16px;font-size:12px}
</style>')

  [void]$sb.AppendLine("<h1>{0}</h1>" -f (HtmlEscape $titleText))
  [void]$sb.AppendLine('<div class="small"><span class="kpi">Busy = 空中占有（%）</span><span class="kpi">Retry/CRC/PLCP = 受信品質（/s）</span><span class="kpi">Changes = ARM変動（/h）</span></div>')
  [void]$sb.AppendLine('<div style="margin:10px 0"><input id="flt" type="search" placeholder="フィルタ（AP/数値/診断を部分一致）..." oninput="filterTable()"></div>')
  [void]$sb.AppendLine('<table id="tbl"><thead><tr>')
  foreach ($c in $cols) { [void]$sb.AppendLine("<th>$c</th>") }
  [void]$sb.AppendLine('</tr></thead><tbody>')

  foreach ($r in $rows) {
    $sev = $r.Severity
    [void]$sb.AppendLine('<tr class="'+ $sev +'">')
    foreach ($c in $cols) {
      if ($c -eq 'Diagnosis') {
        $tags = $r.Diagnosis
        $html = ''
        if ($tags -ne $null) {
          foreach ($t in $tags) { $html += '<span class="tag">'+(HtmlEscape $t)+'</span>' }
        }
        [void]$sb.AppendLine('<td>'+ $html +'</td>')
        continue
      }
      $v = $r.PSObject.Properties[$c].Value
      $text = ''
      if ($null -ne $v) {
        if ($v -is [double] -or $v -is [single]) { $text = ([string]([Math]::Round([double]$v,6))) }
        else { $text = [string]$v }
      }
      [void]$sb.AppendLine('<td>'+ (HtmlEscape $text) +'</td>')
    }
    [void]$sb.AppendLine('</tr>')
  }

  [void]$sb.AppendLine('</tbody></table>')
  [void]$sb.AppendLine('<script>
(function(){
  var tbl=document.getElementById("tbl");
  var ths=tbl.tHead.rows[0].cells;
  var lastCol=-1, asc=true;
  for(var i=0;i<ths.length;i++){
    (function(idx){
      ths[idx].addEventListener("click", function(){
        if(lastCol===idx){ asc=!asc; } else { lastCol=idx; asc=true; }
        sortBy(idx,asc);
      });
    })(i);
  }
  function getVal(td){
    var t=td.textContent; var n=parseFloat(t);
    if(!isNaN(n)) return {n:n, s:t.toLowerCase()};
    return {n:null, s:t.toLowerCase()};
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
      var show=false, tds=trs[i].cells;
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

  # 出力パス
  $htmlDir  = Get-ParentOrCwd -PathLike $OutputHtml
  $nameOnly = $null
  try { $nameOnly = Split-Path -Path $OutputHtml -Leaf } catch { $nameOnly = $null }
  if ([string]::IsNullOrWhiteSpace($nameOnly)) { $nameOnly = "radio_stats_diff.html" }
  $htmlPath = Join-Path -Path $htmlDir -ChildPath $nameOnly

  Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8
  Write-Output ("HTML: {0}" -f $htmlPath)
}