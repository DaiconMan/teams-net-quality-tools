<# 
.SYNOPSIS
  Aruba "show ap debug radio-stats" の2スナップショットから差分/秒を計算しCSV/HTML出力。
  HTML/CSVに「わかりやすい判定（SimpleDiag）」と「おすすめ対策（Tips）」を出力。

.DESCRIPTION
  - PowerShell 5.1 対応。三項演算子未使用。予約語 Host 不使用。
  - OneDrive/日本語/スペース対応（安全な Split-Path/Join-Path）。Cドライブ固定参照なし。
  - Before/After のテキスト出力から主要指標を抽出し、差分/秒や@beaconの瞬時値を算出。
  - 経過秒は "output time" を優先し自動算出（失敗時はファイル更新時刻差→既定900秒）。
  - 解析時の誤検知対策：コマンド行の誤認を防ぎ、メトリクス行のみオブジェクト生成。全空行は出力しない。

.PARAMETER BeforeFile
  先に取得した radio-stats のテキストファイル。
.PARAMETER AfterFile
  後で取得した radio-stats のテキストファイル。
.PARAMETER DurationSec
  経過秒を明示的に上書きしたい場合のみ指定（通常は省略）。
.PARAMETER OutputCsv
  出力CSVのパス。未指定時は AfterFile の親（不可ならカレント）に自動命名。
.PARAMETER OutputHtml
  出力HTMLのパス。指定時は CSV に加えて HTML も生成。
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

# ===== しきい値（必要なら調整） =====
$TH_BusyHigh   = 60    # Busy64s 高
$TH_BusyWarn   = 40    # Busy64s 注意
$TH_ExcessCCA  = 15    # BusyBeacon - (Tx+Rx) 閾値
$TH_OtherHigh  = 20    # CCA_Other 閾値（同チャネル）
$TH_InterfHigh = 10    # CCA_Interference 閾値（非Wi-Fi）
$TH_RetryHigh  = 5     # RxRetry_per_s
$TH_CRCHigh    = 100   # RxCRC_per_s
$TH_PLCPHigh   = 50    # RxPLCP_per_s
$TH_ChgHigh    = 3     # ChannelChanges_per_h
$TH_TxPHigh    = 10    # TxPowerChanges_per_h

# ===== パスユーティリティ =====
function Get-ParentOrCwd {
  param([string]$PathLike)
  $dir = $null
  if (-not [string]::IsNullOrWhiteSpace($PathLike)) {
    if (Test-Path -LiteralPath $PathLike) { try { $dir = Split-Path -LiteralPath $PathLike -Parent } catch { $dir = $null } }
    else { try { $dir = Split-Path -Path $PathLike -Parent } catch { $dir = $null } }
  }
  if ([string]::IsNullOrWhiteSpace($dir)) { try { return (Get-Location).Path } catch { return "." } }
  return $dir
}

# ===== テキスト解析ユーティリティ =====
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
  $cand=@(); foreach($raw in $Lines){ if($raw -match '(?i)(output\s*time|出力(時刻|時間|日時)|生成時刻)'){ $cand+=$raw } }
  if ($cand.Count -eq 0) { return $null }
  foreach($line in $cand){
    $m=[regex]::Match($line,'(?<!\d)(\d{10})(?:\.\d+)?(?!\d)'); if($m.Success){ try{ return ([DateTime]'1970-01-01').AddSeconds([double]$m.Groups[1].Value) }catch{} }
  }
  foreach($line in $cand){
    $m=[regex]::Match($line,'(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})'); if($m.Success){ try{ return [DateTime]::Parse($m.Groups[1].Value,[System.Globalization.CultureInfo]::InvariantCulture) }catch{} }
  }
  foreach($line in $cand){
    $m=[regex]::Match($line,'((\d{1,4})/(\d{1,2})/(\d{1,4})\s+\d{1,2}:\d{2}:\d{2})')
    if($m.Success){
      try{ return [DateTime]::Parse($m.Groups[1].Value,[System.Globalization.CultureInfo]::InvariantCulture) }catch{}
      try{ return [DateTime]::Parse($m.Groups[1].Value,[System.Globalization.CultureInfo]::GetCultureInfo('ja-JP')) }catch{}
      try{ return [DateTime]::Parse($m.Groups[1].Value,[System.Globalization.CultureInfo]::GetCultureInfo('en-US')) }catch{}
    }
  }
  foreach($line in $cand){
    $m=[regex]::Match($line,'([A-Za-z]{3}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}(?:\s+\d{4})?)')
    if($m.Success){
      $s=$m.Groups[1].Value; if($s -notmatch '\s\d{4}$'){ $s = "$s " + (Get-Date).Year }
      try{ return [DateTime]::Parse($s,[System.Globalization.CultureInfo]::GetCultureInfo('en-US')) }catch{}
    }
  }
  return $null
}

# ===== シンプル診断（主原因＋理由＋対策） =====
function Make-SimpleDiagnosis {
  param(
    [double]$Busy64,[double]$BusyB,[double]$TxB,[double]$RxB,
    [double]$CCAO,[double]$CCAI,[double]$Retry,[double]$CRC,[double]$PLCP,
    [double]$ChgPH,[double]$TxPPH
  )

  $deltaCCA = $null
  if ($BusyB -ne $null -and $TxB -ne $null -and $RxB -ne $null) { $deltaCCA = $BusyB - ($TxB + $RxB) }

  # スコア（ヒット数＋強さの合算で簡易評価）
  $scoreCoch = 0; $strengthCoch = 0
  if ($deltaCCA -ne $null -and $deltaCCA -ge $TH_ExcessCCA) { $scoreCoch += 1; $strengthCoch += $deltaCCA }
  if ($CCAO -ne $null -and $CCAO -ge $TH_OtherHigh)        { $scoreCoch += 1; $strengthCoch += $CCAO }
  if ($Busy64 -ne $null -and $Busy64 -ge $TH_BusyWarn)      { $scoreCoch += 1; $strengthCoch += ($Busy64 - $TH_BusyWarn) }

  $scoreInter = 0; $strengthInter = 0
  if ($CCAI -ne $null -and $CCAI -ge $TH_InterfHigh)        { $scoreInter += 1; $strengthInter += $CCAI }
  if ($PLCP -ne $null -and $PLCP -ge $TH_PLCPHigh)          { $scoreInter += 1; $strengthInter += [Math]::Min($PLCP, $TH_PLCPHigh*4)/10.0 }

  $scoreQual = 0; $strengthQual = 0
  if ($Busy64 -ne $null -and $Busy64 -le $TH_BusyWarn) {
    if ($Retry -ne $null -and $Retry -ge $TH_RetryHigh) { $scoreQual += 1; $strengthQual += $Retry }
    if ($CRC   -ne $null -and $CRC   -ge $TH_CRCHigh)   { $scoreQual += 1; $strengthQual += ($CRC/10.0) }
  }

  $scoreBusy = 0; $strengthBusy = 0
  if ($Busy64 -ne $null -and $Busy64 -ge $TH_BusyHigh) { $scoreBusy = 1; $strengthBusy = $Busy64 }

  # 主原因の決定（スコア優先、同点は強さで決定）
  $labels = @('co','inter','qual','busy')
  $scores = @($scoreCoch,$scoreInter,$scoreQual,$scoreBusy)
  $powers = @($strengthCoch,$strengthInter,$strengthQual,$strengthBusy)

  $maxIdx = 0; $i = 0
  while ($i -lt $scores.Length) { if ($scores[$i] -gt $scores[$maxIdx]) { $maxIdx = $i } elseif ($scores[$i] -eq $scores[$maxIdx]) { if ($powers[$i] -gt $powers[$maxIdx]) { $maxIdx = $i } } $i++ }

  $root = $labels[$maxIdx]

  # 一言診断＆根拠＆対策
  $simple = ''; $why = ''; $tips = ''

  if ($root -eq 'co') {
    $simple = "原因：近くの"同じチャンネル"のWi-Fiが強く、電波の取り合いが発生しています。"
    $whyParts = @()
    if ($CCAO -ne $null)     { $whyParts += ("他BSS " + [int]$CCAO + "%") }
    if ($deltaCCA -ne $null) { $whyParts += ("占有差ΔCCA " + [int]$deltaCCA + "pt") }
    if ($Busy64 -ne $null)   { $whyParts += ("Busy64 " + [int]$Busy64 + "%") }
    $why = "根拠：" + ([string]::Join(" / ", $whyParts))
    $tips = "対策：チャネル再配置・再利用距離を拡大／帯域幅を20MHzに縮小／最低基本レートの引き上げ。"
  }
  elseif ($root -eq 'inter') {
    $simple = "原因：Wi-Fi以外の電波ノイズの影響が大きい可能性があります。"
    $whyParts = @()
    if ($CCAI -ne $null)   { $whyParts += ("Interference " + [int]$CCAI + "%") }
    if ($PLCP -ne $null)   { $whyParts += ("PLCP " + [int]$PLCP + "/s") }
    $why = "根拠：" + ([string]::Join(" / ", $whyParts))
    $tips = "対策：別チャネル（非DFS含む）に一時固定して観測／周辺の装置（AV機器・無線機器等）の稼働時間と突き合わせ。"
  }
  elseif ($root -eq 'qual') {
    $simple = "原因：端末の電波が弱い／遮蔽／隠れ端末の可能性が高いです。"
    $whyParts = @("Busyが低め")
    if ($Retry -ne $null) { $whyParts += ("Retry " + [int]$Retry + "/s") }
    if ($CRC   -ne $null) { $whyParts += ("CRC " + [int]$CRC + "/s") }
    $why = "根拠：" + ([string]::Join(" / ", $whyParts))
    $tips = "対策：AP配置・出力レンジの適正化／ローミング閾値の見直し／低速端末の抑制。"
  }
  else {
    $simple = "状況：電波の混雑が高く、全体的にエアタイムが逼迫しています。"
    $why = "根拠：Busy64 " + [int]$Busy64 + "%"
    $tips = "対策：ユーザー密度の分散／帯域幅20MHz化／不要な低速レートの無効化。"
  }

  $oneLine = $simple + " " + $why
  return ,@($oneLine, $tips)
}

function Diff-NonNegative { param([double]$After,[double]$Before)
  if ($After -eq $null -or $Before -eq $null) { return $null }
  $d = $After - $Before
  if ($d -lt 0) { return 0.0 }
  return $d
}

# ===== パーサ（誤認抑止版） =====
function Parse-RadioStatsFile {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
  $lines = Get-Content -LiteralPath $Path -Encoding UTF8
  $outTime = Extract-OutputTime -Lines $lines

  $result = @{}
  $ap = ''; $radio = ''

  foreach ($raw in $lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    $mApName = [regex]::Match($line, '(?i)\bap[-\s_]*name\s+([A-Za-z0-9_\-\.:]+)\b')
    if ($mApName.Success) { $ap = $mApName.Groups[1].Value }

    $mHdr = [regex]::Match($line, '^(?i)\s*AP\s+([^\s]+).*?\bRadio\s+([01])\b')
    if ($mHdr.Success) { $ap = $mHdr.Groups[1].Value; $radio = $mHdr.Groups[2].Value }

    $mRadio = [regex]::Match($line, '(?i)\bradio\s+([01])\b')
    if ($mRadio.Success) { $radio = $mRadio.Groups[1].Value }

    # メトリクス行のみ採用
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

    $key = ''
    if (-not [string]::IsNullOrWhiteSpace($ap)) { if (-not [string]::IsNullOrWhiteSpace($radio)) { $key = "$ap|$radio" } else { $key = "$ap|?" } }
    else { if (-not [string]::IsNullOrWhiteSpace($radio)) { $key = "Unknown|$radio" } else { $key = "Unknown|?" } }

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
  return (New-Object psobject -Property @{ Data = $result; OutputTime = $outTime })
}

# ===== メイン =====
$beforeObj = Parse-RadioStatsFile -Path $BeforeFile
$afterObj  = Parse-RadioStatsFile -Path $AfterFile

$before = $beforeObj.Data; $after = $afterObj.Data
$beforeTime = $beforeObj.OutputTime; $afterTime = $afterObj.OutputTime

if ($DurationSec -le 0) {
  $sec = 0
  if ($beforeTime -ne $null -and $afterTime -ne $null) { try { $sec = [int][Math]::Abs(($afterTime - $beforeTime).TotalSeconds) } catch { $sec = 0 } }
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

# CSVパス
if ([string]::IsNullOrWhiteSpace($OutputCsv)) {
  $outDir = Get-ParentOrCwd -PathLike $AfterFile
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $OutputCsv = Join-Path -Path $outDir -ChildPath ("aruba_radio_stats_diff_{0}.csv" -f $ts)
}

# CSVヘッダ（SimpleDiag/Tipsを追加）
$header = @(
  'AP','Radio','DurationSec',
  'RxRetry_per_s','RxCRC_per_s','RxPLCP_per_s',
  'ChannelChanges_per_h','TxPowerChanges_per_h',
  'Busy1s_pct','Busy4s_pct','Busy64s_pct',
  'BusyBeacon_pct','TxBeacon_pct','RxBeacon_pct',
  'CCA_Our_pct','CCA_Other_pct','CCA_Interference_pct',
  'SimpleDiag','Tips'
) -join ','
Set-Content -LiteralPath $OutputCsv -Value $header -Encoding UTF8

# 出力用
$rows=@()
$keys = New-Object System.Collections.Generic.HashSet[string]
foreach ($k in $before.Keys) { [void]$keys.Add($k) }
foreach ($k in $after.Keys)  { [void]$keys.Add($k) }

foreach ($k in $keys) {
  $b=$null;$a=$null
  if ($before.ContainsKey($k)) { $b=$before[$k] }
  if ($after.ContainsKey($k))  { $a=$after[$k] }

  $ap='';$radio=''
  if ($a -ne $null) { $ap=$a.AP; $radio=$a.Radio }
  if ([string]::IsNullOrWhiteSpace($ap) -and $b -ne $null) { $ap=$b.AP }
  if ([string]::IsNullOrWhiteSpace($radio) -and $b -ne $null) { $radio=$b.Radio }

  # 差分
  $dRetry = Diff-NonNegative $a.RxRetry $b.RxRetry
  $dCRC   = Diff-NonNegative $a.RxCRC   $b.RxCRC
  $dPLCP  = Diff-NonNegative $a.RxPLCP  $b.RxPLCP
  $dChg   = Diff-NonNegative $a.ChannelChanges $b.ChannelChanges
  $dTxPw  = Diff-NonNegative $a.TxPowerChanges $b.TxPowerChanges

  $retry_ps=$null;$crc_ps=$null;$plcp_ps=$null;$chg_ph=$null;$txp_ph=$null
  if ($dRetry -ne $null) { $retry_ps=[Math]::Round($dRetry/$DurationSec,6) }
  if ($dCRC   -ne $null) { $crc_ps  =[Math]::Round($dCRC  /$DurationSec,6) }
  if ($dPLCP  -ne $null) { $plcp_ps =[Math]::Round($dPLCP /$DurationSec,6) }
  if ($dChg   -ne $null) { $chg_ph  =[Math]::Round(($dChg*3600.0)/$DurationSec,6) }
  if ($dTxPw  -ne $null) { $txp_ph  =[Math]::Round(($dTxPw*3600.0)/$DurationSec,6) }

  function Pick-AfterFirst { param($afterV,$beforeV) if ($afterV -ne $null) { return $afterV } return $beforeV }
  $busy1s = Pick-AfterFirst $a.Busy1s $b.Busy1s
  $busy4s = Pick-AfterFirst $a.Busy4s $b.Busy4s
  $busy64 = Pick-AfterFirst $a.Busy64s $b.Busy64s
  $busyB  = Pick-AfterFirst $a.BusyBeacon $b.BusyBeacon
  $txB    = Pick-AfterFirst $a.TxBeacon   $b.TxBeacon
  $rxB    = Pick-AfterFirst $a.RxBeacon   $b.RxBeacon
  $ccaO   = Pick-AfterFirst $a.CCA_Our $b.CCA_Our
  $ccaOt  = Pick-AfterFirst $a.CCA_Other $b.CCA_Other
  $ccaI   = Pick-AfterFirst $a.CCA_Interference $b.CCA_Interference

  # 何も値が無ければスキップ（空行抑止）
  $hasAny=$false
  foreach($vv in @($retry_ps,$crc_ps,$plcp_ps,$chg_ph,$txp_ph,$busy1s,$busy4s,$busy64,$busyB,$txB,$rxB,$ccaO,$ccaOt,$ccaI)){
    if($vv -ne $null -and $vv -ne ''){ $hasAny=$true }
  }
  if(-not $hasAny){ continue }

  # 一言診断＋対策
  $sd = Make-SimpleDiagnosis -Busy64 $busy64 -BusyB $busyB -TxB $txB -RxB $rxB `
                             -CCAO $ccaOt -CCAI $ccaI -Retry $retry_ps -CRC $crc_ps -PLCP $plcp_ps `
                             -ChgPH $chg_ph -TxPPH $txp_ph
  $simple = $sd[0]; $tips = $sd[1]

  # CSV
  $vals = @(
    $ap,$radio,$DurationSec,
    $retry_ps,$crc_ps,$plcp_ps,
    $chg_ph,$txp_ph,
    $busy1s,$busy4s,$busy64,
    $busyB,$txB,$rxB,
    $ccaO,$ccaOt,$ccaI,
    $simple,$tips
  ) | ForEach-Object { if ($_ -eq $null) { '' } else { $_.ToString() } }

  $escaped=@(); foreach($v in $vals){ if($v -match '[,"]'){ $escaped+=('"{0}"' -f ($v -replace '"','""')) } else { $escaped+=$v } }
  Add-Content -LiteralPath $OutputCsv -Value ($escaped -join ',') -Encoding UTF8

  # HTML用
  $rows += (New-Object psobject -Property @{
    AP=$ap; Radio=$radio; DurationSec=$DurationSec;
    RxRetry_per_s=$retry_ps; RxCRC_per_s=$crc_ps; RxPLCP_per_s=$plcp_ps;
    ChannelChanges_per_h=$chg_ph; TxPowerChanges_per_h=$txp_ph;
    Busy1s_pct=$busy1s; Busy4s_pct=$busy4s; Busy64s_pct=$busy64;
    BusyBeacon_pct=$busyB; TxBeacon_pct=$txB; RxBeacon_pct=$rxB;
    CCA_Our_pct=$ccaO; CCA_Other_pct=$ccaOt; CCA_Interference_pct=$ccaI;
    SimpleDiag=$simple; Tips=$tips
  })
}

Write-Output ("CSV : {0}" -f $OutputCsv)

# ===== HTML =====
if (-not [string]::IsNullOrWhiteSpace($OutputHtml)) {
  $outDir = Get-ParentOrCwd -PathLike $OutputHtml
  if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

  $titleText = $Title
  if ([string]::IsNullOrWhiteSpace($titleText)) {
    $bt=''; $at=''
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
    'SimpleDiag','Tips'
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
.diag{font-weight:600}
.tip{font-size:12px;color:#333}
</style>')
  [void]$sb.AppendLine("<h1>{0}</h1>" -f (HtmlEscape $titleText))
  [void]$sb.AppendLine('<div class="small">ヘッダークリックでソート／検索でフィルタ</div>')
  [void]$sb.AppendLine('<div style="margin:10px 0"><input id="flt" type="search" placeholder="フィルタ（AP/数値/文言）..." oninput="filterTable()"></div>')
  [void]$sb.AppendLine('<table id="tbl"><thead><tr>')
  foreach ($c in $cols) { [void]$sb.AppendLine("<th>$c</th>") }
  [void]$sb.AppendLine('</tr></thead><tbody>')

  foreach ($r in $rows) {
    [void]$sb.AppendLine('<tr>')
    foreach ($c in $cols) {
      $v = $r.PSObject.Properties[$c].Value
      if ($c -eq 'SimpleDiag') {
        [void]$sb.AppendLine('<td class="diag">'+ (HtmlEscape $v) +'</td>'); continue
      }
      if ($c -eq 'Tips') {
        [void]$sb.AppendLine('<td class="tip">'+ (HtmlEscape $v) +'</td>'); continue
      }
      $text=''; if ($null -ne $v) { if ($v -is [double] -or $v -is [single]) { $text=([string]([Math]::Round([double]$v,6))) } else { $text=[string]$v } }
      [void]$sb.AppendLine('<td>'+ (HtmlEscape $text) +'</td>')
    }
    [void]$sb.AppendLine('</tr>')
  }

  [void]$sb.AppendLine('</tbody></table>')
  [void]$sb.AppendLine('<script>
(function(){
  var tbl=document.getElementById("tbl");
  var ths=tbl.tHead.rows[0].cells; var lastCol=-1, asc=true;
  for(var i=0;i<ths.length;i++){
    (function(idx){
      ths[idx].addEventListener("click", function(){
        if(lastCol===idx){ asc=!asc; } else { lastCol=idx; asc=true; }
        sortBy(idx,asc);
      });
    })(i);
  }
  function getVal(td){ var t=td.textContent; var n=parseFloat(t); if(!isNaN(n)) return {n:n,s:t.toLowerCase()}; return {n:null,s:t.toLowerCase()}; }
  function cmp(a,b,ascFlag){
    if(a.n!==null&&b.n!==null){ if(a.n<b.n) return ascFlag?-1:1; if(a.n>b.n) return ascFlag?1:-1; return 0; }
    if(a.s<b.s) return ascFlag?-1:1; if(a.s>b.s) return ascFlag?1:-1; return 0;
  }
  function sortBy(col,ascFlag){
    var tbody=tbl.tBodies[0]; var rows=[].slice.call(tbody.rows);
    rows.sort(function(r1,r2){ var a=getVal(r1.cells[col]); var b=getVal(r2.cells[col]); return cmp(a,b,ascFlag); });
    for(var i=0;i<rows.length;i++){ tbody.appendChild(rows[i]); }
  }
  window.filterTable=function(){
    var q=document.getElementById("flt").value.toLowerCase(); var trs=tbl.tBodies[0].rows;
    for(var i=0;i<trs.length;i++){
      var show=false, tds=trs[i].cells;
      for(var j=0;j<tds.length;j++){ var t=tds[j].textContent.toLowerCase(); if(t.indexOf(q)>=0){ show=true; break; } }
      trs[i].style.display = show? "":"none";
    }
  };
})();
</script>')
  $html = $sb.ToString()

  $htmlDir  = Get-ParentOrCwd -PathLike $OutputHtml
  $nameOnly = $null; try { $nameOnly = Split-Path -Path $OutputHtml -Leaf } catch { $nameOnly = $null }
  if ([string]::IsNullOrWhiteSpace($nameOnly)) { $nameOnly = "radio_stats_diff.html" }
  $htmlPath = Join-Path -Path $htmlDir -ChildPath $nameOnly

  Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8
  Write-Output ("HTML: {0}" -f $htmlPath)
}
exit 0