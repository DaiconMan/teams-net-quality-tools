<# 
.SYNOPSIS
  Aruba "show ap debug radio-stats" スナップショットの差分/秒を算出し、CSV/HTML を生成（JST表示）。
  - 上部カードに簡易診断（Simple/Tips/副次）を表示
  - ただし CSV/HTML の「行末列」としての SimpleDiag/Tips は出力しない（要望対応）
  - Output Time が「... UTC」でも厳密にUTCとして取り込み→JST変換
  - Output Time が無い場合、ファイル更新時刻をJSTで表示にフォールバック
  - -SnapshotFiles はフォルダ/ワイルドカード/ファイル混在OK（展開後2ファイル以上）
  - Channel/Band 抽出、LAA/NR-U候補表示、時間帯コメント

.NOTES
  - PowerShell 5.1対応。三項演算子(?)不使用。
  - Join-Path は -Path 使用。予約語 Host は使用しない。
  - OneDrive/日本語/スペースパス配慮。Cドライブ固定参照なし。
#>

[CmdletBinding()]
param(
  # 単区間比較
  [string]$BeforeFile,
  [string]$AfterFile,
  [int]$DurationSec,

  # 時系列集計（展開後に2ファイル以上が必要）: ファイル/ワイルドカード/ディレクトリ混在可
  [string[]]$SnapshotFiles,

  # 出力
  [string]$OutputCsv,
  [string]$OutputHtml,
  [string]$Title
)

# ===== しきい値 =====
$TH_BusyHigh   = 60
$TH_BusyWarn   = 40
$TH_ExcessCCA  = 15
$TH_OtherHigh  = 20
$TH_InterfHigh = 10
$TH_RetryHigh  = 5
$TH_CRCHigh    = 100
$TH_PLCPHigh   = 50
$TH_ChgHigh    = 3
$TH_TxPHigh    = 10

# LAA/NR-U 疑い（保守的）
$TH_LAA_MinInterf = 12
$TH_LAA_MaxOther  = 10
$TH_LAA_MinBusy   = 35

# ===== JST 変換（UTC/Local/Unspecified → JST） =====
function Convert-ToJst {
  param([Nullable[DateTime]]$dt)
  if ($dt -eq $null) { return $null }
  try { $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById("Tokyo Standard Time") } catch { $tz = [System.TimeZoneInfo]::Local }
  try {
    if ($dt.Value.Kind -eq [System.DateTimeKind]::Utc) {
      return [System.TimeZoneInfo]::ConvertTimeFromUtc($dt.Value, $tz)
    } elseif ($dt.Value.Kind -eq [System.DateTimeKind]::Local) {
      return [System.TimeZoneInfo]::ConvertTime($dt.Value, $tz)
    } else {
      $utc = [DateTime]::SpecifyKind($dt.Value, [System.DateTimeKind]::Utc)
      return [System.TimeZoneInfo]::ConvertTimeFromUtc($utc, $tz)
    }
  } catch { return $dt }
}

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

# ===== 表示ユーティリティ =====
function HtmlEscape { param([string]$s)
  if ($null -eq $s) { return '' }
  $r = $s.Replace('&','&amp;'); $r = $r.Replace('<','&lt;'); $r = $r.Replace('>','&gt;')
  $r = $r.Replace('"','&quot;'); $r = $r.Replace("'",'&#39;'); return $r
}
function Sanitize-Text { param([string]$s) if ($null -eq $s) { return '' } return $s.Replace('"','"') }

# ===== 数値抽出 =====
function Get-LastNumber { param([string]$Line)
  if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
  $m = [regex]::Matches($Line, '(-?\d+(?:\.\d+)?)')
  if ($m.Count -gt 0) { return [double]$m[$m.Count-1].Value }
  return $null
}
function TryExtractPercentTriplet {
  param([string]$Line, [ref]$Busy1s, [ref]$Busy4s, [ref]$Busy64s)
  $ok = $false
  $m1 = [regex]::Match($Line, '\b1s\b[^0-9\-]*(-?\d+(?:\.\d+)?)'); if ($m1.Success) { $Busy1s.Value = [double]$m1.Groups[1].Value; $ok = $true }
  $m4 = [regex]::Match($Line, '\b4s\b[^0-9\-]*(-?\d+(?:\.\d+)?)'); if ($m4.Success) { $Busy4s.Value = [double]$m4.Groups[1].Value; $ok = $true }
  $m64= [regex]::Match($Line, '\b64s\b[^0-9\-]*(-?\d+(?:\.\d+)?)'); if ($m64.Success){ $Busy64s.Value= [double]$m64.Groups[1].Value; $ok = $true }
  return $ok
}

# ===== 文字列→UTC（DateTime） =====
function Parse-AnyToUtc {
  param([string]$s, [string]$cultureName)
  $ci = $null
  try { if (-not [string]::IsNullOrWhiteSpace($cultureName)) { $ci = [System.Globalization.CultureInfo]::GetCultureInfo($cultureName) } } catch { $ci = $null }
  if ($ci -eq $null) { $ci = [System.Globalization.CultureInfo]::InvariantCulture }

  $dto = [System.DateTimeOffset]::MinValue
  $ok = [System.DateTimeOffset]::TryParse($s, $ci, [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)
  if ($ok) { return $dto.UtcDateTime }

  $dt = [DateTime]::MinValue
  $ok2 = [DateTime]::TryParse($s, $ci, [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dt)
  if ($ok2) {
    if ($dt.Kind -ne [System.DateTimeKind]::Utc) { $dt = [DateTime]::SpecifyKind($dt, [System.DateTimeKind]::Utc) }
    return $dt
  }

  try {
    $dt2 = [DateTime]::Parse($s, $ci)
    $dt2 = [DateTime]::SpecifyKind($dt2, [System.DateTimeKind]::Utc)
    return $dt2
  } catch { return $null }
}

# ===== Output Time 抽出（UTCで返す） =====
function Extract-OutputTime {
  param([string[]]$Lines)
  $cand=@(); foreach($raw in $Lines){ if($raw -match '(?i)(output\s*time|出力(時刻|時間|日時)|生成時刻)'){ $cand+=$raw } }
  if ($cand.Count -eq 0) { return $null }

  # 1) Unix epoch (sec) → UTC
  foreach($line in $cand){
    $m=[regex]::Match($line,'(?<!\d)(\d{10})(?:\.\d+)?(?!\d)')
    if($m.Success){ try{ $sec=[long]([double]$m.Groups[1].Value); return ([System.DateTimeOffset]::FromUnixTimeSeconds($sec)).UtcDateTime }catch{} }
  }

  # 2) ISO 8601 / "YYYY-MM-DD HH:mm:ssZ" / +hh:mm → UTC扱い
  foreach($line in $cand){
    $m=[regex]::Match($line,'(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(?:Z|[+\-]\d{2}:\d{2})?)')
    if($m.Success){ $u = Parse-AnyToUtc -s $m.Groups[1].Value -cultureName ''; if ($u -ne $null) { return $u } }
  }

  # 2.5) "YYYY-MM-DD HH:mm:ss UTC/GMT" → UTCとして解釈（★追加）
  foreach($line in $cand){
    $m=[regex]::Match($line,'(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})\s*(UTC|GMT)\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if($m.Success){
      try{
        $dt = [DateTime]::Parse($m.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture)
        $dt = [DateTime]::SpecifyKind($dt, [System.DateTimeKind]::Utc)
        return $dt
      }catch{}
    }
  }

  # 3) "MM/DD/YYYY HH:mm:ss" 等 → UTCとして扱う
  foreach($line in $cand){
    $m=[regex]::Match($line,'((\d{1,4})/(\d{1,2})/(\d{1,4})\s+\d{1,2}:\d{2}:\d{2})')
    if($m.Success){
      $u = Parse-AnyToUtc -s $m.Groups[1].Value -cultureName 'ja-JP'
      if ($u -eq $null) { $u = Parse-AnyToUtc -s $m.Groups[1].Value -cultureName 'en-US' }
      if ($u -ne $null) { return $u }
    }
  }

  # 4) "Mon 12 23:59:01 [YYYY]" → UTCとして扱う
  foreach($line in $cand){
    $m=[regex]::Match($line,'([A-Za-z]{3}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}(?:\s+\d{4})?)')
    if($m.Success){
      $s=$m.Groups[1].Value
      if($s -notmatch '\s\d{4}$'){ $s = "$s " + (Get-Date).Year }
      $u = Parse-AnyToUtc -s $s -cultureName 'en-US'
      if ($u -ne $null) { return $u }
    }
  }
  return $null
}

# ===== チャネル→バンド推定 =====
function Get-BandFromChannel {
  param([Nullable[int]]$Channel)
  if ($Channel -eq $null) { return '' }
  $ch = [int]$Channel
  if ($ch -ge 1 -and $ch -le 14)  { return '2.4GHz' }
  if ($ch -ge 32 -and $ch -le 196){ return '5GHz' }
  return ''
}

# ===== 診断（カード生成用） =====
function Make-Diagnosis {
  param(
    [double]$Busy64,[double]$BusyB,[double]$TxB,[double]$RxB,
    [double]$CCAO,[double]$CCAI,[double]$Retry,[double]$CRC,[double]$PLCP,
    [double]$ChgPH,[double]$TxPPH,
    [Nullable[int]]$Channel
  )

  $deltaCCA = $null
  if ($BusyB -ne $null -and $TxB -ne $null -and $RxB -ne $null) { $deltaCCA = $BusyB - ($TxB + $RxB) }

  $scoreCoch=0;$pCoch=0.0
  if ($deltaCCA -ne $null -and $deltaCCA -ge $TH_ExcessCCA){$scoreCoch+=1;$pCoch+=$deltaCCA}
  if ($CCAO -ne $null -and $CCAO -ge $TH_OtherHigh){$scoreCoch+=1;$pCoch+=$CCAO}
  if ($Busy64 -ne $null -and $Busy64 -ge $TH_BusyWarn){$scoreCoch+=1;$pCoch+=($Busy64-$TH_BusyWarn)}

  $scoreInter=0;$pInter=0.0
  if ($CCAI -ne $null -and $CCAI -ge $TH_InterfHigh){$scoreInter+=1;$pInter+=$CCAI}
  if ($PLCP -ne $null -and $PLCP -ge $TH_PLCPHigh){$scoreInter+=1;$pInter+=[Math]::Min($PLCP,$TH_PLCPHigh*4)/10.0}

  $scoreQual=0;$pQual=0.0
  if ($Busy64 -ne $null -and $Busy64 -le $TH_BusyWarn){
    if ($Retry -ne $null -and $Retry -ge $TH_RetryHigh){$scoreQual+=1;$pQual+=$Retry}
    if ($CRC -ne $null -and $CRC -ge $TH_CRCHigh){$scoreQual+=1;$pQual+=($CRC/10.0)}
  }

  $scoreBusy=0;$pBusy=0.0
  if ($Busy64 -ne $null -and $Busy64 -ge $TH_BusyHigh){$scoreBusy=1;$pBusy=$Busy64}

  $labels=@('co','inter','qual','busy')
  $scores=@($scoreCoch,$scoreInter,$scoreQual,$scoreBusy)
  $power =@($pCoch,$pInter,$pQual,$pBusy)
  $maxIdx=0;$i=0
  while($i -lt $scores.Length){
    if($scores[$i] -gt $scores[$maxIdx]){$maxIdx=$i}
    elseif($scores[$i] -eq $scores[$maxIdx]){ if($power[$i] -gt $power[$maxIdx]){$maxIdx=$i} }
    $i++
  }
  $root=$labels[$maxIdx]

  $simple='';$why='';$tips='';$sev='sev-ok'
  $secList=New-Object System.Collections.Generic.List[string]

  if($root -eq 'co'){
    $simple="原因：近くの同一チャネルのWi-Fiが強く、電波の取り合いが発生しています。"
    $whyParts=@(); if($CCAO -ne $null){$whyParts+=("他BSS "+[int]$CCAO+"%")}
    if($deltaCCA -ne $null){$whyParts+=("占有差ΔCCA "+[int]$deltaCCA+"pt")}
    if($Busy64 -ne $null){$whyParts+=("Busy64 "+[int]$Busy64+"%")}
    $why="根拠：" + ([string]::Join(" / ",$whyParts))
    $tips="対策：チャネル再配置・再利用距離の拡大／帯域幅20MHz化／最低基本レートの引き上げ。"
    $sev='sev-warn'
  } elseif($root -eq 'inter'){
    $simple="原因：Wi-Fi以外の電波ノイズの影響が大きい可能性があります。"
    $whyParts=@(); if($CCAI -ne $null){$whyParts+=("Interference "+[int]$CCAI+"%")}
    if($PLCP -ne $null){$whyParts+=("PLCP "+[int]$PLCP+"/s")}
    $why="根拠：" + ([string]::Join(" / ",$whyParts))
    $tips="対策：別チャネル（非DFS含む）に一時固定して観測／周辺装置の稼働時間と突き合わせ。"
    $sev='sev-warn'
  } elseif($root -eq 'qual'){
    $simple="原因：端末の電波が弱い・遮蔽・隠れ端末の可能性が高いです。"
    $why="根拠：Busyが低め / Retry "+[int]$Retry+"/s / CRC "+[int]$CRC+"/s"
    $tips="対策：AP配置・出力レンジの適正化／ローミング閾値の見直し／低速端末の抑制。"
    $sev='sev-warn'
  } else {
    $simple="状況：電波の混雑が高く、全体的にエアタイムが逼迫しています。"
    $why="根拠：Busy64 "+[int]$Busy64+"%"
    $tips="対策：ユーザー密度の分散／帯域幅20MHz化／不要な低速レートの無効化。"
    $sev='sev-warn'
  }

  if($root -ne 'inter' -and ($scoreInter -gt 0)){
    $sec="副次：非Wi-Fi干渉の疑い（Interference "+([int]$CCAI)+"%、PLCP "+([int]$PLCP)+"/s）"
    $secList.Add((Sanitize-Text $sec))
  }
  if($root -ne 'co' -and ($scoreCoch -gt 0)){
    $sec="副次：同一チャネル混在（他BSS "+([int]$CCAO)+"%、ΔCCA "+([int]$deltaCCA)+"pt、Busy64 "+([int]$Busy64)+"%）"
    $secList.Add((Sanitize-Text $sec))
  }
  if($root -ne 'qual' -and ($scoreQual -gt 0)){
    $sec="副次：端末側の品質低下（Retry "+([int]$Retry)+"/s、CRC "+([int]$CRC)+"/s）"
    $secList.Add((Sanitize-Text $sec))
  }
  if($root -ne 'busy' -and ($scoreBusy -gt 0)){
    $sec="副次：混雑（Busy64 "+([int]$Busy64)+"%）"
    $secList.Add((Sanitize-Text $sec))
  }

  # LAA/NR-U 候補
  $band = Get-BandFromChannel -Channel $Channel
  $laaHit = $false
  if (-not [string]::IsNullOrWhiteSpace($band)) {
    if ($band -eq '5GHz') {
      if ($CCAI -ne $null -and $CCAI -ge $TH_LAA_MinInterf) {
        $okOther = $false
        if ($CCAO -eq $null) { $okOther = $true } else { if ($CCAO -lt $TH_LAA_MaxOther) { $okOther = $true } }
        if ($okOther) {
          if ($Busy64 -ne $null -and $Busy64 -ge $TH_LAA_MinBusy) { $laaHit = $true }
        }
      }
    }
  }
  if ($laaHit) {
    $msg = "候補：LAA/NR-Uの可能性（5GHz ch"+$Channel+"、Interf "+([int]$CCAI)+"%、Other "+([int]$CCAO)+"%、Busy64 "+([int]$Busy64)+"%）"
    $secList.Add((Sanitize-Text $msg))
  }

  $oneLine=Sanitize-Text ($simple+" "+$why)
  $tips   =Sanitize-Text $tips
  $mainScore = $scores[$maxIdx] + ($power[$maxIdx] / 100.0)

  return New-Object psobject -Property @{
    Simple=$oneLine; Tips=$tips; Secondary=$secList; Severity=$sev; Score=$mainScore; Band=$band
  }
}

function Diff-NonNegative { param([double]$After,[double]$Before)
  if ($After -eq $null -or $Before -eq $null) { return $null }
  $d = $After - $Before
  if ($d -lt 0) { return 0.0 }
  return $d
}

# ===== パーサ =====
function Parse-RadioStatsFile {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
  $lines = Get-Content -LiteralPath $Path -Encoding UTF8
  $outTime = Extract-OutputTime -Lines $lines  # UTC DateTime

  $result = @{}; $ap = ''; $radio = ''
  foreach ($raw in $lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    $mApName = [regex]::Match($line, '(?i)\bap[-\s_]*name\s+([A-Za-z0-9_\-\.:]+)\b')
    if ($mApName.Success) { $ap = $mApName.Groups[1].Value }

    $mHdr = [regex]::Match($line, '^(?i)\s*AP\s+([^\s]+).*?\bRadio\s+([01])\b')
    if ($mHdr.Success) { $ap = $mHdr.Groups[1].Value; $radio = $mHdr.Groups[2].Value }

    $mRadio = [regex]::Match($line, '(?i)\bradio\s+([01])\b')
    if ($mRadio.Success) { $radio = $mRadio.Groups[1].Value }

    $isMetric = $false
    if ($line -match '(?i)\bRx\s*retry\s*frames\b') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bRX?\s*CRC\b.*\bError') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bRX?\s*PLCP\b.*\bError') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bChannel\s*Changes\b') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bTX\s*Power\s*Changes\b') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bChannel\s*busy\b') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\b(Ch|Tx|Rx)\s*Time\s*perct\s*@\s*beacon') { $isMetric = $true }
    if (-not $isMetric -and $line -match '(?i)\bCCA\b.*\b(bss|interference)\b') { $isMetric = $true }

    # チャネル抽出
    $channelNum = $null
    if ($line -match '(?i)\b(Current|Primary|Operating)\s*Channel\s*[:=]\s*(\d{1,3})\b') {
      try { $channelNum = [int]$Matches[2] } catch { $channelNum = $null }
    } elseif ($line -match '(?i)\bChannel\s*[:=]\s*(\d{1,3})\b' -and ($line -notmatch '(?i)Channel\s*Changes')) {
      try { $channelNum = [int]$Matches[1] } catch { $channelNum = $null }
    }
    if (-not $isMetric -and $channelNum -eq $null) { continue }

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
        CCA_Our = $null; CCA_Other = $null; CCA_Interference = $null;
        Channel = $null
      }
      $result[$key] = $obj
    }
    $cur = $result[$key]

    if     ($line -match '(?i)\bRx\s*retry\s*frames\b') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxRetry = [double]$v } }
    elseif ($line -match '(?i)\bRX?\s*CRC\b.*\bError')  { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxCRC  = [double]$v } }
    elseif ($line -match '(?i)\bRX?\s*PLCP\b.*\bError') { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxPLCP = [double]$v } }
    elseif ($line -match '(?i)\bChannel\s*Changes\b')   { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.ChannelChanges = [double]$v } }
    elseif ($line -match '(?i)\bTX\s*Power\s*Changes\b'){ $v = Get-LastNumber $line; if ($v -ne $null) { $cur.TxPowerChanges = [double]$v } }

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
    elseif ($line -match '(?i)\bTx\s*Time\s*perct\s*@\s*beacon'){ $v = Get-LastNumber $line; if ($v -ne $null) { $cur.TxBeacon   = [double]$v } }
    elseif ($line -match '(?i)\bRx\s*Time\s*perct\s*@\s*beacon'){ $v = Get-LastNumber $line; if ($v -ne $null) { $cur.RxBeacon   = [double]$v } }

    if     ($line -match '(?i)\bCCA\b.*\bour\b.*\bbss\b')      { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.CCA_Our          = [double]$v } }
    elseif ($line -match '(?i)\bCCA\b.*\bother\b.*\bbss\b')    { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.CCA_Other        = [double]$v } }
    elseif ($line -match '(?i)\bCCA\b.*\binterference\b')      { $v = Get-LastNumber $line; if ($v -ne $null) { $cur.CCA_Interference = [double]$v } }

    if ($channelNum -ne $null) { $cur.Channel = [int]$channelNum }
  }

  return New-Object psobject -Property @{ Data = $result; OutputTime = $outTime; Path = $Path }
}

# ===== -SnapshotFiles 展開（非再帰） =====
function Expand-SnapshotInputs {
  param([string[]]$Inputs)
  $files = New-Object System.Collections.Generic.List[string]
  if ($Inputs -eq $null) { return $files }

  foreach ($p in $Inputs) {
    if ([string]::IsNullOrWhiteSpace($p)) { continue }

    if (Test-Path -LiteralPath $p) {
      $it = $null
      try { $it = Get-Item -LiteralPath $p -ErrorAction Stop } catch { $it = $null }
      if ($it -ne $null) {
        if ($it.PSIsContainer) {
          $items = @()
          try { $items = Get-ChildItem -LiteralPath $it.FullName -File -ErrorAction Stop } catch { $items = @() }
          foreach ($f in $items) { $files.Add($f.FullName) }
        } else {
          $files.Add($it.FullName)
        }
      }
    } else {
      $parent = $null; $leaf = $null
      try { $parent = Split-Path -Path $p -Parent } catch { $parent = $null }
      try { $leaf   = Split-Path -Path $p -Leaf } catch { $leaf = $p }
      if ([string]::IsNullOrWhiteSpace($parent)) { $parent = "." }
      $patternPath = $null
      try { $patternPath = Join-Path -Path $parent -ChildPath $leaf } catch { $patternPath = $p }
      $items = @()
      try { $items = Get-ChildItem -Path $patternPath -File -ErrorAction Stop } catch { $items = @() }
      foreach ($f in $items) { $files.Add($f.FullName) }
    }
  }

  $set = New-Object System.Collections.Generic.HashSet[string]
  $out = New-Object System.Collections.Generic.List[string]
  foreach ($f in $files) { if (-not $set.Contains($f)) { $set.Add($f) | Out-Null; $out.Add($f) } }
  return $out
}

# ===== 入力展開（単区間 or 時系列） =====
$segments = @()

if ($SnapshotFiles -and $SnapshotFiles.Count -ge 1) {
  $fileList = Expand-SnapshotInputs -Inputs $SnapshotFiles
  if ($fileList.Count -lt 2) {
    throw "SnapshotFiles: 指定から2つ以上のファイルが見つかりません。フォルダ直下に2つ以上置くか、複数ファイル/ワイルドカードを指定してください。"
  }

  $parsed = @()
  foreach ($f in $fileList) { $parsed += (Parse-RadioStatsFile -Path $f) }
  $sorted = $parsed | Sort-Object { if ($_.OutputTime -ne $null) { $_.OutputTime } else { [System.IO.File]::GetLastWriteTime($_.Path) } }

  for ($i=0; $i -lt $sorted.Count-1; $i++) {
    $b = $sorted[$i]; $a = $sorted[$i+1]
    $bt = $b.OutputTime; $at = $a.OutputTime
    $sec = 0
    if ($bt -ne $null -and $at -ne $null) { try { $sec = [int][Math]::Abs(($at - $bt).TotalSeconds) } catch { $sec = 0 } }
    if ($sec -le 0) {
      try {
        $t1 = [System.IO.File]::GetLastWriteTime($b.Path)
        $t2 = [System.IO.File]::GetLastWriteTime($a.Path)
        $sec = [int]([Math]::Abs(($t2 - $t1).TotalSeconds))
      } catch { $sec = 0 }
    }
    if ($sec -le 0) { $sec = 900 }

    $segments += (New-Object psobject -Property @{
      Before = $b; After = $a; DurationSec = $sec;
      StartJst = (Convert-ToJst $bt); EndJst = (Convert-ToJst $at)
    })
  }
}
elseif (-not [string]::IsNullOrWhiteSpace($BeforeFile) -and -not [string]::IsNullOrWhiteSpace($AfterFile)) {
  $b = Parse-RadioStatsFile -Path $BeforeFile
  $a = Parse-RadioStatsFile -Path $AfterFile
  $bt = $b.OutputTime; $at = $a.OutputTime

  if ($DurationSec -le 0) {
    $sec = 0
    if ($bt -ne $null -and $at -ne $null) { try { $sec = [int][Math]::Abs(($at - $bt).TotalSeconds) } catch { $sec = 0 } }
    if ($sec -le 0) {
      try {
        $t1 = [System.IO.File]::GetLastWriteTime($BeforeFile)
        $t2 = [System.IO.File]::GetLastWriteTime($AfterFile)
        $sec = [int]([Math]::Abs(($t2 - $t1).TotalSeconds))
      } catch { $sec = 0 }
    }
    if ($sec -le 0) { $sec = 900 }
    $DurationSec = $sec
  } else { $sec = $DurationSec }

  $segments += (New-Object psobject -Property @{
    Before = $b; After = $a; DurationSec = $sec;
    StartJst = (Convert-ToJst $bt); EndJst = (Convert-ToJst $at)
  })
}
else {
  throw "単区間比較は -BeforeFile/-AfterFile、時系列集計は -SnapshotFiles に（フォルダ/ワイルドカード/ファイルのいずれかを）1個以上指定してください（展開後2つ以上のファイルが必要）。"
}

# ===== CSV 出力先 =====
if ([string]::IsNullOrWhiteSpace($OutputCsv)) {
  $baseRef = $null
  if ($segments.Count -gt 0) { $baseRef = $segments[$segments.Count-1].After.Path } else { $baseRef = $AfterFile }
  $outDir = Get-ParentOrCwd -PathLike $baseRef
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $OutputCsv = Join-Path -Path $outDir -ChildPath ("aruba_radio_stats_diff_{0}.csv" -f $ts)
}

# ===== 列定義（CSV/HTML） SimpleDiag/Tips は含めない =====
$colDefs = @(
  @('AP','AP 名'),
  @('Radio','ラジオ番号（0/1 等）'),
  @('Channel','運用チャネル番号（取得できた場合）'),
  @('Band','2.4GHz/5GHz（チャネルから推定）'),
  @('DurationSec','比較区間の秒数'),
  @('StartJST','区間開始（日本時間）'),
  @('EndJST','区間終了（日本時間）'),
  @('RxRetry_per_s','受信再送（/s, 差分/秒）'),
  @('RxCRC_per_s','受信CRCエラー（/s, 差分/秒）'),
  @('RxPLCP_per_s','受信PLCPエラー（/s, 差分/秒）'),
  @('ChannelChanges_per_h','チャネル変更（/h）'),
  @('TxPowerChanges_per_h','送信出力変更（/h）'),
  @('Busy1s_pct','1秒平均 空中占有（%）'),
  @('Busy4s_pct','4秒平均 空中占有（%）'),
  @('Busy64s_pct','64秒平均 空中占有（%）'),
  @('BusyBeacon_pct','@beacon時の占有（%）'),
  @('TxBeacon_pct','@beacon時のTx（%）'),
  @('RxBeacon_pct','@beacon時のRx（%）'),
  @('CCA_Our_pct','CCA内訳: 自BSS（%）'),
  @('CCA_Other_pct','CCA内訳: 他BSS（%）'),
  @('CCA_Interference_pct','CCA内訳: 非Wi-Fi干渉（%）')
)

# CSV ヘッダ出力
$csvHeader = ($colDefs | ForEach-Object { $_[0] }) -join ','
Set-Content -LiteralPath $OutputCsv -Value $csvHeader -Encoding UTF8

# ===== 集計 =====
$rows = @()
$cards = @()
$hourBuckets = @{}

# 出力時刻のフォールバック関数（JST）
function Get-FileTimeJst {
  param([string]$p)
  try { return Convert-ToJst ([System.IO.File]::GetLastWriteTime($p)) } catch { return $null }
}

foreach ($seg in $segments) {
  $before = $seg.Before.Data
  $after  = $seg.After.Data
  $dur    = $seg.DurationSec

  # 表示用時刻（JST）：OutputTime優先、無ければ更新時刻にフォールバック
  $startDisp = $seg.StartJst
  if ($startDisp -eq $null -and $seg.Before -ne $null -and -not [string]::IsNullOrWhiteSpace($seg.Before.Path)) {
    $startDisp = Get-FileTimeJst -p $seg.Before.Path
  }
  $endDisp = $seg.EndJst
  if ($endDisp -eq $null -and $seg.After -ne $null -and -not [string]::IsNullOrWhiteSpace($seg.After.Path)) {
    $endDisp = Get-FileTimeJst -p $seg.After.Path
  }

  # キーの和集合
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

    $dRetry = Diff-NonNegative $a.RxRetry $b.RxRetry
    $dCRC   = Diff-NonNegative $a.RxCRC   $b.RxCRC
    $dPLCP  = Diff-NonNegative $a.RxPLCP  $b.RxPLCP
    $dChg   = Diff-NonNegative $a.ChannelChanges $b.ChannelChanges
    $dTxPw  = Diff-NonNegative $a.TxPowerChanges $b.TxPowerChanges

    $retry_ps=$null;$crc_ps=$null;$plcp_ps=$null;$chg_ph=$null;$txp_ph=$null
    if ($dRetry -ne $null) { $retry_ps=[Math]::Round($dRetry/$dur,6) }
    if ($dCRC   -ne $null) { $crc_ps  =[Math]::Round($dCRC  /$dur,6) }
    if ($dPLCP  -ne $null) { $plcp_ps =[Math]::Round($dPLCP /$dur,6) }
    if ($dChg   -ne $null) { $chg_ph  =[Math]::Round(($dChg*3600.0)/$dur,6) }
    if ($dTxPw  -ne $null) { $txp_ph  =[Math]::Round(($dTxPw*3600.0)/$dur,6) }

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
    $chan   = Pick-AfterFirst $a.Channel $b.Channel
    $band   = Get-BandFromChannel -Channel $chan

    # 空行抑止
    $hasAny=$false
    foreach($vv in @($retry_ps,$crc_ps,$plcp_ps,$chg_ph,$txp_ph,$busy1s,$busy4s,$busy64,$busyB,$txB,$rxB,$ccaO,$ccaOt,$ccaI)){
      if($vv -ne $null -and $vv -ne ''){ $hasAny=$true }
    }
    if(-not $hasAny){ continue }

    # 診断（カード用に生成）
    $d = Make-Diagnosis -Busy64 $busy64 -BusyB $busyB -TxB $txB -RxB $rxB `
                        -CCAO $ccaOt -CCAI $ccaI -Retry $retry_ps -CRC $crc_ps -PLCP $plcp_ps `
                        -ChgPH $chg_ph -TxPPH $txp_ph -Channel $chan

    # === CSV 1行 ===
    $rowObj = New-Object psobject -Property @{
      AP=$ap; Radio=$radio; Channel=$chan; Band=$band; DurationSec=$dur;
      StartJST=$(if ($startDisp -ne $null){ $startDisp.ToString('yyyy-MM-dd HH:mm:ss') } else { '' });
      EndJST  =$(if ($endDisp   -ne $null){ $endDisp.ToString('yyyy-MM-dd HH:mm:ss') } else { '' });
      RxRetry_per_s=$retry_ps; RxCRC_per_s=$crc_ps; RxPLCP_per_s=$plcp_ps;
      ChannelChanges_per_h=$chg_ph; TxPowerChanges_per_h=$txp_ph;
      Busy1s_pct=$busy1s; Busy4s_pct=$busy4s; Busy64s_pct=$busy64;
      BusyBeacon_pct=$busyB; TxBeacon_pct=$txB; RxBeacon_pct=$rxB;
      CCA_Our_pct=$ccaO; CCA_Other_pct=$ccaOt; CCA_Interference_pct=$ccaI
    }

    $vals = @()
    foreach ($def in $colDefs) {
      $name = $def[0]
      $v = $rowObj.PSObject.Properties[$name].Value
      if ($v -eq $null) { $vals += '' } else { $vals += $v.ToString() }
    }
    $escaped=@(); foreach($v in $vals){ if($v -match '[,"]'){ $escaped+=('"{0}"' -f ($v -replace '"','""')) } else { $escaped+=$v } }
    Add-Content -LiteralPath $OutputCsv -Value ($escaped -join ',') -Encoding UTF8

    # === HTMLテーブル用行 ===
    $rows += $rowObj

    # === 上部カード ===
    $cards += (New-Object psobject -Property @{
      AP=$ap; Radio=$radio; Channel=$chan; Band=$band;
      EndJST=$(if ($endDisp -ne $null){ $endDisp.ToString('yyyy-MM-dd HH:mm') } else { '' });
      Simple=$d.Simple; Tips=$d.Tips; Secondary=$d.Secondary; Severity=$d.Severity; Score=$d.Score
    })

    # === 時間帯バケット ===
    if ($endDisp -ne $null) {
      $key = $endDisp.ToString('HH')
      if (-not $hourBuckets.ContainsKey($key)) {
        $hourBuckets[$key] = New-Object psobject -Property @{ N=0; Busy64=0.0; CCAO=0.0; CCAI=0.0; PLCP=0.0; CRC=0.0; Retry=0.0 }
      }
      $h = $hourBuckets[$key]; $h.N++
      if ($busy64 -ne $null){ $h.Busy64 += $busy64 }
      if ($ccaOt -ne $null){ $h.CCAO   += $ccaOt }
      if ($ccaI  -ne $null){ $h.CCAI   += $ccaI }
      if ($plcp_ps -ne $null){ $h.PLCP += $plcp_ps }
      if ($crc_ps  -ne $null){ $h.CRC  += $crc_ps }
      if ($retry_ps- ne $null){ $h.Retry+= $retry_ps }
    }
  }
}

Write-Output ("CSV : {0}" -f $OutputCsv)

# ===== HTML 出力 =====
if (-not [string]::IsNullOrWhiteSpace($OutputHtml)) {
  $baseRef = $OutputHtml
  $outDir = Get-ParentOrCwd -PathLike $baseRef
  if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

  $firstSeg = $null; $lastSeg = $null
  if ($segments.Count -gt 0) { $firstSeg = $segments[0]; $lastSeg = $segments[$segments.Count-1] }
  $btStr = ''; $atStr = ''
  if ($firstSeg -ne $null -and $firstSeg.StartJst -ne $null) { $btStr = $firstSeg.StartJst.ToString('yyyy-MM-dd HH:mm:ss') }
  if ($lastSeg  -ne $null -and $lastSeg.EndJst   -ne $null)   { $atStr = $lastSeg.EndJst.ToString('yyyy-MM-dd HH:mm:ss') }

  $titleText = $Title
  if ([string]::IsNullOrWhiteSpace($titleText)) {
    if (-not [string]::IsNullOrWhiteSpace($btStr) -or -not [string]::IsNullOrWhiteSpace($atStr)) {
      $titleText = "Aruba Radio Stats Diff（JST） " + $btStr + " → " + $atStr
    } else { $titleText = "Aruba Radio Stats Diff（JST）" }
  }

  # 列は colDefs に準拠（SimpleDiag/Tips 列は無し）
  $cols = $colDefs

  # カード（上位10件）
  $topCards = $cards | Sort-Object -Property @{Expression='Score';Descending=$true} | Select-Object -First 10

  # 時間帯サマリ
  $hourComments = New-Object System.Text.StringBuilder
  $hours = $hourBuckets.Keys | Sort-Object
  foreach ($hh in $hours) {
    $h = $hourBuckets[$hh]
    if ($h.N -le 0) { continue }
    $avgBusy = [Math]::Round($h.Busy64 / $h.N, 1)
    $avgO    = [Math]::Round($h.CCAO   / $h.N, 1)
    $avgI    = [Math]::Round($h.CCAI   / $h.N, 1)
    $avgP    = [Math]::Round($h.PLCP   / $h.N, 1)
    $avgC    = [Math]::Round($h.CRC    / $h.N, 1)
    $avgR    = [Math]::Round($h.Retry  / $h.N, 1)

    $msg = $hh + "時台："
    $notes=@()
    if ($avgO -ge $TH_OtherHigh -and $avgBusy -ge $TH_BusyWarn) { $notes += ("同一チャネル影響が強い（他BSS " + $avgO + "%）") }
    if ($avgI -ge $TH_InterfHigh -and $avgP -ge $TH_PLCPHigh)   { $notes += ("非Wi-Fi干渉の疑い（Interf " + $avgI + "%, PLCP " + $avgP + "/s）") }
    if ($avgBusy -ge $TH_BusyHigh)                              { $notes += ("混雑が高い（Busy64 " + $avgBusy + "%）") }
    if ($notes.Count -eq 0) { $notes += "特筆なし" }
    $null = $hourComments.AppendLine( (Sanitize-Text ($msg + [string]::Join(" / ", $notes))) )
  }

  # ---- HTML ----
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
.diag-card{border:1px solid #e5e5e5;border-radius:6px;padding:8px;margin:6px 0;background:#fafafa}
.diag-card .title{font-weight:600}
.diag-card .sub{color:#444;margin-top:2px}
.diag-card .sec{color:#333;font-size:12px;margin-top:4px}
.sev-warn{background:#fff7e6}
.sev-crit{background:#ffecec}
.help{font-size:12px;color:#333;background:#f7f7f7;border:1px solid #e5e5e5;border-radius:4px;padding:8px;margin:8px 0}
  </style>')

  [void]$sb.AppendLine("<h1>{0}</h1>" -f (HtmlEscape $titleText))
  [void]$sb.AppendLine('<div class="small">※日時はすべて日本時間（JST）で表示。上は重点カード、下は明細テーブルです。</div>')

  # 上部カード（Simple/Tipsはここに表示）
  if ($topCards -and $topCards.Count -gt 0) {
    [void]$sb.AppendLine('<div id="cards">')
    foreach ($c in $topCards) {
      $cls = $c.Severity
      [void]$sb.AppendLine('<div class="diag-card '+$cls+'">')
      $titleLine = $c.AP+" / Radio "+$c.Radio
      if ($c.Channel -ne $null -and $c.Channel -ne '') { $titleLine = $titleLine + " ch" + $c.Channel }
      if (-not [string]::IsNullOrWhiteSpace($c.Band)) { $titleLine = $titleLine + " ("+$c.Band+")" }
      $titleLine = $titleLine + " @ " + $c.EndJST
      [void]$sb.AppendLine('<div class="title">'+ (HtmlEscape $titleLine) +'</div>')
      [void]$sb.AppendLine('<div class="sub">'+ (HtmlEscape $c.Simple) +'</div>')
      if ($c.Secondary -ne $null -and $c.Secondary.Count -gt 0) {
        foreach ($s in $c.Secondary) { [void]$sb.AppendLine('<div class="sec">'+ (HtmlEscape $s) +'</div>') }
      }
      [void]$sb.AppendLine('<div class="sec">'+ (HtmlEscape ("対策："+$c.Tips)) +'</div>')
      [void]$sb.AppendLine('</div>')
    }
    [void]$sb.AppendLine('</div>')
  }

  # フィルタ
  [void]$sb.AppendLine('<div style="margin:10px 0"><input id="flt" type="search" placeholder="フィルタ（AP/数値/文言）..." oninput="filterTable()"></div>')

  # 列の説明（Simple/Tipsは表には無い）
  [void]$sb.AppendLine('<details class="help"><summary>列の見方（クリックで開閉）</summary><div><ul>')
  foreach ($pair in $cols) { [void]$sb.AppendLine('<li><b>'+ (HtmlEscape $pair[0]) +'</b>：'+ (HtmlEscape $pair[1]) +'</li>') }
  [void]$sb.AppendLine('</ul></div></details>')

  # 時間帯サマリ
  $hours = $hourBuckets.Keys | Sort-Object
  if ($hours.Count -gt 0) {
    [void]$sb.AppendLine('<div class="help"><b>時間帯サマリ（JST）</b><br>')
    [void]$sb.AppendLine((HtmlEscape ($hourComments.ToString().Trim())))
    [void]$sb.AppendLine('</div>')
  }

  # 明細テーブル（Simple/Tips列なし）
  [void]$sb.AppendLine('<table id="tbl"><thead><tr>')
  foreach ($pair in $cols) { $name=$pair[0]; $desc=$pair[1]; [void]$sb.AppendLine('<th title="'+ (HtmlEscape $desc) +'">'+ (HtmlEscape $name) +'</th>') }
  [void]$sb.AppendLine('</tr></thead><tbody>')
  foreach ($r in $rows) {
    [void]$sb.AppendLine('<tr>')
    foreach ($pair in $cols) {
      $c = $pair[0]; $v = $r.PSObject.Properties[$c].Value
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
  $nameOnly = $null; try { $nameOnly = Split-Path -Path $OutputHtml -Leaf } catch { $nameOnly = $null }
  if ([string]::IsNullOrWhiteSpace($nameOnly)) { $nameOnly = "radio_stats_diff.html" }
  $htmlPath = Join-Path -Path $outDir -ChildPath $nameOnly
  Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8
  Write-Output ("HTML: {0}" -f $htmlPath)
}
exit 0