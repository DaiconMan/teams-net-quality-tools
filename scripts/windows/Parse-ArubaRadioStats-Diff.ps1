<# 
.SYNOPSIS
  Aruba "show ap debug radio-stats" スナップショットの差分/秒を算出し、CSV/HTML を生成（JST表示）。
  - AP名は「**ファイルの親フォルダ名**」を採用（radio-stats 内に AP 名が無い前提）。
  - 同フォルダの "show ap bss-table" / "show aps" / "show ap debug radio-info" / "show ap arm history"
    テキストから Channel / Band を補完（バックアップ参照）
  - -SnapshotFiles はフォルダ/ワイルドカード/ファイル混在OK（展開後2ファイル以上）
  - radio-stats 以外（arm history 等）はスナップショット対象外（補完専用）
  - LAA/NR-U 候補検出、時間帯コメント
  - PowerShell 5.1対応・三項演算子不使用・OneDrive/日本語/スペース対応・呼び出し演算子(&)不使用
#>
[CmdletBinding()]
param(
  [string]$BeforeFile,
  [string]$AfterFile,
  [int]$DurationSec,
  [string[]]$SnapshotFiles,
  [string]$OutputCsv,
  [string]$OutputHtml,
  [string]$Title,
  [string]$LogFile
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

# ===== ログ基盤 =====
$Script:LogPath = $null
function Write-Log { param([string]$Msg)
  if ([string]::IsNullOrWhiteSpace($Script:LogPath)) { return }
  $ts = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffK"
  $line = "[{0}] {1}" -f $ts, $Msg
  try { Add-Content -LiteralPath $Script:LogPath -Value $line -Encoding UTF8 } catch {}
}
function Init-Log { param([string]$Preferred,[string]$OutputHtmlPath,[string]$OutputCsvPath)
  if (-not [string]::IsNullOrWhiteSpace($Preferred)) {
    $Script:LogPath = $Preferred
    try {
      $dir = Split-Path -Path $Script:LogPath -Parent
      if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
      Set-Content -LiteralPath $Script:LogPath -Value "# Aruba Radio Stats Diff Log (UTF-8)" -Encoding UTF8
    } catch {}
    return
  }
  $baseDir = $null
  if (-not [string]::IsNullOrWhiteSpace($OutputHtmlPath)) {
    try { $baseDir = Split-Path -Path $OutputHtmlPath -Parent } catch { $baseDir = $null }
  }
  if ([string]::IsNullOrWhiteSpace($baseDir) -and -not [string]::IsNullOrWhiteSpace($OutputCsvPath)) {
    try { $baseDir = Split-Path -Path $OutputCsvPath -Parent } catch { $baseDir = $null }
  }
  if ([string]::IsNullOrWhiteSpace($baseDir)) {
    try { $baseDir = (Get-Location).Path } catch { $baseDir = "." }
  }
  if (-not (Test-Path -LiteralPath $baseDir)) {
    try { New-Item -ItemType Directory -Path $baseDir -Force | Out-Null } catch {}
  }
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $Script:LogPath = Join-Path -Path $baseDir -ChildPath ("aruba_radio_stats_diff_LOG_{0}.txt" -f $ts)
  try { Set-Content -LiteralPath $Script:LogPath -Value "# Aruba Radio Stats Diff Log (UTF-8)" -Encoding UTF8 } catch {}
}
function Get-ErrMsg { param($e)
  try {
    if ($e -ne $null) {
      if ($e.Exception -ne $null -and -not [string]::IsNullOrWhiteSpace($e.Exception.Message)) { return $e.Exception.Message }
      return [string]$e
    }
  } catch {}
  return ''
}

function ToLogStr-DT { param([Nullable[DateTime]]$dt,[string]$Label)
  if ($dt -eq $null) { return ("{0}=<null>" -f $Label) }
  $v = $dt.Value
  $kind = $v.Kind.ToString()
  $iso  = $v.ToString("yyyy-MM-dd HH:mm:ss 'Kind='") + $kind
  return ("{0}={1}" -f $Label, $iso)
}

# ===== JST 変換 =====
function Convert-ToJst {
  param([Nullable[DateTime]]$dt,[string]$Context)
  if ($dt -eq $null) { Write-Log ("Convert-ToJst: {0} input=<null>" -f $Context); return $null }
  $tz = $null
  try { $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById("Tokyo Standard Time"); Write-Log ("Convert-ToJst: TZ='Tokyo Standard Time' found") } catch { Write-Log ("Convert-ToJst: TZ lookup failed, will +9h fallback") ; $tz = $null }
  try {
    $v = $dt.Value
    Write-Log ("Convert-ToJst: {0} IN {1}" -f $Context, (ToLogStr-DT $v "dt"))
    if     ($v.Kind -eq [System.DateTimeKind]::Unspecified) { Write-Log "Convert-ToJst: Kind=Unspecified → Assume UTC"; $v = [DateTime]::SpecifyKind($v, [System.DateTimeKind]::Utc) }
    elseif ($v.Kind -eq [System.DateTimeKind]::Local)       { Write-Log "Convert-ToJst: Kind=Local → ConvertTimeToUtc"; $v = [System.TimeZoneInfo]::ConvertTimeToUtc($v) }
    if ($tz -ne $null) { $ret=[System.TimeZoneInfo]::ConvertTimeFromUtc($v, $tz); Write-Log ("Convert-ToJst: {0} OUT {1}" -f $Context, (ToLogStr-DT $ret "jst")); return $ret }
    else               { $ret=$v.AddHours(9); Write-Log ("Convert-ToJst: {0} OUT(+9h fallback) {1}" -f $Context, (ToLogStr-DT $ret "jst")); return $ret }
  } catch {
    Write-Log ("Convert-ToJst: Exception → +9h fallback. " + (Get-ErrMsg $_))
    try { $ret=$dt.Value.AddHours(9); Write-Log ("Convert-ToJst: {0} OUT(+9h after err) {1}" -f $Context, (ToLogStr-DT $ret "jst")); return $ret } catch { return $dt }
  }
}

# ===== テキスト補助 =====
function HtmlEscape { param([string]$s) if ($null -eq $s) { return '' }
  $r=$s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;').Replace("'",'&#39;'); return $r
}
function Sanitize-Text { param([string]$s) if ($null -eq $s) { return '' } return $s.Replace('"','"') }
function Get-ParentOrCwd { param([string]$PathLike)
  $dir=$null
  if (-not [string]::IsNullOrWhiteSpace($PathLike)) {
    if (Test-Path -LiteralPath $PathLike) { try { $dir=Split-Path -LiteralPath $PathLike -Parent } catch { $dir=$null } }
    else { try { $dir=Split-Path -Path $PathLike -Parent } catch { $dir=$null } }
  }
  if ([string]::IsNullOrWhiteSpace($dir)) { try { return (Get-Location).Path } catch { return "." } }
  return $dir
}
function Normalize-WS { param([string]$s)
  if ($null -eq $s) { return $s }
  $s = $s.Replace([char]0x00A0,' ')
  $s = $s.Replace([char]0x2007,' ')
  $s = $s.Replace([char]0x202F,' ')
  $s = $s.Trim()
  $s = ($s -replace '\s+',' ')
  return $s
}

# ===== AP名クレンジング / フォルダから取得 =====
function Is-BadAPToken { param([string]$tok)
  if ([string]::IsNullOrWhiteSpace($tok)) { return $true }
  $t=$tok.ToLower()
  foreach($x in @('radio','stats','radio-stats','radio_stats','show','debug','output','log','arm','history','ap','apname','stat','txt','log')){
    if ($t -eq $x) { return $true }
  }
  if ($t.Length -le 1) { return $true }
  return $false
}
function Clean-APLeaf { param([string]$leaf,[string]$context)
  if ([string]::IsNullOrWhiteSpace($leaf)) { return '' }
  $orig=$leaf
  $base=$leaf
  try { $base=[System.IO.Path]::GetFileNameWithoutExtension($base) } catch {}
  $base = ($base -replace '(?i)\b(show|ap|debug|radio|stats?|output|log|arm|history)\b',' ').Trim()
  $base = ($base -replace '[_\-\.\s]+',' ')
  $parts = @(); if (-not [string]::IsNullOrWhiteSpace($base)) { $parts = $base.Split(' ') }
  Write-Log ("AP Clean: leaf='{0}' base='{1}' parts='{2}' ctx='{3}'" -f $orig,$base,([string]::Join('|',$parts)),$context)
  $cands=@()
  foreach($p in $parts){ if (-not (Is-BadAPToken $p)) { $cands += $p } }
  if ($cands.Count -gt 0) {
    $best=$cands[0]
    for($i=1;$i -lt $cands.Count;$i++){ if($cands[$i].Length -gt $best.Length){ $best=$cands[$i] } }
    Write-Log ("AP Clean: choose='{0}'" -f $best)
    return $best
  }
  return ''
}
function Get-APName-From-Folder { param([string]$Path)
  try { $dir = Split-Path -Path $Path -Parent } catch { $dir = $null }
  if ([string]::IsNullOrWhiteSpace($dir)) { return '' }
  try { $leaf = Split-Path -Path $dir -Leaf } catch { $leaf = $null }
  if ([string]::IsNullOrWhiteSpace($leaf)) { return '' }
  $clean = Clean-APLeaf -leaf $leaf -context "folder"
  Write-Log ("AP FromFolder: dir='{0}' leaf='{1}' clean='{2}'" -f $dir,$leaf,$clean)
  return $clean
}

# ===== 数値/Busy抽出 =====
function Get-LastNumber { param([string]$Line)
  if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
  $m=[regex]::Matches($Line,'(-?\d+(?:\.\d+)?)'); if($m.Count -gt 0){ return [double]$m[$m.Count-1].Value } return $null
}
function TryExtractPercentTriplet {
  param([string]$Line,[ref]$Busy1s,[ref]$Busy4s,[ref]$Busy64s)
  $ok=$false
  $m1=[regex]::Match($Line,'\b1s\b[^0-9\-]*(-?\d+(?:\.\d+)?)(?:\s*%)?'); if($m1.Success){ $Busy1s.Value=[double]$m1.Groups[1].Value; $ok=$true }
  $m4=[regex]::Match($Line,'\b4s\b[^0-9\-]*(-?\d+(?:\.\d+)?)(?:\s*%)?'); if($m4.Success){ $Busy4s.Value=[double]$m4.Groups[1].Value; $ok=$true }
  $m64=[regex]::Match($Line,'\b64s\b[^0-9\-]*(-?\d+(?:\.\d+)?)(?:\s*%)?'); if($m64.Success){ $Busy64s.Value=[double]$m64.Groups[1].Value; $ok=$true }
  if ($ok) { return $true }
  $m=[regex]::Match($Line,'1s[^0-9]{0,5}(\d+(?:\.\d+)?).{0,12}4s[^0-9]{0,5}(\d+(?:\.\d+)?).{0,12}64s[^0-9]{0,5}(\d+(?:\.\d+)?)')
  if ($m.Success) { $Busy1s.Value=[double]$m.Groups[1].Value; $Busy4s.Value=[double]$m.Groups[2].Value; $Busy64s.Value=[double]$m.Groups[3].Value; return $true }
  return $false
}

# ===== Output Time (UTC) 抽出（堅牢版） =====
function Extract-OutputTime {
  param([string[]]$Lines,[string]$Path)
  if ($Lines -eq $null -or $Lines.Count -eq 0) { return $null }
  foreach ($raw in $Lines) {
    $line = ($raw -replace '\r','').Trim()
    if ($line -match '(?i)Output\s*Time\s*[:=]\s*(.+)$') {
      $rhsRaw = $Matches[1]
      $rhs = Normalize-WS $rhsRaw
      Write-Log ("OutputTime: Path='{0}' RawLine='{1}' Rhs(norm)='{2}'" -f $Path,$line,$rhs)

      $core = $null; $suffix = $null
      $mCore=[regex]::Match($rhs,'^(?<core>\d{4}[-/]\d{2}[-/]\d{2}[ T]\d{2}:\d{2}:\d{2}(?:\.\d{1,6})?)\s*(?<suffix>.*)$')
      if($mCore.Success){ $core=$mCore.Groups['core'].Value; $suffix=$mCore.Groups['suffix'].Value.Trim(); Write-Log ("OutputTime: core='{0}' suffix='{1}'" -f $core,$suffix) }

      $dto=[System.DateTimeOffset]::MinValue
      if ([System.DateTimeOffset]::TryParse($rhs,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::AssumeUniversal,[ref]$dto)) {
        $ret=$dto.UtcDateTime; Write-Log ("OutputTime: match=iso8601 parsed {0}" -f ((ToLogStr-DT $ret "utc"))); return $ret
      }

      if (-not [string]::IsNullOrWhiteSpace($core) -and $suffix -match '^(?i)(UTC|GMT)\b') {
        try{
          $fmt=@('yyyy-MM-dd HH:mm:ss','yyyy-MM-dd''T''HH:mm:ss','yyyy/MM/dd HH:mm:ss','yyyy/MM/dd''T''HH:mm:ss','yyyy-MM-dd HH:mm:ss.fff','yyyy-MM-dd''T''HH:mm:ss.fff','yyyy/MM/dd HH:mm:ss.fff','yyyy/MM/dd''T''HH:mm:ss.fff')
          $dtOut=[DateTime]::MinValue
          $ok=[DateTime]::TryParseExact($core,$fmt,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$dtOut)
          if($ok){ $ret=[DateTime]::SpecifyKind($dtOut,[System.DateTimeKind]::Utc); Write-Log ("OutputTime: match=core+UTC parsed {0}" -f ((ToLogStr-DT $ret "utc"))); return $ret }
          else { Write-Log ("OutputTime: core matched but ParseExact failed for '{0}'" -f $core) }
        }catch{ Write-Log ("OutputTime: parse err (core+UTC): " + (Get-ErrMsg $_)) }
      }

      if (-not [string]::IsNullOrWhiteSpace($core) -and $suffix -match '^(?i)(JST|JDT)\b') {
        try{
          $fmt2=@('yyyy-MM-dd HH:mm:ss','yyyy-MM-dd''T''HH:mm:ss','yyyy/MM/dd HH:mm:ss','yyyy/MM/dd''T''HH:mm:ss','yyyy-MM-dd HH:mm:ss.fff','yyyy-MM-dd''T''HH:mm:ss.fff','yyyy/MM/dd HH:mm:ss.fff','yyyy/MM/dd''T''HH:mm:ss.fff')
          $dtJst=[DateTime]::MinValue
          $ok2=[DateTime]::TryParseExact($core,$fmt2,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$dtJst)
          if($ok2){
            $dtJst = [DateTime]::SpecifyKind($dtJst,[System.DateTimeKind]::Unspecified)
            $utc = $dtJst.AddHours(-9)
            $utc = [DateTime]::SpecifyKind($utc,[System.DateTimeKind]::Utc)
            Write-Log ("OutputTime: match=core+JST parsed (JST->UTC) {0}" -f ((ToLogStr-DT $utc "utc")))
            return $utc
          } else { Write-Log ("OutputTime: core+JST matched but ParseExact failed for '{0}'" -f $core) }
        }catch{ Write-Log ("OutputTime: parse err (core+JST): " + (Get-ErrMsg $_)) }
      }

      if ($rhs -match '^\d{10}(\.\d+)?$') {
        try { $sec=[long]([double]$rhs); $ret=([System.DateTimeOffset]::FromUnixTimeSeconds($sec)).UtcDateTime; Write-Log ("OutputTime: match=epoch-seconds parsed {0}" -f ((ToLogStr-DT $ret "utc"))); return $ret } catch { Write-Log ("OutputTime: epoch-seconds parse err: " + (Get-ErrMsg $_)) }
      }
      if ($rhs -match '^\d{13}$') {
        try { $ms=[long]$rhs; $ret=([System.DateTimeOffset]::FromUnixTimeMilliseconds($ms)).UtcDateTime; Write-Log ("OutputTime: match=epoch-millis parsed {0}" -f ((ToLogStr-DT $ret "utc"))); return $ret } catch { Write-Log ("OutputTime: epoch-millis parse err: " + (Get-ErrMsg $_)) }
      }

      if (-not [string]::IsNullOrWhiteSpace($core)) {
        try{
          $fmt3=@('yyyy-MM-dd HH:mm:ss','yyyy-MM-dd''T''HH:mm:ss','yyyy/MM/dd HH:mm:ss','yyyy/MM/dd''T''HH:mm:ss','yyyy-MM-dd HH:mm:ss.fff','yyyy-MM-dd''T''HH:mm:ss.fff','yyyy/MM/dd HH:mm:ss.fff','yyyy/MM/dd''T''HH:mm:ss.fff')
          $dtOut2=[DateTime]::MinValue
          $ok3=[DateTime]::TryParseExact($core,$fmt3,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$dtOut2)
          if($ok3){
            $ret=[DateTime]::SpecifyKind($dtOut2,[System.DateTimeKind]::Utc)
            Write-Log ("OutputTime: match=core(no suffix) as UTC parsed {0}" -f ((ToLogStr-DT $ret "utc")))
            return $ret
          } else {
            Write-Log ("OutputTime: core(no suffix) ParseExact failed for '{0}'" -f $core)
          }
        }catch{ Write-Log ("OutputTime: parse err (core only): " + (Get-ErrMsg $_)) }
      }

      try{
        $cleanRhs = ($rhs -replace '(?i)\b(UTC|GMT|JST|JDT)\b','').Trim()
        $cleanRhs = ($cleanRhs -replace '\s+',' ')
        Write-Log ("OutputTime: fallback cleanRhs='{0}'" -f $cleanRhs)
        $fmt4=@('yyyy-MM-dd HH:mm:ss','yyyy-MM-dd''T''HH:mm:ss','yyyy/MM/dd HH:mm:ss','yyyy/MM/dd''T''HH:mm:ss','yyyy-MM-dd HH:mm:ss.fff','yyyy-MM-dd''T''HH:mm:ss.fff','yyyy/MM/dd HH:mm:ss.fff','yyyy/MM/dd''T''HH:mm:ss.fff')
        $dtOut3=[DateTime]::MinValue
        $ok4=[DateTime]::TryParseExact($cleanRhs,$fmt4,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::None,[ref]$dtOut3)
        if($ok4){
          $isJst = ($suffix -match '^(?i)(JST|JDT)\b') -or ($rhs -match '^(?i).*\b(JST|JDT)\b')
          if($isJst){
            $utc2 = [DateTime]::SpecifyKind($dtOut3,[System.DateTimeKind]::Unspecified).AddHours(-9)
            $utc2 = [DateTime]::SpecifyKind($utc2,[System.DateTimeKind]::Utc)
            Write-Log ("OutputTime: fallback parsed as JST->UTC {0}" -f ((ToLogStr-DT $utc2 "utc")))
            return $utc2
          } else {
            $ret2=[DateTime]::SpecifyKind($dtOut3,[System.DateTimeKind]::Utc)
            Write-Log ("OutputTime: fallback parsed as UTC {0}" -f ((ToLogStr-DT $ret2 "utc")))
            return $ret2
          }
        } else {
          Write-Log ("OutputTime: fallback ParseExact failed for cleanRhs='{0}'" -f $cleanRhs)
        }
      }catch{ Write-Log ("OutputTime: final fallback err: " + (Get-ErrMsg $_)) }

      Write-Log ("OutputTime: all patterns failed for Rhs='{0}'" -f $rhs)
    }
  }
  Write-Log ("OutputTime: not found in '{0}'" -f $Path)
  return $null
}

# ===== 種別判定 =====
function Is-RadioStatsLines { param([string[]]$Lines)
  if ($Lines -eq $null -or $Lines.Count -eq 0) { return $false }
  foreach ($raw in $Lines) {
    $line = ($raw -replace '\r','')
    if ($line -match '(?i)show\s+ap\s+debug\s+radio-stats') { return $true }
    if ($line -match '(?i)\bChannel\s*busy\b')              { return $true }
    if ($line -match '(?i)\bRx\s*retry\b')                  { return $true }
    if ($line -match '(?i)\bRX\s*PLCP\s*Errors?\b')         { return $true }
    if ($line -match '(?i)\bCCA\b')                         { return $true }
  }
  return $false
}

# ===== Band/Channel 補助 =====
function Get-BandFromChannel { param([Nullable[int]]$Channel)
  if ($Channel -eq $null) { return '' }
  $ch=[int]$Channel
  if ($ch -ge 1  -and $ch -le 14)  { return '2.4GHz' }
  if ($ch -ge 32 -and $ch -le 196) { return '5GHz' }
  return ''
}
function Guess-Band-From-RadioIndex { param([string]$Radio)
  if ([string]::IsNullOrWhiteSpace($Radio)) { return '' }
  if ($Radio -eq '0') { return '2.4GHz' }
  if ($Radio -eq '1') { return '5GHz' }
  return ''
}
function Map-PhyType-To-Band { param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return '' }
  $t = $s.ToLower()
  if ($t -match '6ghz' -or $t -match '\b6g\b' -or $t -match '\b6e\b' -or $t -match '\beht\b') { return '6GHz' }
  if ($t -match '5ghz' -or $t -match '\b5g\b' -or $t -match '\b11a\b' -or $t -match '\bvht\b' -or $t -match '\bhe\b') { return '5GHz' }
  if ($t -match '2\.4ghz' -or $t -match '\b2g\b' -or $t -match '\b11b\b' -or $t -match '\b11g\b') { return '2.4GHz' }
  return ''
}
function New-BackupSlot { return New-Object psobject -Property @{ Ch24=$null; T24=$null; Ch5=$null; T5=$null; Ch6=$null; T6=$null } }
function Detect-Band-Token { param([string]$line,[Nullable[int]]$ch)
  $s=$line.ToLower()
  if ($s -match '6ghz' -or $s -match '\b6g\b' -or $s -match '6e') { return '6GHz' }
  if ($s -match '\b11a\b' -or $s -match 'a/n' -or $s -match 'vht' -or $s -match '\bhe\b' -or $s -match '\beht\b') { return '5GHz' }
  if ($s -match '\b11b\b' -or $s -match '\b11g\b' -or $s -match '2\.4ghz' -or $s -match '\b2g\b') { return '2.4GHz' }
  if ($ch -ne $null) { return (Get-BandFromChannel $ch) }
  return ''
}

function Update-Backup { param([hashtable]$Backup,[string]$ap,[string]$band,[Nullable[int]]$ch,[Nullable[DateTime]]$ts)
  if ([string]::IsNullOrWhiteSpace($ap) -or [string]::IsNullOrWhiteSpace($band) -or $ch -eq $null) { return }
  if (-not $Backup.ContainsKey($ap)) { $Backup[$ap] = New-BackupSlot }
  $slot = $Backup[$ap]
  if     ($band -eq '2.4GHz') { if ($slot.T24 -eq $null -or ($ts -ne $null -and $ts -gt $slot.T24)) { $slot.Ch24=[int]$ch; $slot.T24=$ts; Write-Log ("Backup Update: AP='{0}' Band=2.4GHz Ch={1} ts={2}" -f $ap,$ch,$ts) } }
  elseif ($band -eq '5GHz')   { if ($slot.T5  -eq $null -or ($ts -ne $null -and $ts -gt $slot.T5 )) { $slot.Ch5 =[int]$ch; $slot.T5 =$ts; Write-Log ("Backup Update: AP='{0}' Band=5GHz Ch={1} ts={2}"  -f $ap,$ch,$ts) } }
  elseif ($band -eq '6GHz')   { if ($slot.T6  -eq $null -or ($ts -ne $null -and $ts -gt $slot.T6 )) { $slot.Ch6 =[int]$ch; $slot.T6 =$ts; Write-Log ("Backup Update: AP='{0}' Band=6GHz Ch={1} ts={2}"  -f $ap,$ch,$ts) } }
}

# ===== バックアップパーサ =====
function Parse-BssTable-File { param([string]$Path,[hashtable]$Backup)
  Write-Log ("Backup Parse: BSS-Table '{0}'" -f $Path)
  $ts=$null; try{ $ts=[System.IO.File]::GetLastWriteTime($Path) }catch{}
  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ return }
  foreach ($raw in $lines) {
    $line=($raw -replace '\r','').Trim(); if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $ch=$null; $mCh=[regex]::Match($line,'(?i)\bCh(?:annel)?\s*[:=]?\s*(\d{1,3})\b'); if($mCh.Success){ try{$ch=[int]$mCh.Groups[1].Value}catch{} }
    if ($ch -eq $null) { $mCh2=[regex]::Match($line,'(?i)\b(\d{1,3})\b\s*(?:MHz|HT|EIRP)'); if($mCh2.Success){ try{$ch=[int]$mCh2.Groups[1].Value}catch{} } }
    $band = Detect-Band-Token -line $line -ch $ch
    if ([string]::IsNullOrWhiteSpace($band) -and $ch -ne $null) { $band = Get-BandFromChannel $ch }
    if ($ch -eq $null -or [string]::IsNullOrWhiteSpace($band)) { continue }
    # AP 名は bss-table からは拾わない（radio-stats と一意に紐づかないため）
    # ここではチャネル情報のみをバックアップとして更新しない（曖昧回避）。
  }
}
function Parse-APS-File { param([string]$Path,[hashtable]$Backup)
  Write-Log ("Backup Parse: APS '{0}'" -f $Path)
  $ts=$null; try{ $ts=[System.IO.File]::GetLastWriteTime($Path) }catch{}
  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ return }
  $currentAP=''
  foreach ($raw in $lines) {
    $line=($raw -replace '\r','').Trim(); if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $mHead=[regex]::Match($line,'(?i)^(Name|AP\s*Name)\s*[:=]?\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)'); if($mHead.Success){ $currentAP=$mHead.Groups[2].Value; Write-Log ("APS: AP='{0}' (header)" -f $currentAP) }
    $mInline=[regex]::Match($line,'(?i)\bAP\s*Name\s*[:=]\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)'); if($mInline.Success){ $currentAP=$mInline.Groups[1].Value; Write-Log ("APS: AP='{0}' (inline)" -f $currentAP) }
    if ([string]::IsNullOrWhiteSpace($currentAP)) { continue }
    $mR=[regex]::Match($line,'(?i)\bRadio\s*([012])\s*.*?\bChannel\s*[:=]\s*(\d{1,3})')
    if($mR.Success){
      $r=$mR.Groups[1].Value; $ch=$null; try{$ch=[int]$mR.Groups[2].Value}catch{}
      if($ch -ne $null){ $band=Get-BandFromChannel $ch; if([string]::IsNullOrWhiteSpace($band)){ $band=Guess-Band-From-RadioIndex $r }; Update-Backup -Backup $Backup -ap $currentAP -band $band -ch $ch -ts $ts; continue }
    }
    $mCh=[regex]::Match($line,'(?i)\bChannel\s*[:=]\s*(\d{1,3})')
    if($mCh.Success){ $ch=$null; try{$ch=[int]$mCh.Groups[1].Value}catch{}; if($ch -ne $null){ $band=Get-BandFromChannel $ch; if([string]::IsNullOrWhiteSpace($band)){ $band=Detect-Band-Token -line $line -ch $ch }; Update-Backup -Backup $Backup -ap $currentAP -band $band -ch $ch -ts $ts } }
  }
}
function Parse-RadioInfo-File { param([string]$Path,[hashtable]$Backup)
  Write-Log ("Backup Parse: Radio-Info '{0}'" -f $Path)
  $ts=$null; try{ $ts=[System.IO.File]::GetLastWriteTime($Path) }catch{}
  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ return }
  $ap=''; foreach ($raw in $lines) {
    $line=($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $m1=[regex]::Match($line,'(?i)\bAP\s*Name\s*[:=]\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)'); if($m1.Success){ $ap=$m1.Groups[1].Value; Write-Log ("Radio-Info: AP='{0}'" -f $ap); break }
  }
  if ([string]::IsNullOrWhiteSpace($ap)) {
    try{ $ap=Clean-APLeaf -leaf (Split-Path -Path (Split-Path -Path $Path -Parent) -Leaf) -context "radio-info-folder" }catch{}
  }
  $curRadio=''
  foreach ($raw in $lines) {
    $line=($raw -replace '\r','').Trim()
    if ($line -match '(?i)\bRadio\s*([012])\b') { $curRadio=$Matches[1] }
    $mCh=[regex]::Match($line,'(?i)\bChannel\s*[:=]\s*(\d{1,3})'); if($mCh.Success){
      $ch=$null; try{$ch=[int]$mCh.Groups[1].Value}catch{}; if($ch -ne $null){
        $band=Get-BandFromChannel $ch; if([string]::IsNullOrWhiteSpace($band)){ $band=Guess-Band-From-RadioIndex $curRadio }
        Update-Backup -Backup $Backup -ap $ap -band $band -ch $ch -ts $ts
      }
    }
  }
}
function Parse-ARMHistory-File { param([string]$Path,[hashtable]$Backup)
  Write-Log ("Backup Parse: ARM-History '{0}'" -f $Path)
  $ts=$null; try{ $ts=[System.IO.File]::GetLastWriteTime($Path) }catch{}
  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ return }
  $ap = Get-APName-From-Folder -Path $Path
  if ([string]::IsNullOrWhiteSpace($ap)) { try{ $ap=Clean-APLeaf -leaf (Split-Path -Path (Split-Path -Path $Path -Parent) -Leaf) -context "arm-folder" }catch{} }
  $curBand=''
  foreach ($raw in $lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $mPhy=[regex]::Match($line,'(?i)\bPhy[-\s]?Type\s*[:=]\s*([A-Za-z0-9\.]+)'); if($mPhy.Success){ $b = Map-PhyType-To-Band $mPhy.Groups[1].Value; if (-not [string]::IsNullOrWhiteSpace($b)) { $curBand = $b; Write-Log ("ARM: AP='{0}' BandHint='{1}'" -f $ap,$curBand) } }
    $ch=$null
    $mCh1=[regex]::Match($line,'(?i)\bChannel\s*[:=]\s*(\d{1,3})\b')
    if ($mCh1.Success) { try { $ch=[int]$mCh1.Groups[1].Value } catch {} }
    if ($ch -eq $null) {
      $mCh2=[regex]::Match($line,'(?i)\b(Current|New)\s*Channel\s*[:\s]\s*(\d{1,3})\b')
      if ($mCh2.Success) { try { $ch=[int]$mCh2.Groups[2].Value } catch {} }
    }
    if ($ch -eq $null) {
      $mCh3=[regex]::Match($line,'(?i)\bch(?:annel)?\b[^\d]{0,6}(\d{1,3})\b')
      if ($mCh3.Success) { try { $ch=[int]$mCh3.Groups[1].Value } catch {} }
    }
    if ($ch -ne $null) {
      $band=$curBand
      if ([string]::IsNullOrWhiteSpace($band)) { $band = Get-BandFromChannel $ch }
      if (-not [string]::IsNullOrWhiteSpace($ap) -and -not [string]::IsNullOrWhiteSpace($band)) { Update-Backup -Backup $Backup -ap $ap -band $band -ch $ch -ts $ts }
    }
  }
}
function Build-Backup-From-Dirs { param([string[]]$Dirs)
  $bk=@{}; if ($Dirs -eq $null) { return $bk }
  foreach ($d in $Dirs) {
    if ([string]::IsNullOrWhiteSpace($d)) { continue }
    if (-not (Test-Path -LiteralPath $d)) { continue }
    Write-Log ("Backup scan dir: {0}" -f $d)
    $files=@(); try{ $files=Get-ChildItem -LiteralPath $d -File -ErrorAction Stop }catch{ $files=@() }
    foreach ($f in $files) {
      $name=$f.Name.ToLower()
      $isBss   = ($name -match 'bss') -or ($name -match 'bss\-table') -or ($name -match 'show.*bss')
      $isAPS   = ($name -match '\baps\b') -or ($name -match 'show.*aps')
      $isRInfo = ($name -match 'radio-info') -or ($name -match 'show.*radio.*info')
      $isARM   = ($name -match 'arm' -and $name -match 'history')
      if (-not ($isBss -or $isAPS -or $isRInfo -or $isARM)) {
        $peek=@(); try{ $peek=Get-Content -LiteralPath $f.FullName -Encoding UTF8 -TotalCount 40 }catch{ $peek=@() }
        foreach ($l in $peek) {
          $ls=($l -replace '\r','').ToLower()
          if ($ls -match 'bssid' -and $ls -match 'ssid') { $isBss=$true; break }
          if ($ls -match 'ap name' -and ($ls -match 'ip' -or $ls -match 'group')) { $isAPS=$true; break }
          if ($ls -match 'radio' -and $ls -match 'channel') { $isRInfo=$true; break }
          if (($ls -match 'phy' -and $ls -match 'type') -or ($ls -match 'interface:' -and $ls -match 'wifi')) { $isARM=$true; break }
        }
      }
      if     ($isAPS)   { Parse-APS-File       -Path $f.FullName -Backup $bk }
      elseif ($isRInfo) { Parse-RadioInfo-File -Path $f.FullName -Backup $bk }
      elseif ($isARM)   { Parse-ARMHistory-File -Path $f.FullName -Backup $bk }
      # bss-table はチャネル確定キーに乏しいため、ここでは参照のみ（Updateしない）
    }
  }
  return $bk
}

# ---- Helper: Ensure-RadioSlot (top-level) ----
function Ensure-RadioSlot {
  param(
    [hashtable]$Data,
    [string]$ApName,
    [string]$Radio
  )
  if ($null -eq $Data) { $Data = @{} }
  $rk = ($ApName + '|' + $Radio)
  if (-not $Data.ContainsKey($rk)) {
    $Data[$rk] = New-Object psobject -Property @{
      AP=$ApName; Radio=$Radio; Channel=$null;
      RxRetry=$null; RxCRC=$null; RxPLCP=$null;
      ChannelChanges=$null; TxPowerChanges=$null;
      Busy1s=$null; Busy4s=$null; Busy64s=$null;
      BusyBeacon=$null; TxBeacon=$null; RxBeacon=$null;
      CCA_Our=$null; CCA_Other=$null; CCA_Interference=$null
    }
    Write-Log ("Parse-RadioStatsFile: init slot AP='{0}' Radio='{1}'" -f $ApName,$Radio)
  }
  return $Data[$rk]
}

# ===== radio-stats パーサ（AP名はフォルダ） =====
function Parse-RadioStatsFile {
  param([string]$Path)

  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ $lines=@() }
  if (-not (Is-RadioStatsLines -Lines $lines)) {
    Write-Log ("Parse-RadioStatsFile: not radio-stats '{0}'" -f $Path)
    return New-Object psobject -Property @{ Path=$Path; OutputTime=$null; Data=@{}; IsRadio=$false }
  }

  $otUtc   = Extract-OutputTime -Lines $lines -Path $Path

  # ★AP名は親フォルダ名を最優先採用
  $apName  = Get-APName-From-Folder -Path $Path
  if ([string]::IsNullOrWhiteSpace($apName)) {
    # フォールバック：ファイル名からのクレンジング（最終手段）
    try { $apName = Clean-APLeaf -leaf (Split-Path -Path $Path -Leaf) -context "filename" } catch { $apName = '' }
  }
  if ([string]::IsNullOrWhiteSpace($apName)) { $apName = 'Unknown' }
  Write-Log ("AP Resolve: AP='{0}' (from folder/filename), file='{1}'" -f $apName,$Path)

  $data=@{}; $currentRadio=$null; $seenAnyRadio=$false

  foreach ($raw in $lines) {
    $line=($raw -replace '\r','').Trim(); if ([string]::IsNullOrWhiteSpace($line)) { continue }

    $switched = $false
    if ($line -match '(?i)show\s+ap\s+debug\s+radio-stats\s+([012])\b') { $currentRadio=$Matches[1]; $switched=$true }
    elseif ($line -match '(?i)^\s*radio\s*([012])\b')                  { $currentRadio=$Matches[1]; $switched=$true }
    elseif ($line -match '(?i)\binterface\s*[:=]\s*wifi([01])\b')      { $currentRadio=$Matches[1]; $switched=$true }
    elseif ($line -match '(?i)\bRadio\s*([012])\b')                     { $currentRadio=$Matches[1]; $switched=$true }
    if ($switched) { $seenAnyRadio = $true; Write-Log ("Parse-RadioStatsFile: switch -> Radio {0}" -f $currentRadio) }

    if ([string]::IsNullOrWhiteSpace($currentRadio)) { continue }

    $obj = (Ensure-RadioSlot -Data $data -ApName $apName -Radio $currentRadio)

    if ($line -notmatch '(?i)Channel\s+Changes') {
      $mCh1=[regex]::Match($line,'(?i)\bCurrent\s*Channel\s*[:=]\s*(\d{1,3})\b')
      $mCh2=[regex]::Match($line,'(?i)\bChannel\s*[:=]\s*(\d{1,3})\b')
      if ($mCh1.Success -or $mCh2.Success) {
        $val = $null
        if ($mCh1.Success) { try { $val=[int]$mCh1.Groups[1].Value } catch {} }
        if ($val -eq $null -and $mCh2.Success) { try { $val=[int]$mCh2.Groups[1].Value } catch {} }
        if ($val -ne $null) { $obj.Channel=$val; Write-Log ("Parse-RadioStatsFile: Channel from body AP='{0}' Radio='{1}' Ch={2}" -f $apName,$currentRadio,$val) }
      }
    }

    if ($line -match '(?i)\bRx\s*retry\s*frames?\b')  { $n=Get-LastNumber $line; if($n -ne $null){ $obj.RxRetry=[double]$n } }
    if ($line -match '(?i)\bRx\s*CRC\s*Errors?\b')    { $n=Get-LastNumber $line; if($n -ne $null){ $obj.RxCRC  =[double]$n } }
    if ($line -match '(?i)\bRX\s*PLCP\s*Errors?\b')   { $n=Get-LastNumber $line; if($n -ne $null){ $obj.RxPLCP =[double]$n } }
    if ($line -match '(?i)\bChannel\s*Changes?\b')    { $n=Get-LastNumber $line; if($n -ne $null){ $obj.ChannelChanges=[double]$n } }
    if ($line -match '(?i)\bTX\s*Power\s*Changes?\b') { $n=Get-LastNumber $line; if($n -ne $null){ $obj.TxPowerChanges=[double]$n } }

    if ($line -match '(?i)\bChannel\s*busy\b') {
      $b1=$null; $b4=$null; $b64=$null
      if (TryExtractPercentTriplet -Line $line -Busy1s ([ref]$b1) -Busy4s ([ref]$b4) -Busy64s ([ref]$b64)) {
        if ($b1 -ne $null){ $obj.Busy1s=[double]$b1 }
        if ($b4 -ne $null){ $obj.Busy4s=[double]$b4 }
        if ($b64 -ne $null){ $obj.Busy64s=[double]$b64 }
      } else {
        $m1=[regex]::Match($line,'(?i)\b1s\b[^0-9\-]*(-?\d+(?:\.\d+)?)(?:\s*%)?'); if($m1.Success){ $obj.Busy1s=[double]$m1.Groups[1].Value }
        $m4=[regex]::Match($line,'(?i)\b4s\b[^0-9\-]*(-?\d+(?:\.\d+)?)(?:\s*%)?'); if($m4.Success){ $obj.Busy4s=[double]$m4.Groups[1].Value }
        $m64=[regex]::Match($line,'(?i)\b64s\b[^0-9\-]*(-?\d+(?:\.\d+)?)(?:\s*%)?'); if($m64.Success){ $obj.Busy64s=[double]$m64.Groups[1].Value }
      }
    }
    if ($line -match '(?i)\bBusy\s*\(?\s*Beacon\s*\)?') { $n=Get-LastNumber $line; if($n -ne $null){ $obj.BusyBeacon=[double]$n } }
    if ($line -match '(?i)\bTx\s*Beacon\b')             { $n=Get-LastNumber $line; if($n -ne $null){ $obj.TxBeacon=[double]$n } }
    if ($line -match '(?i)\bRx\s*Beacon\b')             { $n=Get-LastNumber $line; if($n -ne $null){ $obj.RxBeacon=[double]$n } }

    if ($line -match '(?i)\bCCA\b') {
      $mOur=[regex]::Match($line,'(?i)\b(our|self|own)\b[^0-9\-]*(-?\d+(?:\.\d+)?)')
      $mOth=[regex]::Match($line,'(?i)\bother\b[^0-9\-]*(-?\d+(?:\.\d+)?)')
      $mInt=[regex]::Match($line,'(?i)\b(interf|non[-\s]?wi[-\s]?fi)\b[^0-9\-]*(-?\d+(?:\.\d+)?)')
      if ($mOur.Success) { $obj.CCA_Our = [double]$mOur.Groups[$mOur.Groups.Count-1].Value }
      if ($mOth.Success) { $obj.CCA_Other = [double]$mOth.Groups[$mOth.Groups.Count-1].Value }
      if ($mInt.Success) { $obj.CCA_Interference = [double]$mInt.Groups[$mInt.Groups.Count-1].Value }
      if (-not $mOur.Success -and -not $mOth.Success -and -not $mInt.Success) {
        $nums=[regex]::Matches($line,'(-?\d+(?:\.\d+)?)')
        if ($nums.Count -ge 3) {
          $obj.CCA_Our=[double]$nums[0].Value; $obj.CCA_Other=[double]$nums[1].Value; $obj.CCA_Interference=[double]$nums[2].Value
        }
      }
    }
  }

  if (-not $seenAnyRadio) {
    $guess='0'
    [void](Ensure-RadioSlot -Data $data -ApName $apName -Radio $guess)
    Write-Log ("Parse-RadioStatsFile: no radio markers; created empty slot radio={0}" -f $guess)
  }

  Write-Log ("Parse-RadioStatsFile: Path='{0}' OutputTimeUTC={1}" -f $Path, (ToLogStr-DT $otUtc "utc"))
  return New-Object psobject -Property @{ Path=$Path; AP=$apName; OutputTime=$otUtc; Data=$data; IsRadio=$true }
}

# ===== -SnapshotFiles 展開（非再帰） =====
function Expand-SnapshotInputs { param([string[]]$Inputs)
  $files=New-Object System.Collections.Generic.List[string]
  if ($Inputs -eq $null) { return $files }
  foreach ($p in $Inputs) {
    if ([string]::IsNullOrWhiteSpace($p)) { continue }
    if (Test-Path -LiteralPath $p) {
      $it=$null; try{ $it=Get-Item -LiteralPath $p -ErrorAction Stop }catch{ $it=$null }
      if ($it -ne $null) {
        if ($it.PSIsContainer) {
          $items=@(); try{ $items=Get-ChildItem -LiteralPath $it.FullName -File -ErrorAction Stop }catch{ $items=@() }
          foreach ($f in $items) { $files.Add($f.FullName) }
        } else { $files.Add($it.FullName) }
      }
    } else {
      $parent=$null; $leaf=$null
      try{ $parent=Split-Path -Path $p -Parent }catch{}
      try{ $leaf  =Split-Path -Path $p -Leaf }catch{ $leaf=$p }
      if ([string]::IsNullOrWhiteSpace($parent)) { $parent="." }
      $pattern=$null; try{ $pattern=Join-Path -Path $parent -ChildPath $leaf }catch{ $pattern=$p }
      $items=@(); try{ $items=Get-ChildItem -Path $pattern -File -ErrorAction Stop }catch{ $items=@() }
      foreach ($f in $items) { $files.Add($f.FullName) }
    }
  }
  $set=New-Object System.Collections.Generic.HashSet[string]; $out=New-Object System.Collections.Generic.List[string]
  foreach ($f in $files) { if (-not $set.Contains($f)) { [void]$set.Add($f); $out.Add($f) } }
  return $out
}

# ===== 期間補正 =====
function Fix-Duration {
  param([Nullable[datetime]]$bt,[Nullable[datetime]]$at,[string]$bPath,[string]$aPath)
  $sec=0
  if ($bt -ne $null -and $at -ne $null) {
    try { $sec=[int][Math]::Abs(($at - $bt).TotalSeconds) } catch { $sec=0 }
  }
  if ($sec -lt 30) {
    try {
      $t1=[System.IO.File]::GetLastWriteTime($bPath)
      $t2=[System.IO.File]::GetLastWriteTime($aPath)
      $sec2=[int]([Math]::Abs(($t2 - $t1).TotalSeconds))
      if ($sec2 -ge 30) { Write-Log ("Duration fix: using filetime {0}s (b={1} a={2})" -f $sec2,$t1,$t2); return $sec2 }
    } catch {}
    $sec=900
    Write-Log ("Duration fix: too small; fallback to default 900s")
  }
  return $sec
}

# ===== 安全なレート算出 =====
function SafeRate {
  param([Nullable[double]]$After,[Nullable[double]]$Before,[int]$Dur,[string]$Label,[string]$AP,[string]$Radio)
  if ($After -eq $null -or $Before -eq $null) { return $null }
  if ($Dur -le 0) { return $null }
  $diff = $After - $Before
  if ($diff -lt 0) { $diff = 0.0 }
  $rate = $diff / [double]$Dur
  Write-Log ("RateCalc: {0} AP='{1}' R={2} before={3} after={4} dur={5}s rate={6}" -f $Label,$AP,$Radio,$Before,$After,$Dur,$rate)
  if ($rate -gt 100000) {
    Write-Log ("RateCalc WARN: absurd {0} -> drop (>{1}/s)" -f $Label,100000)
    return $null
  }
  return [Math]::Round($rate,6)
}

# ===== 集計/出力 =====
Init-Log -Preferred $LogFile -OutputHtmlPath $OutputHtml -OutputCsvPath $OutputCsv
Write-Log "=== START ==="
Write-Log ("Args: Before='{0}' After='{1}' Snapshots='{2}' Html='{3}' Csv='{4}'" -f $BeforeFile,$AfterFile,($SnapshotFiles -join ';'),$OutputHtml,$OutputCsv)

$segments=@()
$allInputFiles=@()
$radioParsed=@()

if ($SnapshotFiles -and $SnapshotFiles.Count -ge 1) {
  $fileList = Expand-SnapshotInputs -Inputs $SnapshotFiles
  Write-Log ("Expand-Snapshot-Inputs: {0} files" -f $fileList.Count)
  if ($fileList.Count -lt 2) { throw "SnapshotFiles: 指定から2つ以上のファイルが見つかりません。フォルダ直下に2つ以上置くか、複数ファイル/ワイルドカードを指定してください。" }
  $allInputFiles = $fileList

  foreach ($f in $fileList) {
    $lines=@(); try{ $lines=Get-Content -LiteralPath $f -Encoding UTF8 -TotalCount 60 }catch{ $lines=@() }
    if (Is-RadioStatsLines -Lines $lines) {
      Write-Log ("Pick radio-stats: {0}" -f $f)
      $radioParsed += (Parse-RadioStatsFile -Path $f)
    } else {
      Write-Log ("Skip(non radio-stats): {0}" -f $f)
    }
  }
  if ($radioParsed.Count -lt 2) { throw "radio-stats の候補が2つ未満でした。arm history / aps / bss-table / radio-info はスナップショット対象外です。radio-stats の出力ファイルを2つ以上配置してください。" }

  $sorted = $radioParsed | Sort-Object { if ($_.OutputTime -ne $null) { $_.OutputTime } else { [System.IO.File]::GetLastWriteTime($_.Path) } }

  for ($i=0; $i -lt $sorted.Count-1; $i++) {
    $b=$sorted[$i]; $a=$sorted[$i+1]
    $sec = Fix-Duration -bt $b.OutputTime -at $a.OutputTime -bPath $b.Path -aPath $a.Path
    $sj = Convert-ToJst -dt $b.OutputTime -Context ("seg{0}-startUtc" -f $i)
    $ej = Convert-ToJst -dt $a.OutputTime -Context ("seg{0}-endUtc"   -f $i)
    $segments += (New-Object psobject -Property @{ Before=$b; After=$a; DurationSec=$sec; StartJst=$sj; EndJst=$ej })
    Write-Log ("Segment[{0}] durSec={1} StartJST={2} EndJST={3}" -f $i,$sec,$sj,$ej)
  }
}
elseif (-not [string]::IsNullOrWhiteSpace($BeforeFile) -and -not [string]::IsNullOrWhiteSpace($AfterFile)) {
  $b = Parse-RadioStatsFile -Path $BeforeFile
  $a = Parse-RadioStatsFile -Path $AfterFile
  if (-not $b.IsRadio -or -not $a.IsRadio) { throw "指定ファイルが radio-stats 形式ではありません。radio-stats 出力を指定してください。" }

  if ($DurationSec -le 0) {
    $DurationSec = Fix-Duration -bt $b.OutputTime -at $a.OutputTime -bPath $BeforeFile -aPath $AfterFile
  }
  $sj = Convert-ToJst -dt $b.OutputTime -Context "single-startUtc"
  $ej = Convert-ToJst -dt $a.OutputTime -Context "single-endUtc"
  $segments += (New-Object psobject -Property @{ Before=$b; After=$a; DurationSec=$DurationSec; StartJst=$sj; EndJst=$ej })
  Write-Log ("Single durSec={0} StartJST={1} EndJST={2}" -f $DurationSec,$sj,$ej)
}
else {
  throw "単区間比較は -BeforeFile/-AfterFile、時系列集計は -SnapshotFiles に（フォルダ/ワイルドカード/ファイルのいずれかを）1個以上指定してください（展開後2つ以上の radio-stats ファイルが必要）。"
}

# ===== 補完用バックアップ =====
$dirs = New-Object System.Collections.Generic.HashSet[string]
foreach ($seg in $segments) {
  try { $d1 = Split-Path -Path $seg.Before.Path -Parent } catch { $d1 = $null }
  try { $d2 = Split-Path -Path $seg.After.Path  -Parent } catch { $d2 = $null }
  if (-not [string]::IsNullOrWhiteSpace($d1)) { [void]$dirs.Add($d1) }
  if (-not [string]::IsNullOrWhiteSpace($d2)) { [void]$dirs.Add($d2) }
}
$dirList=@(); foreach ($d in $dirs) { $dirList += $d }
Write-Log ("Backup scan dirs: " + ($dirList -join ';'))
$BackupMap = Build-Backup-From-Dirs -Dirs $dirList
Write-Log ("Backup keys: " + ([string]::Join(',', $BackupMap.Keys)))

# ===== CSV 出力先 =====
if ([string]::IsNullOrWhiteSpace($OutputCsv)) {
  $baseRef = $segments[$segments.Count-1].After.Path
  $outDir = Get-ParentOrCwd -PathLike $baseRef
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $OutputCsv = Join-Path -Path $outDir -ChildPath ("aruba_radio_stats_diff_{0}.csv" -f $ts)
}
Write-Log ("OutputCsv: " + $OutputCsv)

# ===== 列定義 =====
$colDefs = @(
  @('AP','AP 名（親フォルダ名）'),
  @('Radio','ラジオ番号（0/1 等）'),
  @('Channel','運用チャネル番号（本文/補完）'),
  @('Band','2.4GHz/5GHz/6GHz（推定/補完）'),
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

# ===== CSV ヘッダ =====
$csvHeader = ($colDefs | ForEach-Object { $_[0] }) -join ','
Set-Content -LiteralPath $OutputCsv -Value $csvHeader -Encoding UTF8

# ===== 診断（簡易スコア） =====
function Make-Diagnosis {
  param(
    [double]$Busy64,[double]$BusyB,[double]$TxB,[double]$RxB,
    [double]$CCAO,[double]$CCAI,[double]$Retry,[double]$CRC,[double]$PLCP,
    [double]$ChgPH,[double]$TxPPH,[Nullable[int]]$Channel
  )
  $deltaCCA=$null
  if ($BusyB -ne $null -and $TxB -ne $null -and $RxB -ne $null) { $deltaCCA = $BusyB - ($TxB + $RxB) }

  $scoreCoch=0;$pCoch=0.0
  if ($deltaCCA -ne $null -and $deltaCCA -ge $TH_ExcessCCA){$scoreCoch+=1;$pCoch+=$deltaCCA}
  if ($CCAO -ne $null -and $CCAO -ge $TH_OtherHigh){$scoreCoch+=1;$pCoch+=$CCAO}
  if ($Busy64 -ne $null -and $Busy64 -ge $TH_BusyWarn){$scoreCoch+=1;$pCoch+=($Busy64-$TH_BusyWarn)}

  $scoreInter=0;$pInter=0.0
  if ($CCAI -ne $null -and $CCAI -ge $TH_InterfHigh){$scoreInter+=1;$pInter+=$CCAI}
  if ($PLCP -ne $null -and $PLCP -ge $TH_PLCPHigh){$scoreInter+=1;$pInter+=[Math]::Min($PLCP,$TH_PLCPHigh*4)/10.0}

  $scoreQual=0;$pQual=0.0
  if ($Busy64 -le $TH_BusyWarn){
    if ($Retry -ge $TH_RetryHigh){$scoreQual+=1;$pQual+=$Retry}
    if ($CRC   -ge $TH_CRCHigh ){$scoreQual+=1;$pQual+=($CRC/10.0)}
  }

  $scoreBusy=0;$pBusy=0.0
  if ($Busy64 -ge $TH_BusyHigh){$scoreBusy=1;$pBusy=$Busy64}

  $labels=@('co','inter','qual','busy')
  $scores=@($scoreCoch,$scoreInter,$scoreQual,$scoreBusy)
  $power =@($pCoch,$pInter,$pQual,$pBusy)
  $maxIdx=0; for($i=1;$i -lt $scores.Length;$i++){ if($scores[$i] -gt $scores[$maxIdx] -or ($scores[$i] -eq $scores[$maxIdx] -and $power[$i] -gt $power[$maxIdx])){ $maxIdx=$i } }
  $root=$labels[$maxIdx]

  $simple='';$why='';$tips='';$sev='sev-ok'
  $secList=New-Object System.Collections.Generic.List[string]

  if($root -eq 'co'){
    $simple="原因：近くの同一チャネルのWi-Fiが強く、電波の取り合いが発生しています。"
    $whyParts=@(); if($CCAO -ne $null){$whyParts+=("他BSS "+[int]$CCAO+"%")}
    if($deltaCCA -ne $null){$whyParts+=("占有差ΔCCA "+[int]$deltaCCA+"pt")}
    if($Busy64 -ne $null){$whyParts+=("Busy64 "+[int]$Busy64+"%")}
    $why="根拠："+([string]::Join(" / ",$whyParts))
    $tips="対策：チャネル再配置・再利用距離の拡大／帯域幅20MHz化／最低基本レートの引き上げ。"
    $sev='sev-warn'
  } elseif($root -eq 'inter'){
    $simple="原因：Wi-Fi以外の電波ノイズの影響が大きい可能性があります。"
    $whyParts=@(); if($CCAI -ne $null){$whyParts+=("Interference "+[int]$CCAI+"%")}
    if($PLCP -ne $null){$whyParts+=("PLCP "+[int]$PLCP+"/s")}
    $why="根拠："+([string]::Join(" / ",$whyParts))
    $tips="対策：別チャネル（非DFS含む）に一時固定して観測／周辺装置の稼働時間と突き合わせ。"
    $sev='sev-warn'
  } elseif($root -eq 'qual'){
    $simple="原因：端末の電波が弱い・遮蔽・隠れ端末の可能性が高いです。"
    $why="根拠：Busyが低め／Retry "+[int]$Retry+"/s／CRC "+[int]$CRC+"/s"
    $tips="対策：AP配置・出力レンジの適正化／ローミング閾値の見直し／低速端末の抑制。"
    $sev='sev-warn'
  } else {
    $simple="状況：電波の混雑が高く、全体的にエアタイムが逼迫しています。"
    $why="根拠：Busy64 "+[int]$Busy64+"%"
    $tips="対策：ユーザー密度の分散／帯域幅20MHz化／不要な低速レートの無効化。"
    $sev='sev-warn'
  }

  $band = Get-BandFromChannel -Channel $Channel
  $laaHit=$false
  if ($band -eq '5GHz') {
    if ($CCAI -ge $TH_LAA_MinInterf) {
      $okOther = $true
      if ($CCAO -ne $null) { if ($CCAO -ge $TH_LAA_MaxOther) { $okOther=$false } }
      if ($okOther -and $Busy64 -ge $TH_LAA_MinBusy) { $laaHit=$true }
    }
  }
  if ($laaHit) { $secList.Add((Sanitize-Text ("候補：LAA/NR-Uの可能性（5GHz ch"+$Channel+"、Interf "+([int]$CCAI)+"%、Other "+([int]$CCAO)+"%、Busy64 "+([int]$Busy64)+"%）"))) }

  $oneLine=Sanitize-Text ($simple+" "+$why)
  $tips   =Sanitize-Text $tips
  $mainScore = $scores[$maxIdx] + ($power[$maxIdx] / 100.0)

  return New-Object psobject -Property @{ Simple=$oneLine; Tips=$tips; Secondary=$secList; Severity=$sev; Score=$mainScore; Band=$band }
}

# ===== HTML 生成 =====
$rows=@()
$cards=@()
$hourBuckets=@{}

function Get-FileTimeJst {
  param([string]$p,[string]$ctx)
  try {
    $t=[System.IO.File]::GetLastWriteTime($p)
    Write-Log ("FileTime: {0}={1}" -f $ctx,$t)
    return (Convert-ToJst -dt $t -Context ("{0}-filetime" -f $ctx))
  } catch { return $null }
}

foreach ($seg in $segments) {
  $before=$seg.Before.Data; $after=$seg.After.Data; $dur=$seg.DurationSec

  $startDisp=$seg.StartJst
  if ($startDisp -eq $null -and $seg.Before -ne $null -and -not [string]::IsNullOrWhiteSpace($seg.Before.Path)) { $startDisp=Get-FileTimeJst -p $seg.Before.Path -ctx "start" }
  $endDisp=$seg.EndJst
  if ($endDisp -eq $null -and $seg.After -ne $null -and -not [string]::IsNullOrWhiteSpace($seg.After.Path)) { $endDisp=Get-FileTimeJst -p $seg.After.Path -ctx "end" }

  $keys=New-Object System.Collections.Generic.HashSet[string]
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

    $retry_ps = SafeRate -After $a.RxRetry -Before $b.RxRetry -Dur $dur -Label "RxRetry" -AP $ap -Radio $radio
    $crc_ps   = SafeRate -After $a.RxCRC   -Before $b.RxCRC   -Dur $dur -Label "RxCRC"   -AP $ap -Radio $radio
    $plcp_ps  = SafeRate -After $a.RxPLCP  -Before $b.RxPLCP  -Dur $dur -Label "RxPLCP"  -AP $ap -Radio $radio
    $chg_s    = SafeRate -After $a.ChannelChanges -Before $b.ChannelChanges -Dur $dur -Label "ChannelChanges" -AP $ap -Radio $radio
    $txp_s    = SafeRate -After $a.TxPowerChanges -Before $b.TxPowerChanges -Dur $dur -Label "TxPowerChanges" -AP $ap -Radio $radio
    $chg_ph=$null; if ($chg_s -ne $null){ $chg_ph=[Math]::Round($chg_s*3600,6) }
    $txp_ph=$null; if ($txp_s -ne $null){ $txp_ph=[Math]::Round($txp_s*3600,6) }

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

    if ([string]::IsNullOrWhiteSpace($band) -and $chan -ne $null) { $band=Get-BandFromChannel $chan }

    $hasAny=$false
    foreach($vv in @($retry_ps,$crc_ps,$plcp_ps,$chg_ph,$txp_ph,$busy1s,$busy4s,$busy64,$busyB,$txB,$rxB,$ccaO,$ccaOt,$ccaI,$chan)){
      if($vv -ne $null -and $vv -ne ''){ $hasAny=$true }
    }
    if(-not $hasAny){ continue }

    $d = Make-Diagnosis -Busy64 $busy64 -BusyB $busyB -TxB $txB -RxB $rxB `
                        -CCAO $ccaOt -CCAI $ccaI -Retry $retry_ps -CRC $crc_ps -PLCP $plcp_ps `
                        -ChgPH $chg_ph -TxPPH $txp_ph -Channel $chan

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
    $vals=@(); foreach ($def in $colDefs) { $name=$def[0]; $v=$rowObj.PSObject.Properties[$name].Value; if ($v -eq $null) { $vals+='' } else { $vals+=$v } }
    $escaped=@(); foreach($v in $vals){ if($v -match '[,"]'){ $escaped+=('"{0}"' -f ($v -replace '"','""')) } else { $escaped+=$v } }
    Add-Content -LiteralPath $OutputCsv -Value ($escaped -join ',') -Encoding UTF8

    $rows += $rowObj
    $cards += (New-Object psobject -Property @{
      AP=$ap; Radio=$radio; Channel=$chan; Band=$band;
      EndJST=$(if ($endDisp -ne $null){ $endDisp.ToString('yyyy-MM-dd HH:mm') + " JST" } else { '' });
      Simple=$d.Simple; Tips=$d.Tips; Secondary=$d.Secondary; Severity=$d.Severity; Score=$d.Score
    })

    if ($endDisp -ne $null) {
      $key=$endDisp.ToString('HH')
      if (-not $hourBuckets.ContainsKey($key)) { $hourBuckets[$key] = New-Object psobject -Property @{ N=0; Busy64=0.0; CCAO=0.0; CCAI=0.0; PLCP=0.0; CRC=0.0; Retry=0.0 } }
      $h=$hourBuckets[$key]; $h.N++
      if ($busy64 -ne $null){ $h.Busy64 += $busy64 }
      if ($ccaOt -ne $null){ $h.CCAO   += $ccaOt }
      if ($ccaI  -ne $null){ $h.CCAI   += $ccaI }
      if ($plcp_ps -ne $null){ $h.PLCP += $plcp_ps }
      if ($crc_ps  -ne $null){ $h.CRC  += $crc_ps }
      if ($retry_ps -ne $null){ $h.Retry += $retry_ps }
    }
  }
}

Write-Output ("CSV : {0}" -f $OutputCsv)

# ===== HTML =====
if (-not [string]::IsNullOrWhiteSpace($OutputHtml)) {
  $baseRef=$OutputHtml
  $outDir=Get-ParentOrCwd -PathLike $baseRef
  if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

  $firstSeg=$null; $lastSeg=$null
  if ($segments.Count -gt 0) { $firstSeg=$segments[0]; $lastSeg=$segments[$segments.Count-1] }
  $btStr=''; $atStr=''
  if ($firstSeg -ne $null -and $firstSeg.StartJst -ne $null) { $btStr=$firstSeg.StartJst.ToString('yyyy-MM-dd HH:mm:ss') + " JST" }
  if ($lastSeg  -ne $null -and $lastSeg.EndJst   -ne $null)   { $atStr=$lastSeg.EndJst.ToString('yyyy-MM-dd HH:mm:ss') + " JST" }

  $titleText=$Title
  if ([string]::IsNullOrWhiteSpace($titleText)) {
    if (-not [string]::IsNullOrWhiteSpace($btStr) -or -not [string]::IsNullOrWhiteSpace($atStr)) {
      $titleText = "Aruba Radio Stats Diff（JST） " + $btStr + " → " + $atStr
    } else { $titleText = "Aruba Radio Stats Diff（JST）" }
  }

  $cols=$colDefs
  $topCards = $cards | Sort-Object -Property @{Expression='Score';Descending=$true} | Select-Object -First 10

  $sb=New-Object System.Text.StringBuilder
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

  if ($topCards -and $topCards.Count -gt 0) {
    [void]$sb.AppendLine('<div id="cards">')
    foreach ($c in $topCards) {
      $cls=$c.Severity
      [void]$sb.AppendLine('<div class="diag-card '+$cls+'">')
      $titleLine = $c.AP+" / Radio "+$c.Radio
      if ($c.Channel -ne $null -and $c.Channel -ne '') { $titleLine = $titleLine + " ch" + $c.Channel }
      if (-not [string]::IsNullOrWhiteSpace($c.Band)) { $titleLine = $titleLine + " ("+$c.Band+")" }
      $titleLine = $titleLine + " @ " + $c.EndJST
      [void]$sb.AppendLine('<div class="title">'+ (HtmlEscape $titleLine) +'</div>')
      [void]$sb.AppendLine('<div class="sub">'+ (HtmlEscape $c.Simple) +'</div>')
      if ($c.Secondary -ne $null -and $c.Secondary.Count -gt 0) { foreach ($s in $c.Secondary) { [void]$sb.AppendLine('<div class="sec">'+ (HtmlEscape $s) +'</div>') } }
      [void]$sb.AppendLine('<div class="sec">'+ (HtmlEscape ("対策："+$c.Tips)) +'</div>')
      [void]$sb.AppendLine('</div>')
    }
    [void]$sb.AppendLine('</div>')
  }

  [void]$sb.AppendLine('<div style="margin:10px 0"><input id="flt" type="search" placeholder="フィルタ（AP/数値/文言）..." oninput="filterTable()"></div>')

  [void]$sb.AppendLine('<details class="help"><summary>列の見方（クリックで開閉）</summary><div><ul>')
  foreach ($pair in $colDefs) { [void]$sb.AppendLine('<li><b>'+ (HtmlEscape $pair[0]) +'</b>：'+ (HtmlEscape $pair[1]) +'</li>') }
  [void]$sb.AppendLine('</ul></div></details>')

  [void]$sb.AppendLine('<table id="tbl"><thead><tr>')
  foreach ($pair in $colDefs) { $name=$pair[0]; $desc=$pair[1]; [void]$sb.AppendLine('<th title="'+ (HtmlEscape $desc) +'">'+ (HtmlEscape $name) +'</th>') }
  [void]$sb.AppendLine('</tr></thead><tbody>')
  foreach ($r in $rows) {
    [void]$sb.AppendLine('<tr>')
    foreach ($pair in $colDefs) {
      $c=$pair[0]; $v=$r.PSObject.Properties[$c].Value
      $text=''; if ($null -ne $v) { if ($v -is [double] -or $v -is [single]) { $text=([string]([Math]::Round([double]$v,6))) } else { $text=[string]$v } }
      if ($c -eq 'StartJST' -or $c -eq 'EndJST') { if ($text -ne '') { $text = $text + ' JST' } }
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
    (function(idx){ ths[idx].addEventListener("click", function(){ if(lastCol===idx){ asc=!asc; } else { lastCol=idx; asc=true; } sortBy(idx,asc); }); })(i);
  }
  function getVal(td){ var t=td.textContent; var n=parseFloat(t); if(!isNaN(n)) return {n:n,s:t.toLowerCase()}; return {n:null,s:t.toLowerCase()}; }
  function cmp(a,b,ascFlag){ if(a.n!==null&&b.n!==null){ if(a.n<b.n) return ascFlag?-1:1; if(a.n>b.n) return ascFlag?1:-1; return 0; } if(a.s<b.s) return ascFlag?-1:1; if(a.s>b.s) return ascFlag?1:-1; return 0; }
  function sortBy(col,ascFlag){ var tbody=tbl.tBodies[0]; var rows=[].slice.call(tbody.rows); rows.sort(function(r1,r2){ var a=getVal(r1.cells[col]); var b=getVal(r2.cells[col]); return cmp(a,b,ascFlag); }); for(var i=0;i<rows.length;i++){ tbody.appendChild(rows[i]); } }
  window.filterTable=function(){ var q=document.getElementById("flt").value.toLowerCase(); var trs=tbl.tBodies[0].rows; for(var i=0;i<trs.length;i++){ var show=false, tds=trs[i].cells; for(var j=0;j<tds.length;j++){ var t=tds[j].textContent.toLowerCase(); if(t.indexOf(q)>=0){ show=true; break; } } trs[i].style.display = show? "":"none"; } };
})();
</script>')
  $html=$sb.ToString()
  $nameOnly=$null; try{ $nameOnly=Split-Path -Path $OutputHtml -Leaf }catch{}
  if ([string]::IsNullOrWhiteSpace($nameOnly)) { $nameOnly="radio_stats_diff.html" }
  $htmlPath=Join-Path -Path $outDir -ChildPath $nameOnly
  Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8
  Write-Output ("HTML: {0}" -f $htmlPath)
  Write-Log ("HTML written: {0}" -f $htmlPath)
}
Write-Log "=== END ==="
exit 0
