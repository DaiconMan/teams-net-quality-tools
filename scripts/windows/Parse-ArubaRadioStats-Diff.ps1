<# 
.SYNOPSIS
  Aruba "show ap debug radio-stats" スナップショットの差分/秒を算出し、CSV/HTML を生成（JST表示）。
  - 同フォルダの "show ap bss-table" / "show aps" / "show ap debug radio-info" / "show ap arm history"
    テキストから AP名・Channel・Band を補完（バックアップ参照）
  - -SnapshotFiles はフォルダ/ワイルドカード/ファイル混在OK（展開後2ファイル以上）
  - radio-stats 以外（arm history 等）はスナップショットとしては除外（補完専用）
  - LAA/NR-U 候補検出、時間帯コメント
  - PowerShell 5.1対応・三項演算子不使用・Host未使用・日本語/スペース/OneDrive対応
#>

[CmdletBinding()]
param(
  [string]$BeforeFile,
  [string]$AfterFile,
  [int]$DurationSec,
  [string[]]$SnapshotFiles,
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

# ===== JST 変換 =====
function Convert-ToJst {
  param([Nullable[DateTime]]$dt)
  if ($dt -eq $null) { return $null }
  $tz = $null
  try { $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById("Tokyo Standard Time") } catch { $tz = $null }
  try {
    $v = $dt.Value
    if     ($v.Kind -eq [System.DateTimeKind]::Unspecified) { $v = [DateTime]::SpecifyKind($v, [System.DateTimeKind]::Utc) }
    elseif ($v.Kind -eq [System.DateTimeKind]::Local)       { $v = [System.TimeZoneInfo]::ConvertTimeToUtc($v) }
    if ($tz -ne $null) { return [System.TimeZoneInfo]::ConvertTimeFromUtc($v, $tz) } else { return $v.AddHours(9) }
  } catch { try { return $dt.Value.AddHours(9) } catch { return $dt } }
}

# ===== 補助 =====
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

# ===== 数値/時刻抽出 =====
function Get-LastNumber { param([string]$Line)
  if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
  $m=[regex]::Matches($Line,'(-?\d+(?:\.\d+)?)'); if($m.Count -gt 0){ return [double]$m[$m.Count-1].Value } return $null
}

# Busy 1s/4s/64s の抽出を強化
function TryExtractPercentTriplet {
  param(
    [string]$Line,
    [ref]$Busy1s,
    [ref]$Busy4s,
    [ref]$Busy64s
  )
  $ok=$false
  # 代表例: "Channel busy 1s: 64% 4s: 42% 64s: 42%"
  $m1=[regex]::Match($Line,'\b1s\b[^0-9\-]*(-?\d+(?:\.\d+)?)(?:\s*%)?'); if($m1.Success){ $Busy1s.Value=[double]$m1.Groups[1].Value; $ok=$true }
  $m4=[regex]::Match($Line,'\b4s\b[^0-9\-]*(-?\d+(?:\.\d+)?)(?:\s*%)?'); if($m4.Success){ $Busy4s.Value=[double]$m4.Groups[1].Value; $ok=$true }
  $m64=[regex]::Match($Line,'\b64s\b[^0-9\-]*(-?\d+(?:\.\d+)?)(?:\s*%)?'); if($m64.Success){ $Busy64s.Value=[double]$m64.Groups[1].Value; $ok=$true }
  if ($ok) { return $true }

  # 代表例: "1s/4s/64s: 64/42/42" or "1s:64 4s:42 64s:42"
  $m=[regex]::Match($Line,'1s[^0-9]{0,5}(\d+(?:\.\d+)?).{0,12}4s[^0-9]{0,5}(\d+(?:\.\d+)?).{0,12}64s[^0-9]{0,5}(\d+(?:\.\d+)?)')
  if ($m.Success) {
    $Busy1s.Value=[double]$m.Groups[1].Value
    $Busy4s.Value=[double]$m.Groups[2].Value
    $Busy64s.Value=[double]$m.Groups[3].Value
    return $true
  }
  return $false
}

# Output Time を UTC として取り出す
function Extract-OutputTime {
  param([string[]]$Lines)
  if ($Lines -eq $null -or $Lines.Count -eq 0) { return $null }
  foreach ($raw in $Lines) {
    $line = ($raw -replace '\r','').Trim()
    if ($line -match '(?i)Output\s*Time\s*[:=]\s*(.+)$') {
      $rhs = $Matches[1].Trim()

      # 1) "2025-08-22 03:38:51 UTC"（コロン直後に日時が来てもOK）
      $mUtc=[regex]::Match($rhs,'^(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})\s*(UTC|GMT)\b',[System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
      if($mUtc.Success){
        try{
          $dt=[DateTime]::Parse($mUtc.Groups[1].Value,[System.Globalization.CultureInfo]::InvariantCulture)
          return [DateTime]::SpecifyKind($dt,[System.DateTimeKind]::Utc)
        }catch{}
      }

      # 2) epoch 秒 / ミリ秒
      if ($rhs -match '^\d{10}(\.\d+)?$') { try { $sec=[long]([double]$rhs); return ([System.DateTimeOffset]::FromUnixTimeSeconds($sec)).UtcDateTime } catch {} }
      if ($rhs -match '^\d{13}$')        { try { $ms=[long]$rhs; return ([System.DateTimeOffset]::FromUnixTimeMilliseconds($ms)).UtcDateTime } catch {} }

      # 3) ISO8601 with offset/Z
      $dto=[System.DateTimeOffset]::MinValue
      if ([System.DateTimeOffset]::TryParse($rhs,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::AssumeUniversal,[ref]$dto)) { return $dto.UtcDateTime }

      # 4) 最後の手段：UTC と仮定
      try { $dt2=[DateTime]::Parse($rhs,[System.Globalization.CultureInfo]::InvariantCulture); return [DateTime]::SpecifyKind($dt2,[System.DateTimeKind]::Utc) } catch {}
    }
  }
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

# ===== AP名推定／補正 =====
function Guess-APName { param([string[]]$Lines,[string]$Path)
  foreach ($raw in $Lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $m1=[regex]::Match($line,'(?i)\bAP\s*Name\s*[:=]\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)'); if($m1.Success){ return $m1.Groups[1].Value }
    $m2=[regex]::Match($line,'^\s*([A-Za-z0-9][A-Za-z0-9_\-\.:\(\)\/\\]+)\s*[#>]\s*'); if($m2.Success){ return $m2.Groups[1].Value }
    $m3=[regex]::Match($line,'(?i)\bAP\s*[:=]\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)'); if($m3.Success){ return $m3.Groups[1].Value }
  }
  $leaf=$null; try{ $leaf=[System.IO.Path]::GetFileNameWithoutExtension($Path) }catch{}
  if (-not [string]::IsNullOrWhiteSpace($leaf)) {
    $cand = ($leaf -replace '(?i)show|ap|debug|radio|stats|stat|output|log|arm|history','' -replace '[_\-\.\s]+$','' -replace '^\s+','').Trim()
    if ($cand -match '^[A-Za-z0-9].+') { return $cand }
  }
  $p=$null; try{ $p=Split-Path -Path $Path -Parent }catch{}
  if (-not [string]::IsNullOrWhiteSpace($p)) {
    $dirLeaf=$null; try{ $dirLeaf=Split-Path -Path $p -Leaf }catch{}
    if (-not [string]::IsNullOrWhiteSpace($dirLeaf)) { return $dirLeaf }
  }
  return ''
}
function Fix-AP-From-Backup { param([string]$ap,[string]$path,[hashtable]$Backup)
  if ($Backup -eq $null -or $Backup.Keys.Count -eq 0) { return $ap }
  if (-not [string]::IsNullOrWhiteSpace($ap) -and $Backup.ContainsKey($ap)) { return $ap }
  $leaf=''; try{ $leaf=[System.IO.Path]::GetFileName($path) }catch{}
  $full=$path
  foreach ($k in $Backup.Keys) {
    if (-not [string]::IsNullOrWhiteSpace($k)) {
      if ($leaf -and ($leaf -like "*$k*")) { return $k }
      if ($full -and ($full -like "*$k*")) { return $k }
    }
  }
  if ($Backup.Keys.Count -eq 1) { foreach($k in $Backup.Keys){ return $k } }
  return $ap
}

# ===== バンド判定 =====
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

# ===== バックアップ格納 =====
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
  if     ($band -eq '2.4GHz') { if ($slot.T24 -eq $null -or ($ts -ne $null -and $ts -gt $slot.T24)) { $slot.Ch24=[int]$ch; $slot.T24=$ts } }
  elseif ($band -eq '5GHz')   { if ($slot.T5  -eq $null -or ($ts -ne $null -and $ts -gt $slot.T5 )) { $slot.Ch5 =[int]$ch; $slot.T5 =$ts } }
  elseif ($band -eq '6GHz')   { if ($slot.T6  -eq $null -or ($ts -ne $null -and $ts -gt $slot.T6 )) { $slot.Ch6 =[int]$ch; $slot.T6 =$ts } }
}

# ===== バックアップ取り込み =====
function Parse-BssTable-File { param([string]$Path,[hashtable]$Backup)
  $ts=$null; try{ $ts=[System.IO.File]::GetLastWriteTime($Path) }catch{}
  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ return }
  foreach ($raw in $lines) {
    $line=($raw -replace '\r','').Trim(); if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $ap=''; $m1=[regex]::Match($line,'(?i)\bAP\s*Name\s*[:=]\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)'); if($m1.Success){ $ap=$m1.Groups[1].Value }
    if ([string]::IsNullOrWhiteSpace($ap)) { $m2=[regex]::Match($line,'(?i)\bAP-?Name\b[:=]?\s*([^\s,]+)'); if($m2.Success){ $ap=$m2.Groups[1].Value } }
    if ([string]::IsNullOrWhiteSpace($ap)) {
      if ($line -match '(?i)\bch(annel)?\b' -or $line -match '(?i)\bssid\b' -or $line -match '(?i)\bbssid\b') {
        $parts=$line -split '\s+'; if ($parts.Length -ge 2) { $ap=$parts[$parts.Length-1] }
      }
    }
    if ([string]::IsNullOrWhiteSpace($ap)) { continue }
    $ch=$null; $mCh=[regex]::Match($line,'(?i)\bCh(?:annel)?\s*[:=]?\s*(\d{1,3})\b'); if($mCh.Success){ try{$ch=[int]$mCh.Groups[1].Value}catch{} }
    if ($ch -eq $null) { $mCh2=[regex]::Match($line,'(?i)\b(\d{1,3})\b\s*(?:MHz|HT|EIRP)'); if($mCh2.Success){ try{$ch=[int]$mCh2.Groups[1].Value}catch{} } }
    $band = Detect-Band-Token -line $line -ch $ch
    if ([string]::IsNullOrWhiteSpace($band) -and $ch -ne $null) { $band = Get-BandFromChannel $ch }
    if ([string]::IsNullOrWhiteSpace($band) -or $ch -eq $null) { continue }
    Update-Backup -Backup $Backup -ap $ap -band $band -ch $ch -ts $ts
  }
}
function Parse-APS-File { param([string]$Path,[hashtable]$Backup)
  $ts=$null; try{ $ts=[System.IO.File]::GetLastWriteTime($Path) }catch{}
  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ return }
  $currentAP=''
  foreach ($raw in $lines) {
    $line=($raw -replace '\r','').Trim(); if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $mHead=[regex]::Match($line,'(?i)^(Name|AP\s*Name)\s*[:=]?\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)'); if($mHead.Success){ $currentAP=$mHead.Groups[2].Value }
    $mInline=[regex]::Match($line,'(?i)\bAP\s*Name\s*[:=]\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)'); if($mInline.Success){ $currentAP=$mInline.Groups[1].Value }
    if ([string]::IsNullOrWhiteSpace($currentAP)) { $parts=$line -split '\s+'; if ($parts.Length -ge 1 -and $parts[0] -match '^[A-Za-z0-9_\-\.:\(\)\/\\]+$'){ $currentAP=$parts[0] } }
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
  $ts=$null; try{ $ts=[System.IO.File]::GetLastWriteTime($Path) }catch{}
  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ return }
  $ap=''; foreach ($raw in $lines) {
    $line=($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $m1=[regex]::Match($line,'(?i)\bAP\s*Name\s*[:=]\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)'); if($m1.Success){ $ap=$m1.Groups[1].Value; break }
  }
  if ([string]::IsNullOrWhiteSpace($ap)) {
    $leaf=$null; try{ $leaf=[System.IO.Path]::GetFileNameWithoutExtension($Path) }catch{}
    if (-not [string]::IsNullOrWhiteSpace($leaf)) { $ap=$leaf }
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
  $ts=$null; try{ $ts=[System.IO.File]::GetLastWriteTime($Path) }catch{}
  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ return }
  $ap = Guess-APName -Lines $lines -Path $Path
  if ([string]::IsNullOrWhiteSpace($ap)) { $leaf=$null; try{ $leaf=[System.IO.Path]::GetFileNameWithoutExtension($Path) }catch{}; if (-not [string]::IsNullOrWhiteSpace($leaf)) { $ap=$leaf } }
  $curBand=''
  foreach ($raw in $lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $mPhy=[regex]::Match($line,'(?i)\bPhy[-\s]?Type\s*[:=]\s*([A-Za-z0-9\.]+)'); if($mPhy.Success){ $b = Map-PhyType-To-Band $mPhy.Groups[1].Value; if (-not [string]::IsNullOrWhiteSpace($b)) { $curBand = $b } }
    if ($line -match '(?i)\binterface\s*[:=]\s*([A-Za-z0-9_\-\.]+)') { $null = $Matches[1] } # ブロック切替のみ
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
      if     ($isBss)   { Parse-BssTable-File   -Path $f.FullName -Backup $bk }
      elseif ($isAPS)   { Parse-APS-File       -Path $f.FullName -Backup $bk }
      elseif ($isRInfo) { Parse-RadioInfo-File -Path $f.FullName -Backup $bk }
      elseif ($isARM)   { Parse-ARMHistory-File -Path $f.FullName -Backup $bk }
    }
  }
  return $bk
}

# ===== radio-stats パーサ =====
function Parse-RadioStatsFile {
  param([string]$Path)

  $lines=@(); try{ $lines=Get-Content -LiteralPath $Path -Encoding UTF8 }catch{ $lines=@() }
  if (-not (Is-RadioStatsLines -Lines $lines)) {
    return New-Object psobject -Property @{ Path=$Path; OutputTime=$null; Data=@{}; IsRadio=$false }
  }

  $otUtc = Extract-OutputTime -Lines $lines
  $apName = Guess-APName -Lines $lines -Path $Path

  $data=@{}; $currentRadio=''
  $ensure = {
    param([string]$r)
    $rk = ($apName + '|' + $r)
    if (-not $data.ContainsKey($rk)) {
      $data[$rk] = New-Object psobject -Property @{
        AP=$apName; Radio=$r; Channel=$null;
        RxRetry=$null; RxCRC=$null; RxPLCP=$null;
        ChannelChanges=$null; TxPowerChanges=$null;
        Busy1s=$null; Busy4s=$null; Busy64s=$null;
        BusyBeacon=$null; TxBeacon=$null; RxBeacon=$null;
        CCA_Our=$null; CCA_Other=$null; CCA_Interference=$null
      }
    } else {
      if ([string]::IsNullOrWhiteSpace($data[$rk].AP) -and -not [string]::IsNullOrWhiteSpace($apName)) { $data[$rk].AP=$apName }
    }
    return $data[$rk]
  }

  foreach ($raw in $lines) {
    $line=($raw -replace '\r','').Trim(); if ([string]::IsNullOrWhiteSpace($line)) { continue }
    if ($line -match '(?i)show\s+ap\s+debug\s+radio-stats\s+([012])\b') { $currentRadio=$Matches[1] }
    elseif ($line -match '(?i)\bradio\s*([012])\b') { $currentRadio=$Matches[1] }
    if ([string]::IsNullOrWhiteSpace($currentRadio)) {
      $leaf=$null; try{ $leaf=Split-Path -Path $Path -Leaf }catch{}
      if (-not [string]::IsNullOrWhiteSpace($leaf)) { $mfn=[regex]::Match($leaf,'(?i)\b([012])\b'); if($mfn.Success){ $currentRadio=$mfn.Groups[1].Value } }
    }
    if ([string]::IsNullOrWhiteSpace($currentRadio)) { $currentRadio='0' }
    $obj = & $ensure $currentRadio

    if ($line -match '(?i)\bRx\s*retry\s*frames?\b')  { $n=Get-LastNumber $line; if($n -ne $null){ $obj.RxRetry=[double]$n } }
    if ($line -match '(?i)\bRx\s*CRC\s*Errors?\b')    { $n=Get-LastNumber $line; if($n -ne $null){ $obj.RxCRC  =[double]$n } }
    if ($line -match '(?i)\bRX\s*PLCP\s*Errors?\b')   { $n=Get-LastNumber $line; if($n -ne $null){ $obj.RxPLCP =[double]$n } }
    if ($line -match '(?i)\bChannel\s*Changes?\b')    { $n=Get-LastNumber $line; if($n -ne $null){ $obj.ChannelChanges=[double]$n } }
    if ($line -match '(?i)\bTX\s*Power\s*Changes?\b') { $n=Get-LastNumber $line; if($n -ne $null){ $obj.TxPowerChanges=[double]$n } }

    if ($line -match '(?i)\bChannel\s*busy\b') {
      # ★ 修正：PSReference を正しく [ref] で渡す
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

  return New-Object psobject -Property @{ Path=$Path; OutputTime=$otUtc; Data=$data; IsRadio=$true }
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

# ===== 入力の読み込み =====
$segments=@()
$allInputFiles=@()
$radioParsed=@()

if ($SnapshotFiles -and $SnapshotFiles.Count -ge 1) {
  $fileList = Expand-SnapshotInputs -Inputs $SnapshotFiles
  if ($fileList.Count -lt 2) { throw "SnapshotFiles: 指定から2つ以上のファイルが見つかりません。フォルダ直下に2つ以上置くか、複数ファイル/ワイルドカードを指定してください。" }
  $allInputFiles = $fileList
  foreach ($f in $fileList) {
    $lines=@(); try{ $lines=Get-Content -LiteralPath $f -Encoding UTF8 -TotalCount 60 }catch{ $lines=@() }
    if (Is-RadioStatsLines -Lines $lines) { $radioParsed += (Parse-RadioStatsFile -Path $f) }
  }
  if ($radioParsed.Count -lt 2) { throw "radio-stats の候補が2つ未満でした。arm history / aps / bss-table / radio-info はスナップショット対象外です。radio-stats の出力ファイルを2つ以上配置してください。" }
  $sorted = $radioParsed | Sort-Object { if ($_.OutputTime -ne $null) { $_.OutputTime } else { [System.IO.File]::GetLastWriteTime($_.Path) } }
  for ($i=0; $i -lt $sorted.Count-1; $i++) {
    $b=$sorted[$i]; $a=$sorted[$i+1]
    $bt=$b.OutputTime; $at=$a.OutputTime
    $sec=0
    if ($bt -ne $null -and $at -ne $null) { try { $sec=[int][Math]::Abs(($at - $bt).TotalSeconds) } catch { $sec=0 } }
    if ($sec -le 0) { try { $t1=[System.IO.File]::GetLastWriteTime($b.Path); $t2=[System.IO.File]::GetLastWriteTime($a.Path); $sec=[int]([Math]::Abs(($t2 - $t1).TotalSeconds)) } catch { $sec=0 } }
    if ($sec -le 0) { $sec=900 }
    $segments += (New-Object psobject -Property @{ Before=$b; After=$a; DurationSec=$sec; StartJst=(Convert-ToJst $bt); EndJst=(Convert-ToJst $at) })
  }
}
elseif (-not [string]::IsNullOrWhiteSpace($BeforeFile) -and -not [string]::IsNullOrWhiteSpace($AfterFile)) {
  $b = Parse-RadioStatsFile -Path $BeforeFile
  $a = Parse-RadioStatsFile -Path $AfterFile
  if (-not $b.IsRadio -or -not $a.IsRadio) { throw "指定ファイルが radio-stats 形式ではありません。radio-stats 出力を指定してください。" }
  $bt=$b.OutputTime; $at=$a.OutputTime
  $allInputFiles = @($BeforeFile,$AfterFile)
  if ($DurationSec -le 0) {
    $sec=0
    if ($bt -ne $null -and $at -ne $null) { try { $sec=[int][Math]::Abs(($at - $bt).TotalSeconds) } catch { $sec=0 } }
    if ($sec -le 0) { try { $t1=[System.IO.File]::GetLastWriteTime($BeforeFile); $t2=[System.IO.File]::GetLastWriteTime($AfterFile); $sec=[int]([Math]::Abs(($t2 - $t1).TotalSeconds)) } catch { $sec=0 } }
    if ($sec -le 0) { $sec=900 }
    $DurationSec=$sec
  } else { $sec=$DurationSec }
  $segments += (New-Object psobject -Property @{ Before=$b; After=$a; DurationSec=$sec; StartJst=(Convert-ToJst $bt); EndJst=(Convert-ToJst $at) })
}
else {
  throw "単区間比較は -BeforeFile/-AfterFile、時系列集計は -SnapshotFiles に（フォルダ/ワイルドカード/ファイルのいずれかを）1個以上指定してください（展開後2つ以上の radio-stats ファイルが必要）。"
}

# ===== 補完用バックアップを構築 =====
$dirs = New-Object System.Collections.Generic.HashSet[string]
foreach ($seg in $segments) {
  try { $d1 = Split-Path -Path $seg.Before.Path -Parent } catch { $d1 = $null }
  try { $d2 = Split-Path -Path $seg.After.Path  -Parent } catch { $d2 = $null }
  if (-not [string]::IsNullOrWhiteSpace($d1)) { [void]$dirs.Add($d1) }
  if (-not [string]::IsNullOrWhiteSpace($d2)) { [void]$dirs.Add($d2) }
}
$dirList=@(); foreach ($d in $dirs) { $dirList += $d }
$BackupMap = Build-Backup-From-Dirs -Dirs $dirList

# ===== CSV 出力先 =====
if ([string]::IsNullOrWhiteSpace($OutputCsv)) {
  $baseRef = $segments[$segments.Count-1].After.Path
  $outDir = Get-ParentOrCwd -PathLike $baseRef
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $OutputCsv = Join-Path -Path $outDir -ChildPath ("aruba_radio_stats_diff_{0}.csv" -f $ts)
}

# ===== 列定義 =====
$colDefs = @(
  @('AP','AP 名'),
  @('Radio','ラジオ番号（0/1 等）'),
  @('Channel','運用チャネル番号（取得/補完）'),
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

# ===== 診断 =====
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

  if($root -ne 'inter' -and ($scoreInter -gt 0)){ $secList.Add((Sanitize-Text ("副次：非Wi-Fi干渉の疑い（Interference "+([int]$CCAI)+"%、PLCP "+([int]$PLCP)+"/s）"))) }
  if($root -ne 'co'   -and ($scoreCoch -gt 0)){ $secList.Add((Sanitize-Text ("副次：同一チャネル混在（他BSS "+([int]$CCAO)+"%、ΔCCA "+([int]$deltaCCA)+"pt、Busy64 "+([int]$Busy64)+"%）"))) }
  if($root -ne 'qual' -and ($scoreQual -gt 0)){ $secList.Add((Sanitize-Text ("副次：端末側の品質低下（Retry "+([int]$Retry)+"/s、CRC "+([int]$CRC)+"/s）"))) }
  if($root -ne 'busy' -and ($scoreBusy -gt 0)){ $secList.Add((Sanitize-Text ("副次：混雑（Busy64 "+([int]$Busy64)+"%）"))) }

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

function Diff-NonNegative { param([double]$After,[double]$Before)
  if ($After -eq $null -or $Before -eq $null) { return $null }
  $d=$After-$Before
  if ($d -lt 0) { return 0.0 }
  return $d
}

# ===== 集計 =====
$rows=@()
$cards=@()
$hourBuckets=@{}

function Get-FileTimeJst { param([string]$p) try { return Convert-ToJst ([System.IO.File]::GetLastWriteTime($p)) } catch { return $null } }

foreach ($seg in $segments) {
  $before=$seg.Before.Data; $after=$seg.After.Data; $dur=$seg.DurationSec

  $startDisp=$seg.StartJst
  if ($startDisp -eq $null -and $seg.Before -ne $null -and -not [string]::IsNullOrWhiteSpace($seg.Before.Path)) { $startDisp=Get-FileTimeJst -p $seg.Before.Path }
  $endDisp=$seg.EndJst
  if ($endDisp -eq $null -and $seg.After -ne $null -and -not [string]::IsNullOrWhiteSpace($seg.After.Path)) { $endDisp=Get-FileTimeJst -p $seg.After.Path }

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

    $ap = Fix-AP-From-Backup -ap $ap -path $seg.After.Path -Backup $BackupMap

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

    if (-not [string]::IsNullOrWhiteSpace($ap) -and $BackupMap.ContainsKey($ap)) {
      $slot = $BackupMap[$ap]
      if ($chan -eq $null) {
        $targetBand=$band; if ([string]::IsNullOrWhiteSpace($targetBand)) { $targetBand=Guess-Band-From-RadioIndex $radio }
        if     ($targetBand -eq '2.4GHz' -and $slot.Ch24 -ne $null) { $chan=[int]$slot.Ch24 }
        elseif ($targetBand -eq '5GHz'   -and $slot.Ch5  -ne $null) { $chan=[int]$slot.Ch5 }
        elseif ($targetBand -eq '6GHz'   -and $slot.Ch6  -ne $null) { $chan=[int]$slot.Ch6 }
        else {
          $count=0;$last=''; if($slot.Ch24 -ne $null){$count++;$last='2.4GHz'}; if($slot.Ch5 -ne $null){$count++;$last='5GHz'}; if($slot.Ch6 -ne $null){$count++;$last='6GHz'}
          if ($count -eq 1) { if($last -eq '2.4GHz'){$chan=[int]$slot.Ch24}elseif($last -eq '5GHz'){$chan=[int]$slot.Ch5}else{$chan=[int]$slot.Ch6}; $band=$last }
        }
      }
      if ([string]::IsNullOrWhiteSpace($band)) {
        if ($chan -ne $null) { $band=Get-BandFromChannel $chan }
        if ([string]::IsNullOrWhiteSpace($band)) {
          if ($slot.Ch24 -ne $null -and $slot.Ch5 -eq $null -and $slot.Ch6 -eq $null) { $band='2.4GHz' }
          elseif ($slot.Ch5 -ne $null -and $slot.Ch24 -eq $null -and $slot.Ch6 -eq $null) { $band='5GHz' }
          elseif ($slot.Ch6 -ne $null -and $slot.Ch24 -eq $null -and $slot.Ch5 -eq $null) { $band='6GHz' }
        }
      }
      if ($chan -eq $null -and -not [string]::IsNullOrWhiteSpace($band)) {
        if ($band -eq '2.4GHz' -and $slot.Ch24 -ne $null) { $chan=[int]$slot.Ch24 }
        if ($band -eq '5GHz'   -and $slot.Ch5  -ne $null) { $chan=[int]$slot.Ch5 }
        if ($band -eq '6GHz'   -and $slot.Ch6  -ne $null) { $chan=[int]$slot.Ch6 }
      }
    }
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
    $vals=@(); foreach ($def in $colDefs) { $name=$def[0]; $v=$rowObj.PSObject.Properties[$name].Value; if ($v -eq $null) { $vals+='' } else { $vals+=$v.ToString() } }
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
  foreach ($pair in $cols) { [void]$sb.AppendLine('<li><b>'+ (HtmlEscape $pair[0]) +'</b>：'+ (HtmlEscape $pair[1]) +'</li>') }
  [void]$sb.AppendLine('</ul></div></details>')

  [void]$sb.AppendLine('<table id="tbl"><thead><tr>')
  foreach ($pair in $cols) { $name=$pair[0]; $desc=$pair[1]; [void]$sb.AppendLine('<th title="'+ (HtmlEscape $desc) +'">'+ (HtmlEscape $name) +'</th>') }
  [void]$sb.AppendLine('</tr></thead><tbody>')
  foreach ($r in $rows) {
    [void]$sb.AppendLine('<tr>')
    foreach ($pair in $cols) {
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
}
exit 0