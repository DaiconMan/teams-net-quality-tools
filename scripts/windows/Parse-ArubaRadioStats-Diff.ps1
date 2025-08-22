<# 
.SYNOPSIS
  Aruba "show ap debug radio-stats" スナップショットの差分/秒を算出し、CSV/HTML を生成（JST表示）。
  - 同フォルダの "show ap bss-table" / "show aps" テキストから AP名・Channel・Band を補完（バックアップ参照）
  - 表（CSV/HTML）の行末に SimpleDiag/Tips は出力しない（上部カードのみ）
  - 「Output Time:YYYY-MM-DD HH:mm:ss UTC」等を厳密にUTCとして解釈→JST変換
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

# ===== Output Time パース（UTCで返す） =====
function Extract-OutputTime {
  param([string[]]$Lines)
  if ($Lines -eq $null -or $Lines.Count -eq 0) { return $null }

  foreach ($raw in $Lines) { # "Output Time: ..." の右辺を優先
    $line = ($raw -replace '\r','').Trim()
    if ($line -match '^(?i)\s*Output\s*Time\s*[:=]\s*(.+)$') {
      $rhs = $Matches[1].Trim()

      if ($rhs -match '^\d{10}(\.\d+)?$') { try { $sec=[long]([double]$rhs); return ([System.DateTimeOffset]::FromUnixTimeSeconds($sec)).UtcDateTime } catch {} }
      if ($rhs -match '^\d{13}$') { try { $ms=[long]$rhs; return ([System.DateTimeOffset]::FromUnixTimeMilliseconds($ms)).UtcDateTime } catch {} }

      if ($rhs -match '^(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})\s*(UTC|GMT)\b') {
        try { $dt=[DateTime]::Parse($Matches[1],[System.Globalization.CultureInfo]::InvariantCulture); return [DateTime]::SpecifyKind($dt,[System.DateTimeKind]::Utc) } catch {}
      }

      $dto=[System.DateTimeOffset]::MinValue
      if ([System.DateTimeOffset]::TryParse($rhs,[System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::AssumeUniversal,[ref]$dto)) {
        return $dto.UtcDateTime
      }

      try { $dt2=[DateTime]::Parse($rhs,[System.Globalization.CultureInfo]::InvariantCulture); return [DateTime]::SpecifyKind($dt2,[System.DateTimeKind]::Utc) } catch {}
    }
  }

  $cands=@()
  foreach ($raw in $Lines) { if ($raw -match '(?i)(output\s*time|出力(時刻|時間|日時)|生成時刻)') { $cands+=$raw } }
  if ($cands.Count -eq 0) { return $null }

  foreach($line in $cands){
    $m=[regex]::Match($line,'(?<!\d)(\d{10})(?:\.\d+)?(?!\d)')
    if($m.Success){ try{ $sec=[long]([double]$m.Groups[1].Value); return ([System.DateTimeOffset]::FromUnixTimeSeconds($sec)).UtcDateTime }catch{} }
  }
  foreach($line in $cands){
    $m=[regex]::Match($line,'(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})\s*(UTC|GMT)\b',[System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if($m.Success){ try{ $dt=[DateTime]::Parse($m.Groups[1].Value,[System.Globalization.CultureInfo]::InvariantCulture); return [DateTime]::SpecifyKind($dt,[System.DateTimeKind]::Utc) }catch{} }
  }
  foreach($line in $cands){
    $m=[regex]::Match($line,'(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(?:Z|[+\-]\d{2}:\d{2})?)')
    if($m.Success){ $u = [System.DateTimeOffset]::Parse($m.Groups[1].Value); return $u.UtcDateTime }
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
  if ($ch -ge 1 -and $ch -le 233){ return '' } # 不明（6GHzの詳細chは環境差が大）
  return ''
}

# ===== Radio番号→Bandヒント（一般的なArubaの並び） =====
function Guess-Band-From-RadioIndex {
  param([string]$Radio)
  if ([string]::IsNullOrWhiteSpace($Radio)) { return '' }
  if ($Radio -eq '0') { return '2.4GHz' }
  if ($Radio -eq '1') { return '5GHz' }
  return ''
}

# ===== バックアップ参照（show ap bss-table / show aps） =====
# 格納形： $Backup[AP] = @{ Ch24=int?; T24=DateTime?; Ch5=int?; T5=DateTime?; Ch6=int?; T6=DateTime? }
function New-BackupSlot { return New-Object psobject -Property @{ Ch24=$null; T24=$null; Ch5=$null; T5=$null; Ch6=$null; T6=$null } }

function Detect-Band-Token {
  param([string]$line,[Nullable[int]]$ch)
  $s = $line.ToLower()
  if ($s -match '6ghz' -or $s -match '\b6g\b' -or $s -match '6e') { return '6GHz' }
  if ($s -match '\b11a\b' -or $s -match 'a/n' -or $s -match 'vht' -or $s -match 'he\b' -or $s -match 'eht') { return '5GHz' }
  if ($s -match '\b11b\b' -or $s -match '\b11g\b' -or $s -match '2\.4ghz' -or $s -match '\b2g\b') { return '2.4GHz' }
  if ($ch -ne $null) { return (Get-BandFromChannel $ch) }
  return ''
}

function Update-Backup {
  param([hashtable]$Backup,[string]$ap,[string]$band,[Nullable[int]]$ch,[Nullable[DateTime]]$ts)
  if ([string]::IsNullOrWhiteSpace($ap) -or [string]::IsNullOrWhiteSpace($band) -or $ch -eq $null) { return }
  if (-not $Backup.ContainsKey($ap)) { $Backup[$ap] = New-BackupSlot }
  $slot = $Backup[$ap]
  if ($band -eq '2.4GHz') {
    if ($slot.T24 -eq $null -or ($ts -ne $null -and $ts -gt $slot.T24)) { $slot.Ch24 = [int]$ch; $slot.T24 = $ts }
  } elseif ($band -eq '5GHz') {
    if ($slot.T5 -eq $null -or ($ts -ne $null -and $ts -gt $slot.T5))  { $slot.Ch5  = [int]$ch; $slot.T5  = $ts }
  } elseif ($band -eq '6GHz') {
    if ($slot.T6 -eq $null -or ($ts -ne $null -and $ts -gt $slot.T6))  { $slot.Ch6  = [int]$ch; $slot.T6  = $ts }
  }
}

function Parse-BssTable-File {
  param([string]$Path,[hashtable]$Backup)
  $ts = $null; try { $ts = [System.IO.File]::GetLastWriteTime($Path) } catch {}
  $lines = @(); try { $lines = Get-Content -LiteralPath $Path -Encoding UTF8 } catch { return }
  foreach ($raw in $lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    # AP名
    $ap = ''
    $m1 = [regex]::Match($line,'(?i)\bAP\s*Name\s*[:=]\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)')
    if ($m1.Success) { $ap = $m1.Groups[1].Value }
    if ([string]::IsNullOrWhiteSpace($ap)) {
      $m2 = [regex]::Match($line,'(?i)\bAP-?Name\b[:=]?\s*([^\s,]+)')
      if ($m2.Success) { $ap = $m2.Groups[1].Value }
    }
    if ([string]::IsNullOrWhiteSpace($ap)) {
      # 末尾にAP名が来る系を緩く拾う（空白区切り最後のトークン）
      if ($line -match '(?i)\bch(annel)?\b' -or $line -match '(?i)\bssid\b' -or $line -match '(?i)\bbssid\b') {
        $parts = $line -split '\s+'
        if ($parts.Length -ge 2) { $ap = $parts[$parts.Length-1] }
      }
    }
    if ([string]::IsNullOrWhiteSpace($ap)) { continue }

    # Channel
    $ch = $null
    $mCh = [regex]::Match($line,'(?i)\bCh(?:annel)?\s*[:=]?\s*(\d{1,3})\b')
    if ($mCh.Success) { try { $ch = [int]$mCh.Groups[1].Value } catch {} }
    if ($ch -eq $null) {
      $mCh2 = [regex]::Match($line,'(?i)\b(\d{1,3})\b\s*(?:MHz|HT|EIRP)')
      if ($mCh2.Success) { try { $ch = [int]$mCh2.Groups[1].Value } catch {} }
    }

    # Band
    $band = Detect-Band-Token -line $line -ch $ch
    if ([string]::IsNullOrWhiteSpace($band) -and $ch -ne $null) { $band = Get-BandFromChannel $ch }
    if ([string]::IsNullOrWhiteSpace($band) -or $ch -eq $null) { continue }

    Update-Backup -Backup $Backup -ap $ap -band $band -ch $ch -ts $ts
  }
}

function Parse-APS-File {
  param([string]$Path,[hashtable]$Backup)
  $ts = $null; try { $ts = [System.IO.File]::GetLastWriteTime($Path) } catch {}
  $lines = @(); try { $lines = Get-Content -LiteralPath $Path -Encoding UTF8 } catch { return }

  $currentAP = ''
  foreach ($raw in $lines) {
    $line = ($raw -replace '\r','').Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    # AP名の列／セクション開始をざっくり検出
    $mHead = [regex]::Match($line,'(?i)^(Name|AP\s*Name)\s*[:=]?\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)')
    if ($mHead.Success) { $currentAP = $mHead.Groups[2].Value }

    $mInline = [regex]::Match($line,'(?i)\bAP\s*Name\s*[:=]\s*([A-Za-z0-9_\-\.:\(\)\/\\]+)')
    if ($mInline.Success) { $currentAP = $mInline.Groups[1].Value }

    if ([string]::IsNullOrWhiteSpace($currentAP)) { 
      # テーブル行の先頭がAP名になっている形式（先頭トークン）
      $parts = $line -split '\s+'
      if ($parts.Length -ge 1 -and $parts[0] -match '^[A-Za-z0-9_\-\.:\(\)\/\\]+$') { $currentAP = $parts[0] }
    }

    if ([string]::IsNullOrWhiteSpace($currentAP)) { continue }

    # Radio N Channel = X / "Channel: X (5GHz)" など
    $mR = [regex]::Match($line,'(?i)\bRadio\s*([012])\s*.*?\bChannel\s*[:=]\s*(\d{1,3})')
    if ($mR.Success) {
      $r = $mR.Groups[1].Value
      $ch = $null; try { $ch = [int]$mR.Groups[2].Value } catch {}
      if ($ch -ne $null) {
        $band = Get-BandFromChannel $ch
        if ([string]::IsNullOrWhiteSpace($band)) { $band = Guess-Band-From-RadioIndex $r }
        Update-Backup -Backup $Backup -ap $currentAP -band $band -ch $ch -ts $ts
        continue
      }
    }

    $mCh = [regex]::Match($line,'(?i)\bChannel\s*[:=]\s*(\d{1,3})')
    if ($mCh.Success) {
      $ch = $null; try { $ch = [int]$mCh.Groups[1].Value } catch {}
      if ($ch -ne $null) {
        $band = Get-BandFromChannel $ch
        if ([string]::IsNullOrWhiteSpace($band)) { $band = Detect-Band-Token -line $line -ch $ch }
        Update-Backup -Backup $Backup -ap $currentAP -band $band -ch $ch -ts $ts
      }
    }
  }
}

function Build-Backup-From-Dirs {
  param([string[]]$Dirs)
  $bk = @{}
  if ($Dirs -eq $null) { return $bk }
  foreach ($d in $Dirs) {
    if ([string]::IsNullOrWhiteSpace($d)) { continue }
    if (-not (Test-Path -LiteralPath $d)) { continue }
    $files = @()
    try { $files = Get-ChildItem -LiteralPath $d -File -ErrorAction Stop } catch { $files = @() }
    foreach ($f in $files) {
      $name = $f.Name.ToLower()
      $isBss = ($name -match 'bss') -or ($name -match 'bss\-table') -or ($name -match 'show.*bss')
      $isAPS = ($name -match '\baps\b') -or ($name -match 'show.*aps')
      if (-not $isBss -and -not $isAPS) {
        # ファイル名から判別できない場合は中身を軽く覗いて判定（先頭数十行）
        $peek = @(); try { $peek = Get-Content -LiteralPath $f.FullName -Encoding UTF8 -TotalCount 40 } catch { $peek=@() }
        foreach ($l in $peek) {
          $ls = ($l -replace '\r','').ToLower()
          if ($ls -match 'bssid' -and $ls -match 'ssid') { $isBss = $true; break }
          if ($ls -match 'ap name' -and ($ls -match 'ip' -or $ls -match 'group')) { $isAPS = $true; break }
        }
      }
      if ($isBss) { Parse-BssTable-File -Path $f.FullName -Backup $bk }
      elseif ($isAPS) { Parse-APS-File -Path $f.FullName -Backup $bk }
    }
  }
  return $bk
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
$allInputFiles = @()

if ($SnapshotFiles -and $SnapshotFiles.Count -ge 1) {
  $fileList = Expand-SnapshotInputs -Inputs $SnapshotFiles
  if ($fileList.Count -lt 2) {
    throw "SnapshotFiles: 指定から2つ以上のファイルが見つかりません。フォルダ直下に2つ以上置くか、複数ファイル/ワイルドカードを指定してください。"
  }
  $allInputFiles = $fileList

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
  $allInputFiles = @($BeforeFile,$AfterFile)

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

# ===== 補完用バックアップの構築（radio-info と同じフォルダ想定） =====
$dirs = New-Object System.Collections.Generic.HashSet[string]
foreach ($seg in $segments) {
  try { $d1 = Split-Path -Path $seg.Before.Path -Parent } catch { $d1 = $null }
  try { $d2 = Split-Path -Path $seg.After.Path  -Parent } catch { $d2 = $null }
  if (-not [string]::IsNullOrWhiteSpace($d1)) { [void]$dirs.Add($d1) }
  if (-not [string]::IsNullOrWhiteSpace($d2)) { [void]$dirs.Add($d2) }
}
$dirList = @(); foreach ($d in $dirs) { $dirList += $d }
$BackupMap = Build-Backup-From-Dirs -Dirs $dirList

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

# CSV ヘッダ出力
$csvHeader = ($colDefs | ForEach-Object { $_[0] }) -join ','
Set-Content -LiteralPath $OutputCsv -Value $csvHeader -Encoding UTF8

# ===== 集計 =====
$rows = @()
$cards = @()
$hourBuckets = @{}

function Get-FileTimeJst { param([string]$p) try { return Convert-ToJst ([System.IO.File]::GetLastWriteTime($p)) } catch { return $null } }

foreach ($seg in $segments) {
  $before = $seg.Before.Data
  $after  = $seg.After.Data
  $dur    = $seg.DurationSec

  # 表示用時刻（JST）：OutputTime優先、無ければ更新時刻にフォールバック
  $startDisp = $seg.StartJst
  if ($startDisp -eq $null -and $seg.Before -ne $null -and -not [string]::IsNullOrWhiteSpace($seg.Before.Path)) { $startDisp = Get-FileTimeJst -p $seg.Before.Path }
  $endDisp = $seg.EndJst
  if ($endDisp -eq $null -and $seg.After -ne $null -and -not [string]::IsNullOrWhiteSpace($seg.After.Path)) { $endDisp = Get-FileTimeJst -p $seg.After.Path }

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

    # 差分
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

    # === バックアップ参照で補完 ===
    if ([string]::IsNullOrWhiteSpace($ap) -eq $false -and $BackupMap.ContainsKey($ap)) {
      $slot = $BackupMap[$ap]

      if ($chan -eq $null) {
        # まず既知Band、次にRadioヒント、最後に単一Bandのみのケースを試す
        $targetBand = $band
        if ([string]::IsNullOrWhiteSpace($targetBand)) { $targetBand = Guess-Band-From-RadioIndex $radio }
        if ($targetBand -eq '2.4GHz' -and $slot.Ch24 -ne $null) { $chan = [int]$slot.Ch24 }
        elseif ($targetBand -eq '5GHz' -and $slot.Ch5 -ne $null) { $chan = [int]$slot.Ch5 }
        elseif ($targetBand -eq '6GHz' -and $slot.Ch6 -ne $null) { $chan = [int]$slot.Ch6 }
        else {
          # 単一Bandしか見つからない場合はそれを採用
          $countBands = 0; $lastBand=''
          if ($slot.Ch24 -ne $null){ $countBands++; $lastBand='2.4GHz' }
          if ($slot.Ch5  -ne $null){ $countBands++; $lastBand='5GHz' }
          if ($slot.Ch6  -ne $null){ $countBands++; $lastBand='6GHz' }
          if ($countBands -eq 1) {
            if ($lastBand -eq '2.4GHz'){ $chan = [int]$slot.Ch24 }
            elseif ($lastBand -eq '5GHz'){ $chan = [int]$slot.Ch5 }
            elseif ($lastBand -eq '6GHz'){ $chan = [int]$slot.Ch6 }
            $band = $lastBand
          }
        }
      }

      if ([string]::IsNullOrWhiteSpace($band)) {
        if ($chan -ne $null) { $band = Get-BandFromChannel $chan }
        if ([string]::IsNullOrWhiteSpace($band)) {
          # Channelが無くBandのみ分かるケース
          if ($slot.Ch24 -ne $null -and $slot.Ch5 -eq $null -and $slot.Ch6 -eq $null) { $band='2.4GHz' }
          elseif ($slot.Ch5 -ne $null -and $slot.Ch24 -eq $null -and $slot.Ch6 -eq $null) { $band='5GHz' }
          elseif ($slot.Ch6 -ne $null -and $slot.Ch24 -eq $null -and $slot.Ch5 -eq $null) { $band='6GHz' }
        }
      }

      if ($chan -eq $null -and -not [string]::IsNullOrWhiteSpace($band)) {
        # Bandが決まっていてChannelが空ならBandに応じて埋める
        if ($band -eq '2.4GHz' -and $slot.Ch24 -ne $null) { $chan = [int]$slot.Ch24 }
        if ($band -eq '5GHz'   -and $slot.Ch5  -ne $null) { $chan = [int]$slot.Ch5 }
        if ($band -eq '6GHz'   -and $slot.Ch6  -ne $null) { $chan = [int]$slot.Ch6 }
      }
    }

    if ([string]::IsNullOrWhiteSpace($band) -and $chan -ne $null) { $band = Get-BandFromChannel $chan }

    # 空行抑止
    $hasAny=$false
    foreach($vv in @($retry_ps,$crc_ps,$plcp_ps,$chg_ph,$txp_ph,$busy1s,$busy4s,$busy64,$busyB,$txB,$rxB,$ccaO,$ccaOt,$ccaI,$chan)){
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
      if ($retry_ps -ne $null){ $h.Retry += $retry_ps }
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

  # 列は colDefs に準拠（Simple/Tips 列なし）
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