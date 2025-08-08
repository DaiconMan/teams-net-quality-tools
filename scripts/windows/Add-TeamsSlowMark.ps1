# Save as: Add-TeamsSlowMark.ps1
$LogDir  = Join-Path $env:LOCALAPPDATA "TeamsNet"
$MarkCsv = Join-Path $LogDir "user_marks.csv"
if(!(Test-Path $LogDir)){ New-Item -ItemType Directory -Path $LogDir | Out-Null }
if(!(Test-Path $MarkCsv)){
  "timestamp,username,computer,conn_type,ssid,bssid,signal_pct,comment" | Out-File -FilePath $MarkCsv -Encoding utf8
}

function Get-WifiContext {
  $ssid=$null;$bssid=$null;$signal=$null;$type="wired_or_disconnected"
  $out = (netsh wlan show interfaces 2>$null) -join "`n"
  if($LASTEXITCODE -eq 0 -and $out){
    if($out -match "(?im)^\s*SSID\s*:\s*(.+)$"){ $ssid = $Matches[1].Trim() }
    $m = [regex]::Match($out,"(?im)^\s*BSSID\s*:\s*(([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2})")
    if($m.Success){ $bssid = $m.Groups[1].Value.Replace('-',':').ToLower() }
    if($out -match "(?im)^\s*Signal\s*:\s*([0-9]{1,3})%"){ $signal = [int]$Matches[1] }
    if($ssid -or $bssid){ $type = "wifi" }
  }
  [pscustomobject]@{ type=$type; ssid=$ssid; bssid=$bssid; signal_pct=$signal }
}

$ctx = Get-WifiContext
$comment = Read-Host "（任意）症状メモを入力してEnter（空でもOK）"
$line = "{0},{1},{2},{3},{4},{5},{6},""{7}""" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"),
  $env:USERNAME,$env:COMPUTERNAME,$ctx.type,$ctx.ssid,$ctx.bssid,$ctx.signal_pct,$comment
Add-Content -Path $MarkCsv -Value $line
Write-Host "記録しました：$line"
Start-Sleep -Seconds 1
