# Save as: Add-TeamsSlowMark.ps1
$LogDir  = Join-Path $env:LOCALAPPDATA "TeamsNet"
$MarkCsv = Join-Path $LogDir "user_marks.csv"
if(!(Test-Path $LogDir)){ New-Item -ItemType Directory -Path $LogDir | Out-Null }
if(!(Test-Path $MarkCsv)){
  "timestamp,username,computer,conn_type,ssid,bssid,signal_pct,comment" | Out-File -FilePath $MarkCsv -Encoding utf8
}

function Get-WifiContext {
  $ssid=$null; $bssid=$null; $signal=$null; $type="wired_or_disconnected"

  # netsh 出力をセクション（インターフェイス単位）に分解
  $sections=@()
  $out = netsh.exe wlan show interfaces 2>$null
  if ($LASTEXITCODE -eq 0 -and $out) {
    $buf=@()
    foreach($ln in $out){
      if($ln -match "^\s*(Name|名前)\s*:"){   # 新しいIFセクション開始
        if($buf.Count){ $sections += ,($buf -join "`n"); $buf=@() }
      }
      $buf += $ln
    }
    if($buf.Count){ $sections += ,($buf -join "`n") }

    foreach($sec in $sections){
      # 「接続済み」のセクションだけ採用（英/日対応）
      if($sec -match "(?im)^\s*(State|状態)\s*:\s*(connected|接続されています)"){
        if($sec -match "(?im)^\s*SSID\s*:\s*(.+)$"){ $ssid = $Matches[1].Trim() }
        $m = [regex]::Match($sec,"(?im)^\s*BSSID\s*:\s*(([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2})")
        if($m.Success){ $bssid = $m.Groups[1].Value.Replace('-',':').ToLower() }
        $sm = [regex]::Match($sec,"(?im)^\s*(Signal|シグナル)\s*:\s*([0-9]{1,3})%")
        if($sm.Success){ $signal = [int]$sm.Groups[2].Value }
        $type = "wifi"
        break
      }
    }
  }

  # フォールバック：netshパースに失敗しても、無線IFがUpなら wifi と判定
  if($type -ne "wifi"){
    try{
      $wifi = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Status -eq 'Up' -and ( $_.NdisPhysicalMedium -eq 'Native802_11' -or $_.InterfaceDescription -match 'Wireless|Wi-Fi' ) }
      if($wifi){ $type = "wifi" }
    } catch {}
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
