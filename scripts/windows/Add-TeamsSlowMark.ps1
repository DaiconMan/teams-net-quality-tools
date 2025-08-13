# Save as: Add-TeamsSlowMark.ps1
$LogDir  = Join-Path $env:LOCALAPPDATA "TeamsNet"
$MarkCsv = Join-Path $LogDir "user_marks.csv"
if(!(Test-Path $LogDir)){ New-Item -ItemType Directory -Path $LogDir | Out-Null }
if(!(Test-Path $MarkCsv)){
  "timestamp,username,computer,conn_type,ssid,bssid,signal_pct,comment" | Out-File -FilePath $MarkCsv -Encoding utf8
}

function Get-WifiContext {
  $ssid=$null; $bssid=$null; $signal=$null; $type="wired_or_disconnected"

  # netsh 蜃ｺ蜉帙ｒ縺昴・縺ｾ縺ｾ蜿門ｾ暦ｼ域枚蟄怜喧縺代＠縺ｫ縺上＞闍ｱ謨ｰ蟄励・縺ｿ縺ｧ謚ｽ蜃ｺ・・
  $lines = netsh.exe wlan show interfaces 2>$null
  if ($LASTEXITCODE -eq 0 -and $lines) {
    # 1) 謗･邯壼愛螳壹・ BSSID 縺ｮ譛臥┌縺ｧ陦後≧・・SSID 縺後≠繧後・謗･邯壻ｸｭ縺ｨ縺ｿ縺ｪ縺呻ｼ・
    $bidx = $null
    for($i=0; $i -lt $lines.Count; $i++){
      if ($lines[$i] -match '^\s*BSSID\s*:\s*(([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2})'){
        $bssid = $Matches[1].Replace('-',':').ToLower()
        $type = "wifi"
        $bidx = $i
        break
      }
    }

    if ($type -eq "wifi") {
      # 2) 蜷後§繧ｻ繧ｯ繧ｷ繝ｧ繝ｳ蜀・ｼ亥燕蠕・5陦檎ｨ句ｺｦ・峨°繧・SSID 繧貞叙蠕・
      $start = [Math]::Max(0, $bidx - 15)
      $end   = [Math]::Min($lines.Count-1, $bidx + 15)
      for($j=$start; $j -le $end; $j++){
        if ($null -eq $ssid -and $lines[$j] -match '^\s*SSID\s*:\s*(.+)$') {
          $ssid = $Matches[1].Trim()
        }
        # 3) 繧ｷ繧ｰ繝翫Ν縺ｯ縲・縲阪ｒ蜷ｫ繧譛蛻昴・謨ｰ蛟､% 繧呈鏡縺・ｼ医Λ繝吶Ν縺ｯ險隱樣撼萓晏ｭ假ｼ・
        if ($null -eq $signal -and $lines[$j] -match ':\s*([0-9]{1,3})\s*%') {
          $signal = [int]$Matches[1]
        }
        if ($ssid -and $signal) { break }
      }
    }
  }

  # 繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ・夂┌邱唔F縺袈p縺ｪ繧・type=wifi 縺ｫ・・SID/BSSID荳肴・縺ｧ繧ゑｼ・
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
$comment = Read-Host "・井ｻｻ諢擾ｼ臥裸迥ｶ繝｡繝｢繧貞・蜉帙＠縺ｦEnter・育ｩｺ縺ｧ繧０K・・
$line = "{0},{1},{2},{3},{4},{5},{6},""{7}""" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"),
  $env:USERNAME,$env:COMPUTERNAME,$ctx.type,$ctx.ssid,$ctx.bssid,$ctx.signal_pct,$comment
Add-Content -Path $MarkCsv -Value $line
Write-Host "險倬鹸縺励∪縺励◆・・line"
Start-Sleep -Seconds 1
