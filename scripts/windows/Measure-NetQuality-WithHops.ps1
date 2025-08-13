# Measure-NetQuality-WithHops.ps1 (fixed full)
param(
  [string[]]$Targets = @(
    "world.tr.teams.microsoft.com","teams.microsoft.com",
    "graph.microsoft.com","prod.msocdn.com","aka.ms"
  ),
  [string[]]$HopTargets = @("world.tr.teams.microsoft.com"),
  [int]$SamplesPerCycle = 10,
  [int]$IntervalSeconds = 30,
  [int]$HopPingCount = 5,
  [int]$HopProbeEveryCycles = 10,
  [int]$MaxHops = 25
)

$LogDir = Join-Path $env:LOCALAPPDATA "TeamsNet"
$OutCsv = Join-Path $LogDir "teams_net_quality.csv"
$HopCsv = Join-Path $LogDir "path_hop_quality.csv"
$StateFile = Join-Path $LogDir "state.json"
$MapFile = Join-Path $LogDir "ap_map.csv"

if(!(Test-Path $LogDir)){ New-Item -ItemType Directory -Path $LogDir | Out-Null }
if(!(Test-Path $OutCsv)){
  "timestamp,host,icmp_avg_ms,icmp_jitter_ms,loss_pct,dns_ms,tcp_443_ms,http_head_ms,mos_estimate,conn_type,ssid,bssid,signal_pct,ap_name,roamed,roam_from,roam_to,notes" | Out-File -FilePath $OutCsv -Encoding utf8
}
if(!(Test-Path $HopCsv)){
  "timestamp,target,hop_index,hop_ip,icmp_avg_ms,icmp_jitter_ms,loss_pct,notes,conn_type,ssid,bssid,signal_pct,ap_name,roamed,roam_from,roam_to" | Out-File -FilePath $HopCsv -Encoding utf8
}

function Get-ApName([string]$bssid){
  if(!(Test-Path $MapFile) -or -not $bssid){ return $null }
  try{
    $row = Import-Csv $MapFile | Where-Object { $_.bssid.ToLower() -eq $bssid.ToLower() } | Select-Object -First 1
    if($row){ return $row.ap_name } else { return $null }
  }catch{ return $null }
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

function Measure-DnsTime([string]$target){
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  try { [System.Net.Dns]::GetHostAddresses($target) | Out-Null } catch {}
  $sw.Stop(); [math]::Round($sw.Elapsed.TotalMilliseconds,2)
}

function Measure-TcpTime([string]$target,[int]$port=443){
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $note = ""
  $client = New-Object System.Net.Sockets.TcpClient
  try { $client.Connect($target,$port) } catch { $note = "tcp_fail" }
  $client.Close()
  $sw.Stop()
  $ms = [math]::Round($sw.Elapsed.TotalMilliseconds,2)
  return @($ms, $note)
}

function Measure-HttpHead([string]$target){
  $uri = "https://$target/"
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $note = ""
  try { Invoke-WebRequest -Uri $uri -Method Head -UseBasicParsing -TimeoutSec 10 | Out-Null } catch { $note = "http_fail" }
  $sw.Stop()
  $ms = [math]::Round($sw.Elapsed.TotalMilliseconds,2)
  return @($ms, $note)
}

function Measure-Icmp([string]$target,[int]$count){
  try{
    $pings = Test-Connection -ComputerName $target -Count $count -ErrorAction Stop
    $sent = $count; $recv = @($pings).Count
    $loss = if ($sent -gt 0) { 100.0 * ($sent - $recv) / $sent } else { 0 }
    $rtts = @(); foreach($p in $pings){ $rtts += $p.ResponseTime }
    if($rtts.Count){
      $avg = [math]::Round(($rtts | Measure-Object -Average).Average,2)
      $mean = ($rtts | Measure-Object -Average).Average
      $var = ($rtts | ForEach-Object { ($_ - $mean) * ($_ - $mean) } | Measure-Object -Sum).Sum / $rtts.Count
      $jitter = [math]::Round([math]::Sqrt($var),2)
      return @($avg,$jitter,[math]::Round($loss,2),$null)
    } else { return @($null,$null,100,"icmp_no_reply") }
  }catch{ return @($null,$null,$null,"icmp_blocked") }
}

function Get-HopsForTarget([string]$target,[int]$maxHops=25) {
  # IPv4蠑ｷ蛻ｶ・・4・会ｼ矩・ｼ輔″OFF・・d・会ｼ九ち繧､繝繧｢繧ｦ繝育洒繧・
  # Write-Host $target
  $out = tracert.exe -4 -d -h $maxHops -w 800 $target 2>$null
  if(-not $out){ return @() }
  $ips = @()
  foreach($line in $out){
    if($line -notmatch '^\s*\d+\s') { continue }            # hop逡ｪ蜿ｷ縺ｧ蟋九∪繧玖｡後・縺ｿ
    $m = [regex]::Match($line, '(\d{1,3}(?:\.\d{1,3}){3})')  # 陦悟・縺ｮ譛蛻昴・IPv4
    if($m.Success){ $ips += $m.Value }
  }
  # Write-Host $ips
  return $ips
}

function Measure-HopStats([string]$target,[string[]]$hopIps,[int]$count){
  $idx = 0
  foreach($ip in $hopIps){
    $idx++
    $avg=$null;$jit=$null;$loss=$null;$note=$null
    $avg,$jit,$loss,$note = Measure-Icmp -target $ip -count $count
    [pscustomobject]@{ hop_index=$idx; hop_ip=$ip; avg=$avg; jitter=$jit; loss=$loss; note=$note }
  }
}

# --- CSV繝ｭ繝・け蟇ｾ遲厄ｼ夂ｰ｡譏薙・蝣・欧迚・---
function Append-Line {
  param([string]$Path,[string]$Line)
  try {
    Add-Content -Path $Path -Value $Line -ErrorAction Stop
  } catch [System.IO.IOException] {
    # 譛ｬ菴薙′繝ｭ繝・け荳ｭ・・xcel髢九″縺｣縺ｱ縺ｪ縺礼ｭ会ｼ峨・ .queue 縺ｫ騾驕ｿ
    Add-Content -Path ($Path + ".queue") -Value $Line
  }
}

function Flush-Queue {
  param([string]$Path)
  $q = $Path + ".queue"
  if (Test-Path $q) {
    try {
      # 騾驕ｿ蛻・ｒ荳豌励↓蜷域ｵ・
      Get-Content $q -ErrorAction Stop | Add-Content -Path $Path -ErrorAction Stop
      Remove-Item $q -Force
    } catch {
      # 縺ｾ縺繝ｭ繝・け荳ｭ縺ｪ繧画ｬ｡繧ｵ繧､繧ｯ繝ｫ縺ｧ蜀肴倦謌ｦ
    }
  }
}
# -----------------------------------------------------------------------

# 迥ｶ諷玖ｪｭ縺ｿ霎ｼ縺ｿ
$prevBssid = $null; $cycle = 0
if(Test-Path $StateFile){
  try { $st = Get-Content $StateFile -Raw | ConvertFrom-Json; $prevBssid = $st.prev_bssid; $cycle = [int]$st.cycle } catch {}
}

while($true){
  $cycle++
  $ctx = Get-WifiContext
  $apName = Get-ApName $ctx.bssid

  # 繝ｭ繝ｼ繝溘Φ繧ｰ讀懃衍
  $roamed = ""; $roamFrom = ""; $roamTo = ""
  if($ctx.type -eq "wifi" -and $ctx.bssid){
    if($prevBssid -and ($prevBssid.ToLower() -ne $ctx.bssid.ToLower())){
      $roamed = "roamed"; $roamFrom = $prevBssid; $roamTo = $ctx.bssid
    }
    $prevBssid = $ctx.bssid
  } else { $prevBssid = $null }

  @{ prev_bssid = $prevBssid; cycle = $cycle } | ConvertTo-Json | Set-Content -Path $StateFile -Encoding utf8

  # --- 繧ｨ繝ｳ繝峨ヤ繝ｼ繧ｨ繝ｳ繝・---
  foreach($t in $Targets){
    $dns = Measure-DnsTime $t
    $tcpMs,$tcpNote = Measure-TcpTime $t 443
    $httpMs,$httpNote = Measure-HttpHead $t
    $avg=$null;$jit=$null;$loss=$null;$icmpNote=$null
    $avg,$jit,$loss,$icmpNote = Measure-Icmp -target $t -count $SamplesPerCycle

    # 螳牙・縺ｫ謨ｰ蛟､蛹・
    $rtt = 999.0
    if($null -ne $avg -and -not ($avg -is [System.Array])) { $rtt = [double]$avg }
    elseif($null -ne $tcpMs -and -not ($tcpMs -is [System.Array])) { $rtt = [double]$tcpMs }

    $pl = 0.0
    if($null -ne $loss -and -not ($loss -is [System.Array])) { $pl = [double]$loss }

    $mos = [math]::Round([math]::Max(1,[math]::Min(4.5, 4.5 - 0.0004*$rtt - 0.1*$pl)),2)

    $notes = (@($icmpNote,$tcpNote,$httpNote,$roamed) | Where-Object { $_ -and $_ -ne "" }) -join '+'

    $line = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17}" -f `
      (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $t,$avg,$jit,$loss,$dns,$tcpMs,$httpMs,$mos, `
      $ctx.type,$ctx.ssid,$ctx.bssid,$ctx.signal_pct,$apName,$roamed,$roamFrom,$roamTo,$notes

    Append-Line $OutCsv $line
  }

  # --- 繝帙ャ繝・---
  if( ($cycle % $HopProbeEveryCycles) -eq 0 ){
    foreach($ht in $HopTargets){
      $hops = Get-HopsForTarget -target $ht -maxHops $MaxHops
      if($hops.Count -gt 0){
        $stats = Measure-HopStats -target $ht -hopIps $hops -count $HopPingCount
        $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        foreach($s in $stats){
          $line = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}" -f `
            $ts,$ht,$s.hop_index,$s.hop_ip,$s.avg,$s.jitter,$s.loss,$s.note, `
            $ctx.type,$ctx.ssid,$ctx.bssid,$ctx.signal_pct,$apName,$roamed,$roamFrom,$roamTo
          Append-Line $HopCsv $line
        }
      } else {
        $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $line = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}" -f `
          $ts,$ht,0,"", "", "", "", "tracert_no_reply", `
          $ctx.type,$ctx.ssid,$ctx.bssid,$ctx.signal_pct,$apName,$roamed,$roamFrom,$roamTo
        Append-Line $HopCsv $line
      }
    }
  }

  # Excel髢峨§縺溘ち繧､繝溘Φ繧ｰ縺ｧ蜷域ｵ√ｒ隧ｦ縺ｿ繧・
  Flush-Queue $OutCsv
  Flush-Queue $HopCsv

  Start-Sleep -Seconds $IntervalSeconds
}
