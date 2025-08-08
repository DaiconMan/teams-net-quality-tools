# Measure-NetQuality-WithHops.ps1 (fixed)
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

function Get-ApName($bssid){
  if(!(Test-Path $MapFile) -or -not $bssid){ return $null }
  try{
    $row = Import-Csv $MapFile | Where-Object { $_.bssid.ToLower() -eq $bssid.ToLower() } | Select-Object -First 1
    if($row){ return $row.ap_name } else { return $null }
  }catch{ return $null }
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

function Measure-DnsTime([string]$target){
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  try { [System.Net.Dns]::GetHostAddresses($target) | Out-Null } catch {}
  $sw.Stop(); [math]::Round($sw.Elapsed.TotalMilliseconds,2)
}
function Measure-TcpTime([string]$target,[int]$port=443){
  $sw = [System.Diagnostics.Stopwatch]::StartNew(); $note=""
  $client = New-Object System.Net.Sockets.TcpClient
  try { $client.Connect($target,$port) } catch { $note="tcp_fail" }
  $client.Close(); $sw.Stop()
  ,([math]::Round($sw.Elapsed.TotalMilliseconds,2)),$note
}
function Measure-HttpHead([string]$target){
  $uri = "https://$target/"; $sw=[System.Diagnostics.Stopwatch]::StartNew(); $note=""
  try { Invoke-WebRequest -Uri $uri -Method Head -UseBasicParsing -TimeoutSec 10 | Out-Null } catch { $note="http_fail" }
  $sw.Stop(); ,([math]::Round($sw.Elapsed.TotalMilliseconds,2)),$note
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

function Get-HopsForTarget([string]$target,[int]$maxHops=25){
  $out = tracert.exe -d -h $maxHops -w 800 $target 2>$null
  if(-not $out){ return @() }
  $ips = @()
  foreach($line in $out){
    $m = [regex]::Matches($line, "(\d{1,3}(\.\d{1,3}){3})")
    if($m.Count -gt 0){ $ips += $m[0].Value }
  }
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

# 状態読み込み
$prevBssid = $null; $cycle = 0
if(Test-Path $StateFile){
  try { $st = Get-Content $StateFile -Raw | ConvertFrom-Json; $prevBssid = $st.prev_bssid; $cycle = [int]$st.cycle } catch {}
}

while($true){
  $cycle++
  $ctx = Get-WifiContext
  $apName = Get-ApName $ctx.bssid

  # ローミング検知
  $roamed = ""; $roamFrom = ""; $roamTo = ""
  if($ctx.type -eq "wifi" -and $ctx.bssid){
    if($prevBssid -and ($prevBssid.ToLower() -ne $ctx.bssid.ToLower())){
      $roamed = "roamed"; $roamFrom = $prevBssid; $roamTo = $ctx.bssid
    }
    $prevBssid = $ctx.bssid
  } else { $prevBssid = $null }
  @{ prev_bssid = $prevBssid; cycle = $cycle } | ConvertTo-Json | Set-Content -Path $StateFile -Encoding utf8

  # --- エンドツーエンド ---
  foreach($t in $Targets){
    $dns = Measure-DnsTime $t
    $tcpMs,$tcpNote = Measure-TcpTime $t 443
    $httpMs,$httpNote = Measure-HttpHead $t
    $avg=$null;$jit=$null;$loss=$null;$icmpNote=$null
    $avg,$jit,$loss,$icmpNote = Measure-Icmp -target $t -count $SamplesPerCycle

    $rtt = if($avg){ [double]$avg } elseif($tcpMs){ [double]$tcpMs } else { 999 }
    $pl  = if($loss){ [double]$loss } else { 0 }
    $mos = [math]::Round([math]::Max(1,[math]::Min(4.5, 4.5 - 0.0004*$rtt - 0.1*$pl)),2)

    $notes = (@($icmpNote,$tcpNote,$httpNote,$roamed) | Where-Object { $_ -and $_ -ne "" }) -join '+'
    $line = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17}" -f `
      (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $t,$avg,$jit,$loss,$dns,$tcpMs,$httpMs,$mos, `
      $ctx.type,$ctx.ssid,$ctx.bssid,$ctx.signal_pct,$apName,$roamed,$roamFrom,$roamTo,$notes
    Add-Content -Path $OutCsv -Value $line
  }

  # --- ホップ ---
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
          Add-Content -Path $HopCsv -Value $line
        }
      } else {
        $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $line = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15}" -f `
          $ts,$ht,0,"", "", "", "", "tracert_no_reply", `
          $ctx.type,$ctx.ssid,$ctx.bssid,$ctx.signal_pct,$apName,$roamed,$roamFrom,$roamTo
        Add-Content -Path $HopCsv -Value $line
      }
    }
  }

  Start-Sleep -Seconds $IntervalSeconds
}
