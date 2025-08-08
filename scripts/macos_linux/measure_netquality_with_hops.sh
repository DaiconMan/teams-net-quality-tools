#!/usr/bin/env bash
# Save as: measure_netquality_with_hops.sh
TARGETS=("world.tr.teams.microsoft.com" "teams.microsoft.com" "graph.microsoft.com" "prod.msocdn.com" "aka.ms")
HOP_TARGETS=("world.tr.teams.microsoft.com")
SAMPLES=10
INTERVAL=30
HOP_PING=5
HOP_EVERY=10
MAX_HOPS=25

OUTDIR="${XDG_STATE_HOME:-$HOME/.local/state}/teamsnet"
OUTCSV="$OUTDIR/teams_net_quality.csv"
HOPCSV="$OUTDIR/path_hop_quality.csv"
STATE="$OUTDIR/state.txt"        # prev_bssid|cycle
MAP="$OUTDIR/ap_map.csv"
mkdir -p "$OUTDIR"
[ -f "$OUTCSV" ] || echo "timestamp,host,icmp_avg_ms,icmp_jitter_ms,loss_pct,dns_ms,tcp_443_ms,http_head_ms,mos_estimate,conn_type,ssid,bssid,signal,ap_name,roamed,roam_from,roam_to,notes" > "$OUTCSV"
[ -f "$HOPCSV" ] || echo "timestamp,target,hop_index,hop_ip,icmp_avg_ms,icmp_jitter_ms,loss_pct,notes,conn_type,ssid,bssid,signal,ap_name,roamed,roam_from,roam_to" > "$HOPCSV"

read_state(){ if [ -f "$STATE" ]; then IFS='|' read -r PREV_BSSID CYCLE <"$STATE"; else PREV_BSSID=""; CYCLE=0; fi; }
write_state(){ printf '%s|%s' "$PREV_BSSID" "$CYCLE" > "$STATE"; }

dns_ms(){
  local h="$1"
  if command -v getent >/dev/null 2>&1; then { time getent hosts "$h" >/dev/null 2>&1; } 2> /tmp/dnstime.$$
  elif command -v host >/dev/null 2>&1; then { time host "$h" >/dev/null 2>&1; } 2> /tmp/dnstime.$$
  elif command -v dig  >/dev/null 2>&1; then { time dig +short "$h" >/dev/null 2>&1; } 2> /tmp/dnstime.$$
  else echo ""; return; fi
  awk '/real/ {print $2}' /tmp/dnstime.$$ | sed 's/m/:/; s/s//' | awk -F: '{print ($1*60000)+($2*1000)}'; rm -f /tmp/dnstime.$$
}
tcp_ms(){   curl -s -o /dev/null -w "%{time_connect}\n" "https://$1/" | awk '{print $1*1000}'; }
http_ms(){  curl -s -o /dev/null -w "%{time_starttransfer}\n" "https://$1/" | awk '{print $1*1000}'; }

icmp_stats(){
  local h="$1"; local cnt="${2:-$SAMPLES}"; local tmp=$(mktemp)
  ping -c "$cnt" "$h" > "$tmp" 2>/dev/null || { echo ',,,icmp_blocked'; rm -f "$tmp"; return; }
  local received=$(grep -Eo '[0-9]+ received' "$tmp" | awk '{print $1}')
  local loss=$(awk -v s="$cnt" -v r="$received" 'BEGIN{ if(s>0){print (s-r)*100.0/s}else{print 0}}')
  local rtts=($(grep -Eo 'time=[0-9\.]+' "$tmp" | cut -d= -f2))
  local count=${#rtts[@]}; local avg="" jit=""
  if [ "$count" -gt 0 ]; then
    local sum=0; for v in "${rtts[@]}"; do sum=$(awk -v a="$sum" -v b="$v" 'BEGIN{print a+b}'); done
    avg=$(awk -v s="$sum" -v c="$count" 'BEGIN{print s/c}')
    local varsum=0
    for v in "${rtts[@]}"; do varsum=$(awk -v v="$v" -v m="$avg" -v vs="$varsum" 'BEGIN{d=v-m; print vs + d*d}'); done
    jit=$(awk -v vs="$varsum" -v c="$count" 'BEGIN{print sqrt(vs/c)}')
    echo "$avg,$jit,$loss,"
  else
    echo ",,100,icmp_no_reply"
  fi
  rm -f "$tmp"
}

wifi_ctx(){
  if [ "$(uname)" = "Darwin" ]; then
    AIRPORT="/System/Library/PrivateFrameworks/Apple80211.framework/Versions/Current/Resources/airport"
    if [ -x "$AIRPORT" ]; then
      info="$("$AIRPORT" -I 2>/dev/null)"
      ssid=$(echo "$info"  | awk '/ SSID:/{sub(/^.*SSID: /,"");print}')
      bssid=$(echo "$info" | awk '/ BSSID:/{print $2}' | tr '[:upper:]' '[:lower:]')
      rssi=$(echo "$info"  | awk '/ agrCtlRSSI:/{print $2}')
      [ -n "$ssid$bssid" ] && { echo "wifi|$ssid|$bssid|$rssi"; return; }
    fi
    echo "wired_or_disconnected|||"
  else
    if command -v nmcli >/dev/null 2>&1; then
      line=$(nmcli -t -f active,ssid,bssid,signal dev wifi | grep "^yes:" | head -n1)
      if [ -n "$line" ]; then
        echo "wifi|$(echo "$line" | cut -d: -f2)|$(echo "$line" | cut -d: -f3 | tr '[:upper:]' '[:lower:]')|$(echo "$line" | cut -d: -f4)"
        return
      fi
    fi
    if command -v iw >/dev/null 2>&1; then
      dev=$(iw dev | awk '/Interface/ {print $2; exit}')
      if [ -n "$dev" ]; then
        info=$(iw dev "$dev" link)
        if echo "$info" | grep -q "Connected"; then
          ssid=$(echo "$info" | awk -F'ssid ' '/ssid/ {print $2; exit}')
          bssid=$(echo "$info"| awk '/Connected to/ {print $3; exit}' | tr '[:upper:]' '[:lower:]')
          sig=$(echo "$info"  | awk '/signal/ {print $2; exit}')
          echo "wifi|$ssid|$bssid|$sig"; return
        fi
      fi
    fi
    echo "wired_or_disconnected|||"
  fi
}

ap_name(){
  local b="$1"; [ -f "$MAP" ] || { echo ""; return; }
  awk -F, -v key="$b" 'NR>1 && tolower($1)==tolower(key){print $2; exit}' "$MAP"
}

get_hops(){
  local dst="$1"
  if command -v traceroute >/dev/null 2>&1; then
    traceroute -n -m "$MAX_HOPS" -w 1 "$dst" 2>/dev/null | grep -Eo '([0-9]{1,3}\.){3}[0-9]{1,3}'
  elif command -v mtr >/dev/null 2>&1; then
    mtr -n -r -c 1 "$dst" 2>/dev/null | awk 'NR>1 {print $2}'
  else
    echo ""
  fi
}

read_state
while true; do
  CYCLE=$((CYCLE+1))
  ctx=$(wifi_ctx)
  conn_type=$(echo "$ctx" | cut -d'|' -f1)
  ssid=$(echo "$ctx" | cut -d'|' -f2)
  bssid=$(echo "$ctx" | cut -d'|' -f3)
  signal=$(echo "$ctx"| cut -d'|' -f4)
  apname=$(ap_name "$bssid")

  roamed=""; roam_from=""; roam_to=""
  if [ "$conn_type" = "wifi" ] && [ -n "$bssid" ]; then
    if [ -n "$PREV_BSSID" ] && [ "$PREV_BSSID" != "$bssid" ]; then
      roamed="roamed"; roam_from="$PREV_BSSID"; roam_to="$bssid"
    fi
    PREV_BSSID="$bssid"
  else
    PREV_BSSID=""
  fi
  write_state

  ts=$(date "+%Y-%m-%d %H:%M:%S")
  for h in "${TARGETS[@]}"; do
    dns=$(dns_ms "$h")
    tcp=$(tcp_ms "$h"); tcp_note=$([ -z "$tcp" ] && echo "tcp_fail")
    http=$(http_ms "$h"); http_note=$([ -z "$http" ] && echo "http_fail")
    IFS=',' read icmp_avg icmp_jitter loss_pct icmp_note < <(icmp_stats "$h" "$SAMPLES")
    rtt=${icmp_avg:-$tcp}; rtt=${rtt%%.*}; [ -z "$rtt" ] && rtt=999
    pl=${loss_pct:-0}
    mos=$(awk -v r="$rtt" -v p="$pl" 'BEGIN{v=4.5 - 0.0004*r - 0.1*p; if(v<1)v=1; if(v>4.5)v=4.5; printf("%.2f",v)}')
    notes=$(printf "%s+%s+%s+%s" "$icmp_note" "$tcp_note" "$http_note" "$roamed" | sed 's/^+//; s/+$//; s/++/+/g')
    printf '%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"%s",%s,%s,%s,"%s"\n' \
      "$ts" "$h" "$icmp_avg" "$icmp_jitter" "$loss_pct" "$dns" "$tcp" "$http" "$mos" \
      "$conn_type" "$ssid" "$bssid" "$signal" "$apname" "$roamed" "$roam_from" "$roam_to" "$notes" >> "$OUTCSV"
  done

  if [ $((CYCLE % HOP_EVERY)) -eq 0 ]; then
    for ht in "${HOP_TARGETS[@]}"; do
      mapfile -t hops < <(get_hops "$ht")
      if [ ${#hops[@]} -eq 0 ]; then
        printf '%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"%s",%s,%s,%s\n' \
          "$ts" "$ht" 0 "" "" "" "" "trace_no_reply" "$conn_type" "$ssid" "$bssid" "$signal" "$apname" "$roamed" "$roam_from" "$roam_to" >> "$HOPCSV"
      else
        idx=0
        for ip in "${hops[@]}"; do
          idx=$((idx+1))
          IFS=',' read avg jit loss note < <(icmp_stats "$ip" "$HOP_PING")
          printf '%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"%s",%s,%s,%s\n' \
            "$ts" "$ht" "$idx" "$ip" "$avg" "$jit" "$loss" "$note" \
            "$conn_type" "$ssid" "$bssid" "$signal" "$apname" "$roamed" "$roam_from" "$roam_to" >> "$HOPCSV"
        done
      fi
    done
  fi

  sleep "$INTERVAL"
done
