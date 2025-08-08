#!/usr/bin/env bash
# Save as: add_teams_slow_mark.command (double-clickable on mac, also runs in Linux)
OUTDIR="${XDG_STATE_HOME:-$HOME/.local/state}/teamsnet"
OUTCSV="$OUTDIR/user_marks.csv"
mkdir -p "$OUTDIR"
[ -f "$OUTCSV" ] || echo "timestamp,username,computer,conn_type,ssid,bssid,signal,comment" > "$OUTCSV"

conn_type="wired_or_disconnected"; ssid=""; bssid=""; signal=""
if [ "$(uname)" = "Darwin" ]; then
  AIRPORT="/System/Library/PrivateFrameworks/Apple80211.framework/Versions/Current/Resources/airport"
  if [ -x "$AIRPORT" ]; then
    info="$("$AIRPORT" -I 2>/dev/null)"
    ssid=$(echo "$info"  | awk '/ SSID:/{sub(/^.*SSID: /,"");print}')
    bssid=$(echo "$info" | awk '/ BSSID:/{print $2}' | tr '[:upper:]' '[:lower:]')
    signal=$(echo "$info"  | awk '/ agrCtlRSSI:/{print $2}')
    [ -n "$ssid$bssid" ] && conn_type="wifi"
  fi
else
  if command -v nmcli >/dev/null 2>&1; then
    line=$(nmcli -t -f active,ssid,bssid,signal dev wifi | grep "^yes:" | head -n1)
    if [ -n "$line" ]; then
      conn_type="wifi"
      ssid=$(echo "$line" | cut -d: -f2)
      bssid=$(echo "$line" | cut -d: -f3 | tr '[:upper:]' '[:lower:]')
      signal=$(echo "$line"| cut -d: -f4)
    fi
  elif command -v iw >/dev/null 2>&1; then
    dev=$(iw dev | awk '/Interface/ {print $2; exit}')
    if [ -n "$dev" ]; then
      info=$(iw dev "$dev" link)
      if echo "$info" | grep -q "Connected"; then
        conn_type="wifi"
        ssid=$(echo "$info" | awk -F'ssid ' '/ssid/ {print $2; exit}')
        bssid=$(echo "$info"| awk '/Connected to/ {print $3; exit}' | tr '[:upper:]' '[:lower:]')
        signal=$(echo "$info"  | awk '/signal/ {print $2; exit}')
      fi
    fi
  fi
fi

read -p "（任意）症状メモ（空OK）： " comment
printf '%s,%s,%s,%s,%s,%s,%s,"%s"\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$(whoami)" "$(hostname)" "$conn_type" "$ssid" "$bssid" "$signal" "$comment" >> "$OUTCSV"
echo "記録しました。"
