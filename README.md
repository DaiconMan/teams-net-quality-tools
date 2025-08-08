# teams-net-quality-tools

**Teams の体感遅延の見える化**のためのスクリプト集です。管理者権限なしで動作し、以下を記録します。

- エンドツーエンド品質（RTT/ジッタ/ロス、DNS/TCP/HTTP、簡易 MOS）
- 接続 AP（SSID / BSSID / 信号値）
- ローミング検知（BSSID の変化）
- 途中経路（各ホップ）のジッタ / ロス（指定サイクルごと）
- ユーザーの「今遅い！」ワンクリック時刻メモ

出力は CSV：
- `teams_net_quality.csv` … エンドツーエンド品質＋AP/ローミング
- `path_hop_quality.csv` … 途中経路ホップごとの品質
- `user_marks.csv` … ユーザーの手動マーカー（任意）

---

## フォルダ構成

```
scripts/
  windows/
    Measure-NetQuality-WithHops.ps1
    Add-TeamsSlowMark.ps1
  macos_linux/
    measure_netquality_with_hops.sh
    add_teams_slow_mark.command
samples/
  ap_map.csv
```

`samples/ap_map.csv` は **BSSID → AP 名** の対応表（任意）。

---

## 使い方（Windows）

1) **計測（権限不要）**
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\windows\Measure-NetQuality-WithHops.ps1" -IntervalSeconds 30 -HopProbeEveryCycles 10
```
- 出力：`%LOCALAPPDATA%\TeamsNet\teams_net_quality.csv` / `path_hop_quality.csv`
- 途中経路計測の頻度は `-HopProbeEveryCycles` で調整（例：10秒間隔×10=100秒ごと）

2) **「遅い」マーカー（任意）**
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\windows\Add-TeamsSlowMark.ps1"
```
- 出力：`%LOCALAPPDATA%\TeamsNet\user_marks.csv`

> 注：PowerShell 5.x では **CRLF 改行**が安定です（本リポジトリは `.gitattributes` で PS1=CRLF を強制）。

---

## 使い方（macOS / Linux）

1) 権限付与
```bash
chmod +x ./scripts/macos_linux/measure_netquality_with_hops.sh
chmod +x ./scripts/macos_linux/add_teams_slow_mark.command
```

2) **計測**
```bash
./scripts/macos_linux/measure_netquality_with_hops.sh
```
- 出力：`~/.local/state/teamsnet/teams_net_quality.csv` / `path_hop_quality.csv`

3) **「遅い」マーカー（ダブルクリック可）**
- Finder で `add_teams_slow_mark.command` をダブルクリック

---

## CSV ヘッダーの意味

### `teams_net_quality.csv`
- **timestamp**：計測時刻（PC ローカル）
- **host**：計測宛先 FQDN（例: `world.tr.teams.microsoft.com`）
- **icmp_avg_ms**：ICMP ping の平均 RTT（ms）
- **icmp_jitter_ms**：ICMP RTT のばらつき（標準偏差, ms）
- **loss_pct**：ICMP の損失率（%）
- **dns_ms**：名前解決（DNS）の所要時間（ms）
- **tcp_443_ms**：TCP 443 への接続完了（SYN→ESTABLISHED）までの時間（ms）
- **http_head_ms**：HTTPS で HEAD リクエストが返るまで（TTFB 近似, ms）
- **mos_estimate**：簡易 MOS（1.0〜4.5）。`RTT` と `loss_pct` からの概算値
  *目安：4.0 以上 = 良好 / 3.6〜4.0 = 許容 / 3.6 未満 = 悪化傾向*
- **conn_type**：`wifi` or `wired_or_disconnected`
- **ssid**：接続 SSID（Wi-Fi 時）
- **bssid**：AP の BSSID（AP の無線 MAC）
- **signal_pct**：Wi-Fi 受信感度（% *Windows*）。*mac/Linux は `signal`（dBm）*
- **ap_name**：`ap_map.csv` で BSSID→AP 名に引けた場合の名称
- **roamed**：直前サンプルから BSSID が変わったら `roamed`
- **roam_from / roam_to**：ローミング前後の BSSID
- **notes**：補足フラグ（例: `icmp_blocked` / `http_fail` / `tcp_fail` / `roamed` などの合成）

### `path_hop_quality.csv`
- **timestamp**：同一トレース（同一回）のホップで共通の時刻
- **target**：traceroute の宛先 FQDN
- **hop_index**：ホップ番号（1 から順に付与）
  *注：ICMP に応答したホップだけを記録するため、実 TTL と一致しないことがあります。*
- **hop_ip**：そのホップで応答したルータの IPv4 アドレス
- **icmp_avg_ms / icmp_jitter_ms / loss_pct**：そのホップの ICMP 応答の平均 / ジッタ / 損失
- **notes**：ホップ単位の補足
  - `icmp_blocked` … そのホップが ICMP に応答しない / 遮断
  - `tracert_no_reply` … その回の traceroute で IP が 1 つも取れなかった
- **conn_type / ssid / bssid / signal_pct / ap_name / roamed / roam_from / roam_to**：上と同じ（計測時の接続状況を添付）

### `user_marks.csv`（任意の“遅い！”ボタンの記録）
- **timestamp**：押した時刻
- **username / computer**：ログイン名 / 端末名
- **conn_type / ssid / bssid / signal_pct**：押した瞬間の接続状態
- **comment**：利用者が入力したメモ

---

## AP 名マッピング（任意）

`samples/ap_map.csv` を出力先フォルダ（Windows: `%LOCALAPPDATA%\TeamsNet`、macOS/Linux: `~/.local/state/teamsnet`）にコピーし、必要に応じて追記してください。

```
bssid,ap_name
aa:bb:cc:dd:ee:ff,会議室A-天井AP
aa:bb:cc:dd:ee:f0,執務室-西側
```

---

## よくある調整

- **ホップ計測の負荷**：Windows は `-HopProbeEveryCycles`、mac/Linux は `HOP_EVERY` で頻度を上げ下げ
- **1 ホップあたりの回数**：`-HopPingCount` / `HOP_PING`
- **ターゲット FQDN**：`Targets` / `HOP_TARGETS` を運用に合わせて編集
- **起動直後にも hop 計測**：`($cycle % HopProbeEveryCycles) -eq 0 -or $cycle -eq 1` にする

---

## 既知の挙動・注意

- 一部の宛先/ルータは **ICMP に低優先 or 応答しない**ため、`icmp_blocked` やホップ欠落が発生します。前後ホップとの差やエンドツーエンド指標と併せて判断してください。
- CSV を **Excel で開いている間**は追記できないため、`*.csv.queue` に退避します。Excel を閉じると次サイクルで自動合流します。
- PowerShell 5.x では **CRLF 改行**が安定です（本リポジトリは `.gitattributes` で PS1=CRLF / SH=LF を強制）。

---

## ライセンス

MIT License
