# teams-net-quality-tools

**Teams の体感遅延の見える化**のためのスクリプト集です。管理者権限なしで動作し、以下を記録します。

- エンドツーエンド品質（RTT/ジッタ/ロス、DNS/TCP/HTTP、簡易MOS）
- 接続AP（SSID/BSSID/信号値）
- ローミング検知（BSSIDの変化）
- 途中経路（各ホップ）のジッタ/ロス（指定サイクルごと）
- ユーザーの「今遅い！」ワンクリック時刻メモ

出力は CSV：
- `teams_net_quality.csv` … エンドツーエンド品質＋AP/ローミング
- `path_hop_quality.csv` … 途中経路ホップごとの品質
- `user_marks.csv` … ユーザーの手動マーカー（任意）

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

`ap_map.csv` は **BSSID → AP名** の対応表（任意）。

---

## 使い方（Windows）

1) **計測（権限不要）**
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\windows\Measure-NetQuality-WithHops.ps1" -IntervalSeconds 30 -HopProbeEveryCycles 10
```
- 出力：`%LOCALAPPDATA%\TeamsNet\teams_net_quality.csv` / `path_hop_quality.csv`
- 途中経路計測の頻度は `-HopProbeEveryCycles` で調整

2) **「遅い」マーカー（任意）**
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\windows\Add-TeamsSlowMark.ps1"
```
- 出力：`%LOCALAPPDATA%\TeamsNet\user_marks.csv`

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
- Finderで `add_teams_slow_mark.command` をダブルクリック

---

## AP名マッピング（任意）

`samples/ap_map.csv` をそれぞれの出力先フォルダ（Windows: `%LOCALAPPDATA%\TeamsNet`、macOS/Linux: `~/.local/state/teamsnet`）にコピーし、必要に応じて追記してください。

```
bssid,ap_name
aa:bb:cc:dd:ee:ff,会議室A-天井AP
aa:bb:cc:dd:ee:f0,執務室-西側
```

---

## よくある調整

- **ホップ計測の負荷**：Windowsは `-HopProbeEveryCycles`、mac/Linuxは `HOP_EVERY` で頻度を上げ下げ。
- **1ホップあたりの回数**：`-HopPingCount` / `HOP_PING`。
- **ターゲットFQDN**：`Targets` / `HOP_TARGETS` を実運用に合わせて編集。

> 注：一部のルータは ICMP を制限します。そのホップは `icmp_blocked` と記録されます。前後ホップの差分やエンドツーエンド指標と合わせて読み解いてください。

---

## GitHub への公開手順（例）

1) GitHub で新規リポジトリ（例：`teams-net-quality-tools`）を作成
2) ローカルで：
```bash
cd teams-net-quality-tools
git init
git add .
git commit -m "Initial commit: network quality scripts with AP & hop jitter"
git branch -M main
git remote add origin https://github.com/<YOUR_USER>/teams-net-quality-tools.git
git push -u origin main
```
- Windows の場合は PowerShell でも同様です。

---

## ライセンス

MIT License
