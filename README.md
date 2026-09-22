# WirForce 2026 BYOC 座位查詢

社群座位表（Google Sheets）→ 解析成 `seats.json` → 靜態查詢頁。

## 需要什麼

| | 版本 | 說明 |
|---|---|---|
| Python | 3.9 以上（CI 用 3.12） | `venv`、`http.server` 都是內建的，不用另外裝 |
| openpyxl | 3.1.5 | **唯一的相依套件**，見 `requirements.txt`。只有 xlsx 保留合併儲存格，大桌判斷靠它 |

不需要 Node.js、不需要打包工具、不需要後端。

## 初始化（第一次才要做）

在專案資料夾裡，照自己的系統選一段整個貼上。

### Windows — PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -c "import openpyxl; print(openpyxl.__version__)"
```

### macOS / Linux — bash、zsh

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -c "import openpyxl; print(openpyxl.__version__)"
```

最後一行印出 `3.1.5` 就成功了。

> **PowerShell 兩個坑**
> 1. Windows PowerShell 5.1 **不支援 `&&`**，要把指令串成一行請改用 `;`。
> 2. 被擋下來出現「因為這個系統上已停用指令碼執行」時，先跑
>    `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`（只影響這個視窗），
>    或乾脆不啟用、直接用 `.\.venv\Scripts\python.exe` 代替所有的 `python`。

> **macOS 一個坑**：系統內建的是 `python3`，沒有 `python`。建立 venv 要用 `python3`，
> 但**啟用之後** `python` 就會指向 venv 裡的 3.x，所以後面都寫 `python`。

> **在 Windows 上用 Git Bash / WSL 的話**：指令照 macOS 那段抄，但**啟用路徑不一樣** ——
> Windows 的 venv 產生的是 `Scripts/` 不是 `bin/`，要用 `source .venv/Scripts/activate`。

用 [uv](https://docs.astral.sh/uv/) 的話，上面整段等同於：

```powershell
uv venv; uv pip install -r requirements.txt      # PowerShell
```
```bash
uv venv && uv pip install -r requirements.txt    # macOS / Linux
```

## 跑起來

**每次開新的終端機都要先啟用虛擬環境。** 下面兩段各自從啟用開始，整段貼上即可。

### Windows — PowerShell

```powershell
.\.venv\Scripts\Activate.ps1
python build_seats.py --sheet-id 13n2z8QgxQL-xWLR2K4zy75Xw1aeu6Grf-DuGvZBckXY
python -m http.server 8000
```

### macOS / Linux — bash、zsh

```bash
source .venv/bin/activate
python build_seats.py --sheet-id 13n2z8QgxQL-xWLR2K4zy75Xw1aeu6Grf-DuGvZBckXY
python -m http.server 8000
```

然後開 <http://localhost:8000>。換埠號就改 `8000`，`Ctrl+C` 停掉伺服器。

三行分別在做：

1. **啟用虛擬環境** —— 這是唯一兩個系統不一樣的指令
2. **產生 `seats.json`** —— 直接抓線上版（試算表需開放「知道連結的人可檢視」）。
   跑完會印出每一區的座位數與缺號，有疑慮的地方會印 ⚠ 警告，請對照原表確認
3. **起伺服器** —— 只是 Python 內建的靜態檔案伺服器，兩個系統指令相同

想改用手邊的 xlsx（Google Sheets →「檔案 → 下載 → Microsoft Excel (.xlsx)」），
把第 2 行換成（兩個系統一樣）：

```bash
python build_seats.py 座位表.xlsx
```

> **不要用檔案總管／Finder 直接點開 `index.html`。** 瀏覽器不允許「直接點開的網頁」讀取同資料夾的
> `seats.json`，會退回內建範例資料（只有 A 區 48 格、沒有空白領域）。一定要走上面的
> `http://localhost:8000`。頁面偵測到這個情況時會直接把原因寫在畫面上。

## 部署（Zeabur）

用 `server.py` 跑一個小服務：背景**每 30 分鐘**重抓一次座位表解析成 `seats.json` 放在記憶體，
同時把 `index.html` 和 `seats.json` 送出去。`Dockerfile` 已經寫好，Zeabur 直接吃。

1. Zeabur 新增 Service → Git → 選這個 repo（會自動偵測到 `Dockerfile`）
2. 環境變數只要一個：

   | 變數 | 值 |
   |---|---|
   | `SHEET_ID` | `13n2z8QgxQL-xWLR2K4zy75Xw1aeu6Grf-DuGvZBckXY` |
   | `REFRESH_MINUTES` | `30`（可省略，預設就是 30） |
   | `DATA_DIR` | `/data`（可省略，預設就是 `/data`） |

   `PORT` 由 Zeabur 自己帶入，不用設。
3. 綁網域，開 `/` 就是查詢頁

健康檢查可以指到 `/healthz`，回傳目前座位數、上次抓取時間、失敗次數與解析警告：

```json
{"ok":true,"seats":1012,"occupied":243,"zones":14,"age_seconds":14,"failures":0,"warnings":[...],
 "data":{"path":"/data","exists":true,"writable":true,"persisted":true,"boots":7,
         "first_boot":"2026-09-22T09:00:00+08:00","free_mb":1024,
         "entries":[".volume-probe.json"],"error":null}}
```

`data` 是持久硬碟的狀況：

- `exists` 幾乎一定是 `true`（目錄不在的話這支程式自己會建），**不要拿它判斷有沒有掛到**
- `writable` 是實際寫一個檔案再讀回來的結果，掛載點唯讀或 owner 不對就會是 `false`
- **`persisted` 才是答案**：標記檔在這個行程啟動前就存在，代表資料活過了上一次重啟。
  沒掛 Volume 的話每次部署都是全新容器，這裡會是 `false`
- `boots` 累計啟動次數，數字一直往上加就是硬碟持續留著東西

在 Zeabur 掛硬碟：服務頁 →「硬碟」→「掛載硬碟」，Volume ID 隨便取（例如 `data`），
**掛載目錄填 `/data`**。注意掛載會把該目錄整個清空，所以絕對不要填 `/app`。
啟用 Volume 之後就沒有零停機重啟了，每次部署會完整關閉再啟動。

**抓取失敗時會繼續送上一份成功的資料**，不會讓現場的人看到空頁面；失敗後改成每分鐘重試，
成功才回到 30 分鐘。`seats.json` 只存在記憶體，不寫磁碟，所以容器檔案系統唯讀也沒問題。

本機要跑這個服務：

```powershell
$env:SHEET_ID = "13n2z8QgxQL-xWLR2K4zy75Xw1aeu6Grf-DuGvZBckXY"
python server.py
```
```bash
SHEET_ID=13n2z8QgxQL-xWLR2K4zy75Xw1aeu6Grf-DuGvZBckXY python server.py
```

> `.github/workflows/update-seats.yml`（GitHub Pages 那套）**目前是停用狀態**，先留著當備案。
> 要切回去就 `gh workflow enable "更新座位資料"`，並到 Settings → Pages 把 Source 設成 GitHub Actions。

## 頁面

| path | 內容 | 資料來源 |
|---|---|---|
| `/` | 座位查詢 + 全區地圖 | 地圖工作表（目前「世貿場地」） |
| `/map` | 同上，直接展開地圖 | 同上 |
| `/seat/A0024` | 單一座位的頁面：介紹與活動、同桌／對面／背後／隔壁、整張桌子 | 地圖工作表 + 備註欄 |
| `/zone/01` | 單一空白領域營區的頁面：介紹與活動、在地圖上的位置 | 「空白領域」工作表 |
| `/zones` | 空白領域 14 個營區與介紹 | 「空白領域」工作表（介紹寫在團名那格下方） |
| `/board` | 留言板（**導覽列先不顯示**，網址仍可直接開） | 有「編號／ID／備註」三欄那張的備註欄 |
| `/food` | 美食與設施（**導覽列先不顯示**，網址仍可直接開） | 場地圖上的招牌 |

路由在前端做，後端把這些 path 都送同一份 `index.html`。座位頁也收手打的 `/seat/a24`
（前端會把網址正規化成 `/seat/A0024`）；不是「1 英文 + 1～4 數字」的一律 404，不要讓任何路徑都回 200。

## 頁面功能
- 暱稱／座位編號查詢，結果附同桌、對面、背後、隔壁（打 a24、A024、A0024 都找得到 A0024）
- 每個座位都有自己的網址 `/seat/A0024`，可以直接分享；頁上有「介紹與活動」（原表備註欄）、
  「分享我的座位」（手機會叫出系統分享選單）與「複製這頁連結」。
  大桌兩個編號是同一個人，任一邊的備註兩頁都看得到
- 連結貼到 Discord／LINE 會有預覽（暱稱、排數、說明），由後端產生。
  預覽只放**這個座位自己的**暱稱；對面、隔壁、背後是誰要點進來才看得到，
  免得貼一個連結就順便把鄰居的暱稱攤在聊天室裡
- 空白領域營區可以用團名搜尋（例如「放影旅團」），也有自己的網址 `/zone/01` 可以分享；
  介紹與活動寫在原表「空白領域」分頁、團名正下方那個合併的大格子，換行用 Alt+Enter
- 「看全區地圖」：A/B/C 三區的相對位置圖，格線沿用原表的列欄，障礙物缺口與走道照原表留白
  - 只顯示「有人／空位」和原表底色，暱稱要點下去才出現
  - 三段縮放：全區（整區塞進畫面，點一下先放大避免誤觸）／中／大
  - `A 區／B 區／C 區／空白領域` 快速跳轉；放大後只看得到一小塊，靠這個導航
  - 不是按實際場地比例畫的，也沒有方位

## 解析規則
- 自動挑座位最多的那張工作表（目前是「世貿場地」），不寫死名稱
- 座位編號：原表寫 `A01`、`B123`，輸出一律正規化成 1 英文 + 4 數字（`A0001`、`C0466`）；同一列至少 3 個才算座位列（避免抓到「傳送門」說明文字）
- 空白領域：14 個營區另外放在 `zones`，沒有座位編號，不混進 `seats`
- 同一列會橫跨 A/B/C 三區，先切成「欄位相鄰、同一區」的區塊再逐塊判斷方向
- 方向看編號的蛇行順序（由小到大＝面上、由大到小＝面下），再用原表的「面上／面下」標記和暱稱位置交叉驗證
- 暱稱合併儲存格跨兩三個座位 = 大桌；跨更多的是走道說明，不算暱稱
- 解析有疑慮時會印出 ⚠ 警告，請對照原表確認
