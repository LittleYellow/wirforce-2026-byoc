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

## 部署到 GitHub Pages（自動更新）
1. 建一個 public repo，放入 `index.html`、`build_seats.py`、`requirements.txt`，
   並把 `update-seats.yml` 放到 **`.github/workflows/update-seats.yml`**（目前它在專案根目錄，要自己移過去）
2. Settings → Pages → Source 選 **GitHub Actions**
3. Settings → Secrets and variables → Actions → Variables 新增 `SHEET_ID`
4. Actions 分頁手動跑一次「更新座位資料」，之後每 10 分鐘自動更新

`seats.json` 由 workflow 每次重新產生後直接部署，**不用 commit 進 repo**（`.gitignore` 已經排除）。

## 頁面功能
- 暱稱／座位編號查詢，結果附同桌、對面、背後、隔壁（打 a24、A024、A0024 都找得到 A0024）
- 空白領域營區可以用團名搜尋（例如「放影旅團」）
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
