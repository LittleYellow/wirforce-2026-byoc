#!/usr/bin/env python3
"""
座位查詢的小服務（給 Zeabur 用）。

做兩件事：
  1. 背景每 REFRESH_MINUTES 分鐘重抓一次 Google Sheets，解析成 seats.json 放在記憶體裡
  2. 把 index.html 與 seats.json 送出去

刻意不寫進磁碟：容器的檔案系統可能是唯讀的，而且資料本來就只活在記憶體裡就夠了。
抓失敗時**繼續送上一份成功的資料**，不會把現場的人晾在沒資料的頁面。

環境變數：
  SHEET_ID          必填，Google Sheets 的 ID
  PORT              預設 8080（Zeabur 會自己給）
  REFRESH_MINUTES   預設 30
"""
import gzip
import html
import json
import os
import re
import shutil
import sys
import threading
import time
import traceback
from datetime import datetime, timezone, timedelta
from hashlib import md5
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from build_seats import download_xlsx, fetch_gids, parse

SHEET_ID = os.environ.get("SHEET_ID", "").strip()
PORT = int(os.environ.get("PORT", "8080"))
REFRESH_MINUTES = float(os.environ.get("REFRESH_MINUTES", "30"))
RETRY_SECONDS = 60          # 抓失敗就一分鐘後再試，不用等滿 30 分鐘
TZ_TAIPEI = timezone(timedelta(hours=8))
ROOT = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", "/data"))   # Zeabur 掛上來的持久硬碟
BOOT_MARKER = DATA_DIR / ".volume-probe.json"
BOOT_LOCK = threading.Lock()
BOOT_COUNTED = [False]        # 這個行程有沒有算過啟動次數
BOOT_PERSISTED = [False]


def log(msg):
    print(f"[{datetime.now(TZ_TAIPEI):%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def probe_data_dir() -> dict:
    """真的去寫一個檔案，確認持久硬碟可用。

    光檢查目錄存不存在不夠，而且這支程式自己會把目錄建出來，所以 exists 幾乎一定是 True——
    掛載點有可能是唯讀的，或 owner 不是跑這支程式的使用者，那要 writable 才看得出來。

    persisted 才是「Volume 有沒有真的掛上」的答案：標記檔在這個行程啟動前就存在，
    表示它活過了上一次重啟。沒掛 Volume 的話每次部署都是全新的容器，標記檔不會在。
    boots 是累計啟動次數，數字一直加代表硬碟持續留著東西。
    """
    info = {"path": str(DATA_DIR), "exists": False, "writable": False, "persisted": False,
            "boots": None, "first_boot": None, "free_mb": None, "entries": None, "error": None}
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        info["exists"] = True
        probe = DATA_DIR / ".write-probe"
        stamp = datetime.now(TZ_TAIPEI).isoformat(timespec="seconds")
        probe.write_text(stamp, encoding="utf-8")
        if probe.read_text(encoding="utf-8") != stamp:
            raise OSError("寫進去再讀出來對不起來")
        probe.unlink()
        info["writable"] = True

        with BOOT_LOCK:
            state = {}
            if BOOT_MARKER.exists():
                try:
                    state = json.loads(BOOT_MARKER.read_text(encoding="utf-8"))
                except ValueError:
                    state = {}                       # 檔案壞了就當第一次，不要讓健康檢查掛掉
            if not state.get("first_boot"):
                state = {"first_boot": stamp, "boots": 0}
            if not BOOT_COUNTED[0]:                  # 一個行程只算一次，healthz 被打幾次都一樣
                state["boots"] = state.get("boots", 0) + 1
                BOOT_COUNTED[0] = True
                BOOT_PERSISTED[0] = state["first_boot"] != stamp
                BOOT_MARKER.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
            info["first_boot"] = state["first_boot"]
            info["boots"] = state["boots"]
            info["persisted"] = BOOT_PERSISTED[0]

        info["free_mb"] = round(shutil.disk_usage(DATA_DIR).free / 1048576)
        info["entries"] = sorted(p.name for p in DATA_DIR.iterdir())
    except Exception as e:
        info["error"] = f"{type(e).__name__}: {e}"
    return info


class Asset:
    """一份要送出去的東西：原始 bytes、壓縮過的 bytes、ETag。"""

    def __init__(self, body: bytes, content_type: str):
        self.body = body
        self.gzipped = gzip.compress(body, 6)
        self.etag = '"' + md5(body).hexdigest()[:16] + '"'
        self.content_type = content_type


class Seats:
    """最新一份座位資料。抓失敗時保留上一份。"""

    def __init__(self):
        self._lock = threading.Lock()
        self.asset = None
        self.fetched_at = None      # 上一次「成功」的時間
        self.generated_at = None
        self.total = self.occupied = self.zones = 0
        self.warnings = []
        self.last_error = None
        self.attempts = 0
        self.failures = 0
        self.by_id = {}
        self.generation = 0        # 資料換一次就加一，座位頁的快取靠它作廢

    def update(self, data: dict):
        body = json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode()
        asset = Asset(body, "application/json; charset=utf-8")
        with self._lock:
            self.asset = asset
            self.fetched_at = datetime.now(TZ_TAIPEI)
            self.generated_at = data.get("generated_at")
            self.total = data.get("total", 0)
            self.occupied = data.get("occupied", 0)
            self.zones = len(data.get("zones") or [])
            self.warnings = data.get("warnings") or []
            self.last_error = None
            # 座位頁要即時組社群預覽，另外留一份查得動的索引
            self.by_id = {x["id"]: x for x in data.get("seats") or []}
            self.generation += 1

    def snapshot(self):
        with self._lock:
            return self.asset

    def seat(self, sid):
        """回傳 (座位, 版號)。拿著版號才知道座位頁的快取還算不算數。"""
        with self._lock:
            return self.by_id.get(sid), self.generation

    def health(self):
        with self._lock:
            age = (datetime.now(TZ_TAIPEI) - self.fetched_at).total_seconds() if self.fetched_at else None
            return {
                "ok": self.asset is not None,
                "seats": self.total,
                "occupied": self.occupied,
                "zones": self.zones,
                "generated_at": self.generated_at,
                "age_seconds": round(age) if age is not None else None,
                "refresh_minutes": REFRESH_MINUTES,
                "attempts": self.attempts,
                "failures": self.failures,
                "warnings": self.warnings,
                "last_error": self.last_error,
            }


SEATS = Seats()
PAGE = None                    # index.html 原版，啟動時讀一次
PAGE_HEAD = PAGE_TAIL = None   # 以 <!--meta--> 為界拆兩半，座位頁換掉中間那段
META_OPEN, META_CLOSE = "<!--meta-->", "<!--/meta-->"
SEAT_PATH_RE = re.compile("^/seat/([A-Za-z])0*([0-9]{1,4})$")
SEAT_PAGES = {}                # (座位編號, 網域) -> Asset
SEAT_PAGES_GEN = -1            # 這批快取是用哪個版本的座位資料做的
SEAT_PAGES_LOCK = threading.Lock()
SITE_NAME = "WirForce 2026 BYOC 座位查詢"


def seat_meta(s: dict, url: str) -> str:
    """一個座位的社群預覽。

    貼連結到 Discord／LINE／Threads 時，對方的伺服器只抓 HTML、不執行 JavaScript，
    所以這段一定要在後端就寫好——前端 route() 設的標題它們看不到。

    **刻意不放暱稱**（見 AGENTS.md）：預覽會出現在聊天室裡，貼一個連結就把人的暱稱
    攤開來不是玩家自己同意的事。要改這條先問過 Yellow。
    """
    e = lambda x: html.escape(str(x), quote=True)
    title = "{}　{} 區第 {} 排".format(s["id"], s["area"], s["row_no"])
    bits = [title]
    facing = (s.get("facing") or "").split("（")[0].strip()
    if facing:
        bits.append(facing)
    if s.get("big"):
        bits.append("大桌")
    desc = "，".join(bits) + "。這一頁有同桌、對面、背後、隔壁，以及整張桌子的位置圖。"
    tags = [
        "<title>{}｜{}</title>".format(e(title), e(SITE_NAME)),
        '<meta name="description" content="{}">'.format(e(desc)),
        '<meta property="og:type" content="website">',
        '<meta property="og:site_name" content="{}">'.format(e(SITE_NAME)),
        '<meta property="og:title" content="{}">'.format(e(title)),
        '<meta property="og:description" content="{}">'.format(e(desc)),
        '<meta property="og:url" content="{}">'.format(e(url)),
        '<meta name="twitter:card" content="summary">',
    ]
    return "\n".join(tags) + "\n"


def seat_page(sid: str, base: str):
    """座位頁：內容跟首頁同一份，只有 <head> 那段社群預覽不同。找不到座位回 None。"""
    global SEAT_PAGES_GEN
    s, gen = SEATS.seat(sid)
    if s is None:
        return None
    key = (sid, base)
    with SEAT_PAGES_LOCK:
        if SEAT_PAGES_GEN != gen:      # 座位表更新了，整批重做
            SEAT_PAGES.clear()
            SEAT_PAGES_GEN = gen
        hit = SEAT_PAGES.get(key)
    if hit:
        return hit
    asset = Asset((PAGE_HEAD + seat_meta(s, base + "/seat/" + sid) + PAGE_TAIL).encode(),
                  "text/html; charset=utf-8")
    with SEAT_PAGES_LOCK:
        if SEAT_PAGES_GEN == gen:      # 組的過程中又更新了就不要存，下次重做
            SEAT_PAGES[key] = asset
    return asset


def fetch_once() -> bool:
    """抓一次並更新 SEATS。回傳是否成功。"""
    SEATS.attempts += 1
    path = None
    try:
        path = download_xlsx(SHEET_ID)
        data = parse(path, fetch_gids(SHEET_ID))
        SEATS.update(data)
        log(f"✔ {data['total']} 個座位、{data['occupied']} 個已登記、"
            f"{len(data.get('zones') or [])} 個營區（工作表「{data['sheet']}」）")
        for w in data.get("warnings") or []:
            log(f"  ⚠ {w}")
        return True
    except Exception as e:
        SEATS.failures += 1
        SEATS.last_error = f"{type(e).__name__}: {e}"
        kept = "，繼續沿用上一份資料" if SEATS.asset else "，目前還沒有任何資料可以送"
        log(f"✘ 抓取失敗（{SEATS.last_error}）{kept}")
        traceback.print_exc()
        return False
    finally:
        if path:
            try:
                os.unlink(path)          # download_xlsx 會留下暫存檔，長跑的服務要自己清
            except OSError:
                pass


def refresher(first_ok: bool):
    """啟動時已經抓過一次了，所以這裡先睡再抓。"""
    delay = REFRESH_MINUTES * 60 if first_ok else RETRY_SECONDS
    while True:
        time.sleep(delay)
        delay = REFRESH_MINUTES * 60 if fetch_once() else RETRY_SECONDS


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "wirforce-seat-finder"

    def log_message(self, fmt, *args):
        pass                              # 預設每個請求印一行太吵，只留下面自己印的

    def _send(self, asset: Asset, cache: str):
        if self.headers.get("If-None-Match") == asset.etag:
            self.send_response(304)
            self.send_header("ETag", asset.etag)
            self.send_header("Cache-Control", cache)
            self.end_headers()
            return
        gz = "gzip" in (self.headers.get("Accept-Encoding") or "")
        body = asset.gzipped if gz else asset.body
        self.send_response(200)
        self.send_header("Content-Type", asset.content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("ETag", asset.etag)
        self.send_header("Cache-Control", cache)
        if gz:
            self.send_header("Content-Encoding", "gzip")
        self.send_header("Vary", "Accept-Encoding")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _json(self, code: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def do_HEAD(self):
        self.do_GET()

    # 前端自己做路由，這些 path 一律送同一份 index.html
    PAGES = ("/", "/index.html", "/zones", "/board", "/food", "/map")

    def _base_url(self) -> str:
        """組 og:url 用。Zeabur 在前面擋了一層，真正的協定與網域在 X-Forwarded-* 裡。"""
        proto = (self.headers.get("X-Forwarded-Proto") or "https").split(",")[0].strip()
        host = (self.headers.get("X-Forwarded-Host") or self.headers.get("Host") or "").split(",")[0].strip()
        return proto + "://" + host

    def do_GET(self):
        path = self.path.split("?", 1)[0].rstrip("/") or "/"
        if path in self.PAGES:
            # 頁面本身很少變，但別讓瀏覽器快取太久，改版才推得動
            return self._send(PAGE, "public, max-age=300, must-revalidate")
        m = SEAT_PATH_RE.match(path)
        if m:
            # 原表沒有的編號（打錯字，或號碼被障礙物佔掉）就送原版頁面，
            # 前端會顯示「找不到這個座位」——比直接 404 好，使用者還能繼續查別的
            sid = m.group(1).upper() + m.group(2).zfill(4)
            page = seat_page(sid, self._base_url()) or PAGE
            return self._send(page, "public, max-age=300, must-revalidate")
        if path == "/seats.json":
            asset = SEATS.snapshot()
            if asset is None:
                return self._json(503, {"error": "座位資料還沒抓到，請稍後再試",
                                        "detail": SEATS.last_error})
            return self._send(asset, "no-cache")
        if path in ("/healthz", "/health"):
            h = SEATS.health()
            h["data"] = probe_data_dir()
            return self._json(200 if h["ok"] else 503, h)
        self._json(404, {"error": "not found"})


def main():
    global PAGE
    if not SHEET_ID:
        sys.exit("請設定環境變數 SHEET_ID（Google Sheets 的 ID）")
    page = ROOT / "index.html"
    if not page.exists():
        sys.exit(f"找不到 {page}")
    PAGE = Asset(page.read_bytes(), "text/html; charset=utf-8")
    text = page.read_text(encoding="utf-8")
    if META_OPEN not in text or META_CLOSE not in text:
        sys.exit("index.html 裡找不到 " + META_OPEN + " ... " + META_CLOSE
                 + " 這段，座位頁沒辦法換社群預覽")
    global PAGE_HEAD, PAGE_TAIL
    PAGE_HEAD = text.split(META_OPEN, 1)[0]
    PAGE_TAIL = text.split(META_CLOSE, 1)[1]

    log(f"啟動：每 {REFRESH_MINUTES:g} 分鐘重抓一次，sheet {SHEET_ID}")
    d = probe_data_dir()
    if d["writable"]:
        kept = (f"資料留得住（第 {d['boots']} 次啟動，最早 {d['first_boot']}）"
                if d["persisted"] else "⚠ 但這是全新的目錄，可能根本沒掛到 Volume")
        log(f"持久硬碟 {d['path']}：可寫，剩 {d['free_mb']} MB，{kept}")
    else:
        log(f"⚠ 持久硬碟 {d['path']} 不可用（{d['error']}），"
            f"寫入類功能會失效 —— 檢查 Zeabur 的硬碟掛載設定")
    ok = fetch_once()                     # 先抓一次再開始服務，避免第一個進來的人看到範例資料
    threading.Thread(target=refresher, args=(ok,), daemon=True).start()

    server = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    server.daemon_threads = True
    log(f"開始監聽 0.0.0.0:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        log("收到中斷，關閉中")
        server.shutdown()


if __name__ == "__main__":
    main()
