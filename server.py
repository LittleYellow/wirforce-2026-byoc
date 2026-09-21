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
import json
import os
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


def log(msg):
    print(f"[{datetime.now(TZ_TAIPEI):%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


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

    def snapshot(self):
        with self._lock:
            return self.asset

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
PAGE = None     # index.html，啟動時讀一次


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

    def do_GET(self):
        path = self.path.split("?", 1)[0].rstrip("/") or "/"
        if path in self.PAGES:
            # 頁面本身很少變，但別讓瀏覽器快取太久，改版才推得動
            return self._send(PAGE, "public, max-age=300, must-revalidate")
        if path == "/seats.json":
            asset = SEATS.snapshot()
            if asset is None:
                return self._json(503, {"error": "座位資料還沒抓到，請稍後再試",
                                        "detail": SEATS.last_error})
            return self._send(asset, "no-cache")
        if path in ("/healthz", "/health"):
            h = SEATS.health()
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

    log(f"啟動：每 {REFRESH_MINUTES:g} 分鐘重抓一次，sheet {SHEET_ID}")
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
