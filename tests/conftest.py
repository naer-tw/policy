"""pytest 共用設定：起本機伺服器供 Playwright 測試根相對路徑（/…）。

跟 aabe-deploy 的 `_source/deploy-automation/tests/conftest.py` 同一套寫法，
但 _policy-deploy 是扁平的 GitHub Pages repo（沒有 public/ 子目錄、也沒有
既有的 deploy-automation 工具鏈），所以獨立成一份最小可用版本。
"""
import functools
import http.server
import threading
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture(scope="session")
def repo_root() -> Path:
    return REPO_ROOT


@pytest.fixture(scope="session")
def sync_playwright():
    return pytest.importorskip(
        "playwright.sync_api", reason="playwright 套件未安裝（見 README 或 aabe-deploy 側的安裝法）"
    ).sync_playwright


@pytest.fixture(scope="session")
def local_site(repo_root: Path):
    handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(repo_root))
    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    port = httpd.server_address[1]
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{port}"
    finally:
        httpd.shutdown()
        thread.join(timeout=5)


@pytest.fixture(scope="session")
def browser(sync_playwright):
    with sync_playwright() as p:
        try:
            b = p.chromium.launch()
        except Exception as e:  # pragma: no cover
            pytest.skip(f"無法啟動 Chromium（{e}）")
        yield b
        b.close()


@pytest.fixture()
def page(browser):
    ctx = browser.new_context()
    pg = ctx.new_page()
    yield pg
    ctx.close()
