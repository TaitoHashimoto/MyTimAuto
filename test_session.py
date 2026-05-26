"""セッション確認テスト（打刻しない）"""
import asyncio
from pathlib import Path
from playwright.async_api import async_playwright

BASE_DIR = Path(__file__).parent
STORAGE_FILE = BASE_DIR / "storage.json"

try:
    from config import MYTIM_URL, MYTIM_HOST
except ImportError as _e:
    raise SystemExit(
        "config.py が見つかりません。config.example.py を config.py にコピーして値を設定してください。"
    ) from _e

async def test():
    if not STORAGE_FILE.exists():
        print("[ERROR] storage.json が見つかりません。setup.py を先に実行してください。")
        return

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=[
                "--use-system-default-certificate-store",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        ctx = await browser.new_context(storage_state=str(STORAGE_FILE))
        page = await ctx.new_page()

        print("接続中...")
        await page.goto(MYTIM_URL, timeout=60_000)
        await page.wait_for_load_state("networkidle", timeout=20_000)

        from urllib.parse import urlparse
        current_host = urlparse(page.url).netloc
        print(f"URL: {page.url}")
        print(f"ホスト: {current_host}")

        if current_host == MYTIM_HOST:
            print("\n[OK] セッション有効！ MyTimに接続済み")
        else:
            print("\n[NG] セッション無効。setup.py を再実行してください")

        # ボタン検索
        print("\nボタン検索中...")
        for name in ["出勤", "休憩", "退勤", "Clock In", "Clock Out", "Break"]:
            for selector in [
                f"button:has-text('{name}')",
                f"[role='button']:has-text('{name}')",
                f"a:has-text('{name}')",
            ]:
                el = page.locator(selector).first
                try:
                    if await el.is_visible(timeout=1_500):
                        print(f"  見つかった: '{name}' ({selector})")
                        break
                except Exception:
                    continue

        screenshot_path = BASE_DIR / "screenshot.png"
        await page.screenshot(path=str(screenshot_path))
        print(f"\nスクリーンショット: {screenshot_path}")

        await ctx.close()
        await browser.close()

asyncio.run(test())
