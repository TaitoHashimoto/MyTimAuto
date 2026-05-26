"""
Edge + storage.json でTeamsメッセージが読めるかテスト
"""
import asyncio
import re
from pathlib import Path
from playwright.async_api import async_playwright

BASE_DIR = Path(__file__).parent
SESSION_DIR = BASE_DIR / "session"

try:
    from config import TEAMS_CHAT_ID
except ImportError as _e:
    raise SystemExit(
        "config.py が見つかりません。config.example.py を config.py にコピーして値を設定してください。"
    ) from _e

TEAMS_URL = f"https://teams.microsoft.com/v2/#/chat/{TEAMS_CHAT_ID}"

MSG_BLOCK_PATTERN = re.compile(
    r"日付[：:].*?当月休暇実績時間[：:][\d.]+\s*h?",
    re.DOTALL,
)


async def test():
    async with async_playwright() as p:
        print("Edge + 永続プロファイルで起動中...")
        ctx = await p.chromium.launch_persistent_context(
            str(SESSION_DIR),
            channel="msedge",
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
        )
        page = await ctx.new_page()

        print(f"Teams接続中: {TEAMS_URL}")
        await page.goto(TEAMS_URL, timeout=60_000)
        await page.wait_for_load_state("domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(3_000)

        current_url = page.url
        print(f"現在のURL: {current_url[:100]}")

        if "login.microsoftonline" in current_url:
            print("NG: ログインページにリダイレクトされました（セッション無効）")
            await browser.close()
            return

        if "Unsecured" in await page.inner_text("body") or "Non Compliant" in await page.inner_text("body"):
            print("NG: デバイスコンプライアンスエラー")
            await browser.close()
            return

        # Teams SPA が描画されるまで待機（最大90秒）
        print("メッセージ描画を待機中（最大90秒）...")
        try:
            await page.wait_for_function(
                "() => document.body.innerText.length > 200",
                timeout=90_000,
            )
        except Exception:
            print("描画タイムアウト")

        await page.wait_for_timeout(3_000)
        body = await page.inner_text("body")
        await page.screenshot(path=str(BASE_DIR / "test_edge.png"), full_page=True)

        print(f"取得文字数: {len(body)}")
        print(f"本文先頭200文字: {body[:200]}")

        blocks = MSG_BLOCK_PATTERN.findall(body)
        if blocks:
            print(f"\nOK: 残業メッセージ検出（{len(blocks)}件）")
            print(f"最新メッセージ: {blocks[-1][:80]}")
        else:
            print("\n残業メッセージは未検出（Teamsは読めているが該当メッセージなし）")

        await ctx.close()

asyncio.run(test())
