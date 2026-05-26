"""
MyTim 自動打刻スクリプト
使い方: python punch.py [出勤|休憩|退勤|1|2|3]
"""
import sys
import asyncio
import subprocess
import traceback
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urlparse
from playwright.async_api import async_playwright

# ログ等の日付・時刻はJST固定（PCタイムゾーン設定に依存しない）
JST = timezone(timedelta(hours=9), name="JST")

BASE_DIR = Path(__file__).parent
LOG_FILE = BASE_DIR / "punch.log"
SESSION_DIR = BASE_DIR / "session"
STORAGE_FILE = BASE_DIR / "storage.json"

# 個人/組織固有の値は config.py から読み込む
try:
    from config import MYTIM_URL, MYTIM_HOST
except ImportError as _e:
    raise SystemExit(
        "config.py が見つかりません。"
        "config.example.py を config.py にコピーして値を設定してください。"
    ) from _e

BUTTONS = {
    "出勤": ["勤務開始", "出勤", "出社", "Clock In"],
    "休憩": ["休憩開始", "休憩", "Break"],
    "退勤": ["勤務終了", "退勤", "退社", "Clock Out"],
}

# VBSから数字で受け取る（日本語文字化け回避）
ACTION_MAP = {"1": "出勤", "2": "休憩", "3": "退勤"}

# 状態ごとの操作許可ルール
ALLOWED_STATES = {
    "出勤": ["未出勤", "休憩中"],
    "休憩": ["勤務中"],
    "退勤": ["勤務中"],
}

# 許可されない場合の通知メッセージ
BLOCKED_MESSAGES = {
    ("出勤", "勤務中"): ("出勤できません", "現在勤務中です"),
    ("休憩", "未出勤"): ("休憩できません", "まだ出勤していません"),
    ("休憩", "休憩中"): ("休憩できません", "現在休憩中です"),
    ("退勤", "未出勤"): ("退勤できません", "まだ出勤していません"),
    ("退勤", "休憩中"): ("退勤できません", "現在休憩中です\n先に休憩を終了してください"),
}

# 状態の表示名
STATE_LABEL = {
    "未出勤": "未出勤",
    "勤務中": "勤務中",
    "休憩中": "休憩中",
}


def log(msg: str):
    timestamp = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def on_mytim(url: str) -> bool:
    try:
        return urlparse(url).netloc == MYTIM_HOST
    except Exception:
        return False


def notify(title: str, msg: str):
    ps = f"""
Add-Type -AssemblyName System.Windows.Forms
$n = New-Object System.Windows.Forms.NotifyIcon
$n.Icon = [System.Drawing.SystemIcons]::Application
$n.Visible = $true
$n.BalloonTipTitle = "{title}"
$n.BalloonTipText = "{msg}"
$n.ShowBalloonTip(5000)
Start-Sleep -Seconds 6
$n.Visible = $false
$n.Dispose()
"""
    subprocess.Popen(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
        creationflags=0x08000000,
    )


async def get_state(page) -> str:
    """ページのテキストから現在の打刻状態を判定する"""
    try:
        body = await page.inner_text("body")
        if "休憩中" in body:
            return "休憩中"
        if "勤務中" in body:
            return "勤務中"
    except Exception:
        pass
    return "未出勤"


async def punch_on_page(page, action: str) -> bool:
    # 状態チェック
    state = await get_state(page)
    log(f"現在の状態: {state}")

    if state not in ALLOWED_STATES[action]:
        title, msg = BLOCKED_MESSAGES[(action, state)]
        notify(f"MyTim ⚠ {title}", f"現在の状態: {STATE_LABEL[state]}\n{msg}")
        log(f"[SKIP] {action} は {state} のため押下不可")
        return False

    # ボタンクリック
    for name in BUTTONS[action]:
        for selector in [
            f"button:has-text('{name}')",
            f"[role='button']:has-text('{name}')",
            f"a:has-text('{name}')",
            f"input[value='{name}']",
            f"span:has-text('{name}')",
        ]:
            try:
                el = page.locator(selector).first
                if await el.is_visible(timeout=5_000):
                    await el.click()
                    await page.wait_for_timeout(2_000)
                    notify("MyTim 打刻完了 ✓", f"{action}の打刻が完了しました")
                    log(f"[OK] {action}の打刻完了")
                    return True
            except Exception:
                continue

    screenshot_path = BASE_DIR / f"error_{action}.png"
    await page.screenshot(path=str(screenshot_path))
    notify("MyTim エラー ⚠", f"{action}ボタンが見つかりませんでした\n打刻は完了していません\nMyTimを手動で確認してください")
    log(f"[ERROR] ボタン未検出（現在の状態: {state}）。スクリーンショット: {screenshot_path}")
    return False


async def reauth(action: str):
    """セッション切れ時にEdgeで再認証してから打刻
    channel='msedge' により Windows WAM/SSO が使われ、
    Accenture テナントではサインイン操作が不要になる場合がある。
    """
    SESSION_DIR.mkdir(exist_ok=True)

    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            str(SESSION_DIR),
            channel="msedge",          # Edge SSO / WAM でサインインを省略
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-features=msEdgeProfileSwitcher,AccountConsistencyService,IdentityConsistency,EdgeProfileSwitcherFREDialog",
                "--no-default-browser-check",
                "--no-first-run",
                "--disable-sync",
            ],
        )
        page = await ctx.new_page()
        await page.goto(MYTIM_URL, timeout=60_000)

        if not on_mytim(page.url):
            log("ブラウザでログインしてください（Edge が開いています）...")
            try:
                await page.wait_for_function(
                    f"() => location.hostname === '{MYTIM_HOST}'",
                    timeout=300_000,
                )
            except Exception:
                await ctx.close()
                notify("MyTim エラー", "ログインがタイムアウトしました")
                return

        # MyTim到達後、ページが完全に描画されるまで待機してからボタン操作
        try:
            await page.wait_for_load_state("networkidle", timeout=15_000)
        except Exception:
            pass
        await page.wait_for_timeout(3_000)

        await punch_on_page(page, action)
        await ctx.storage_state(path=str(STORAGE_FILE))
        await ctx.close()


async def do_punch(action: str):
    if not STORAGE_FILE.exists():
        notify("MyTim エラー", "setup.py を実行してログインしてください")
        log("[ERROR] storage.json が見つかりません")
        return

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--use-system-default-certificate-store",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        ctx = await browser.new_context(storage_state=str(STORAGE_FILE))
        page = await ctx.new_page()

        try:
            await page.goto(MYTIM_URL, timeout=60_000)
            await page.wait_for_load_state("networkidle", timeout=20_000)

            if not on_mytim(page.url):
                log(f"セッション切れ（URL: {page.url[:80]}）。再認証が必要です...")
                await ctx.close()
                await browser.close()
                await reauth(action)
                return

            await punch_on_page(page, action)
            await ctx.storage_state(path=str(STORAGE_FILE))

        except Exception as e:
            notify("MyTim エラー", str(e)[:80])
            log(f"[ERROR] {e}\n{traceback.format_exc()}")
        finally:
            await ctx.close()
            await browser.close()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        log("使い方: python punch.py [出勤|休憩|退勤|1|2|3]")
        sys.exit(1)
    arg = sys.argv[1]
    action = ACTION_MAP.get(arg, arg)
    if action not in BUTTONS:
        log(f"不正な引数: {repr(arg)}")
        sys.exit(1)
    log(f"--- 打刻開始: {action} ---")
    try:
        asyncio.run(do_punch(action))
    except Exception as e:
        log(f"[FATAL] {e}\n{traceback.format_exc()}")
