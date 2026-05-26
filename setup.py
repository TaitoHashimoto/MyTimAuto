"""
MyTim 自動打刻 - 初期セットアップ（初回または再認証時に実行）
Chromeでログイン後、セッションを storage.json に保存します。
"""
import asyncio
import subprocess
import sys
from pathlib import Path
from playwright.async_api import async_playwright

BASE_DIR = Path(__file__).parent
SESSION_DIR = BASE_DIR / "session"
STORAGE_FILE = BASE_DIR / "storage.json"

# 個人/組織固有の値は config.py から読み込む
try:
    from config import MYTIM_URL, MYTIM_HOST, TEAMS_CHAT_ID, TEAMS_ANCHOR_MSG
except ImportError as _e:
    raise SystemExit(
        "config.py が見つかりません。"
        "config.example.py を config.py にコピーして値を設定してください。"
    ) from _e


def get_desktop_path() -> Path:
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", "[Environment]::GetFolderPath('Desktop')"],
        capture_output=True, text=True, creationflags=0x08000000,
    )
    path = result.stdout.strip()
    return Path(path) if path else Path.home() / "Desktop"


def create_desktop_shortcuts():
    desktop = get_desktop_path()
    vbs_dir = BASE_DIR
    icons = {
        "出勤": ("MyTim_出勤.lnk", str(vbs_dir / "出勤.vbs"), "77"),
        "休憩": ("MyTim_休憩.lnk", str(vbs_dir / "休憩.vbs"), "130"),
        "退勤": ("MyTim_退勤.lnk", str(vbs_dir / "退勤.vbs"), "131"),
    }
    for label, (link_name, target, icon_idx) in icons.items():
        link_path = desktop / link_name
        ps = f"""
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut('{link_path}')
$sc.TargetPath = 'wscript.exe'
$sc.Arguments = '/nologo "{target}"'
$sc.IconLocation = 'C:\\Windows\\System32\\shell32.dll,{icon_idx}'
$sc.WindowStyle = 7
$sc.Save()
"""
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            creationflags=0x08000000,
        )
        print(f"  ショートカット作成: {link_path}")


async def setup():
    SESSION_DIR.mkdir(exist_ok=True)

    print("=" * 50)
    print("  MyTim 自動打刻 - 初期セットアップ")
    print("=" * 50)
    print()
    print("Chromeブラウザが開きます。")
    print("MyTimにSSOでログインしてください。")
    print("打刻画面が表示されると自動でセッションが保存されます。")
    print()

    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            str(SESSION_DIR),
            headless=False,
            channel="msedge",
            args=[
                "--disable-blink-features=AutomationControlled",
                # Edgeのアカウント検出/プロファイル切替プロンプトを無効化
                "--disable-features=msEdgeProfileSwitcher,AccountConsistencyService,IdentityConsistency,EdgeProfileSwitcherFREDialog",
                "--no-default-browser-check",
                "--no-first-run",
                "--disable-sync",
            ],
        )
        page = await ctx.new_page()

        print("MyTimに接続中...")
        try:
            await page.goto(MYTIM_URL, timeout=60_000)
        except Exception:
            pass

        print("ログインを待機中 (最大5分)...")

        # hostname で判定（redirect_uri 誤検出を防ぐ）
        try:
            await page.wait_for_function(
                f"() => location.hostname === '{MYTIM_HOST}'",
                timeout=300_000,
            )
        except Exception:
            print("警告: タイムアウトしました。")
            await ctx.close()
            return

        print(f"MyTim到達確認: {page.url}")
        print("クッキーを保存中...")

        try:
            await page.wait_for_load_state("networkidle", timeout=15_000)
        except Exception:
            pass
        await page.wait_for_timeout(3_000)

        # Teams のセッションも取得（残業報告自動化に使用）
        # 残業報告チャットに直接遷移し、サインインバナーが消えるまでユーザに任せる
        print()
        print("=" * 50)
        print("  Teamsサインイン確認")
        print("=" * 50)
        print("Teamsを開きます。")
        print("「もう一度サインインする必要があります」バナーが表示されたら、")
        print("画面右上の「サインイン」ボタンをクリックして完了させてください。")
        print("(最大5分待機します。完了すると自動で次に進みます)")
        print()

        teams_chat_url = (
            f"https://teams.microsoft.com/v2/#/l/message/"
            f"{TEAMS_CHAT_ID}/{TEAMS_ANCHOR_MSG}"
            f"?context=%7B%22contextType%22%3A%22chat%22%7D"
        )
        teams_page = await ctx.new_page()
        try:
            await teams_page.goto(
                teams_chat_url, timeout=60_000, wait_until="domcontentloaded"
            )
            await teams_page.wait_for_timeout(5_000)
            # 「Use the web app instead」が出たらクリック
            try:
                web_btn = teams_page.locator("button:has-text('Use the web app instead')")
                if await web_btn.count() > 0:
                    await web_btn.first.click()
                    await teams_page.wait_for_timeout(5_000)
            except Exception:
                pass

            # ユーザにサインインを促し、完了報告（ENTERキー）を待つ
            print()
            print("============================================================")
            print(" 【手順】 Edge ブラウザでTeamsの状態を確認してください")
            print("============================================================")
            print()
            print(" ▼ パターンA: 「Edge プロファイルの切り替え」が出た場合")
            print("   → 中央付近の青いリンク")
            print("     『サインアウトして別のアカウントでサインインしてください』")
            print("     をクリックしてください")
            print("   → 改めて自分のAccentureアカウントでサインインしてください")
            print()
            print(" ▼ パターンB: 「もう一度サインインする必要があります」")
            print("              バナーが上部に出ている場合")
            print("   → 右側の「サインイン」ボタンをクリックして")
            print("     サインインを完了させてください")
            print()
            print(" ▼ パターンC: 既にチャットが表示されている")
            print("   → そのまま次へ進んでOK")
            print()
            print(" 共通: チャット「[TAISHO]残業連絡」の最新メッセージが")
            print("       見えるようになったら、このターミナルに戻って ENTER")
            print("       キーを押してください")
            print("============================================================")
            print()
            # input() を別スレッドで実行（asyncio イベントループをブロックしない）
            try:
                await asyncio.to_thread(input, ">>> 上記が完了したら ENTER キーを押してください: ")
            except (EOFError, KeyboardInterrupt):
                print("中断されました。")
                return

            await teams_page.wait_for_timeout(3_000)
            print(f"Teams URL: {teams_page.url}")
        except Exception as e:
            print(f"Teams接続スキップ: {e}")
        finally:
            await teams_page.close()

        # セッション（クッキー）を JSON ファイルに保存（MyTim + Teams 両方）
        await ctx.storage_state(path=str(STORAGE_FILE))
        await ctx.close()

    print()
    print(f"セッション保存完了: {STORAGE_FILE}")
    print()
    print("デスクトップショートカットを作成中...")
    create_desktop_shortcuts()
    print()
    print("=" * 50)
    print("  セットアップ完了！")
    print("=" * 50)
    print()
    print("デスクトップのショートカットをダブルクリックするだけで打刻できます。")
    print("  MyTim_出勤.lnk / MyTim_休憩.lnk / MyTim_退勤.lnk")


if __name__ == "__main__":
    asyncio.run(setup())
