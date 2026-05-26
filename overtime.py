"""
MyTim 残業超過報告 自動化スクリプト
Edge 永続プロファイルで Teams を読み取り、残業申請メッセージを検出したら
MyTim の「標準労働時間超過報告」フォームを自動送信する。
14:00〜18:00（平日）の間に10分ごとに実行されることを想定。
"""
import re
import json
import html
import asyncio
import subprocess
import traceback
from datetime import datetime, date, timezone, timedelta
from pathlib import Path
from urllib.parse import urlparse
from playwright.async_api import async_playwright

# 全ての日付・時刻はJST固定で扱う（PCのタイムゾーン設定に依存しないため）
# tzdataパッケージなしで動かすため timezone(timedelta) 形式を使う
JST = timezone(timedelta(hours=9), name="JST")


def now_jst() -> datetime:
    """現在時刻（JST）を返す。"""
    return datetime.now(JST)


def today_jst() -> date:
    """本日の日付（JST）を返す。"""
    return datetime.now(JST).date()


BASE_DIR = Path(__file__).parent
LOG_FILE = BASE_DIR / "overtime.log"
STORAGE_FILE = BASE_DIR / "storage.json"
SESSION_DIR = BASE_DIR / "session"
STATE_FILE = BASE_DIR / "overtime_state.json"

# 個人/組織固有の値は config.py から読み込む（config.example.py を参照）
try:
    from config import (
        MYTIM_URL,
        MYTIM_HOST,
        TEAMS_CHAT_ID,
        TEAMS_ANCHOR_MSG,
        MY_NAME,
        MY_EMAIL,
    )
except ImportError as _e:
    raise SystemExit(
        "config.py が見つかりません。"
        "config.example.py を config.py にコピーして値を設定してください。"
    ) from _e

# l/message URL はランチャー経由になるため、
# ランチャーが使う /_#/l/message/ を /v2/#/l/message/ に変換した直接URLを使用
# このURLを使うとサイドバーで対象チャットがハイライトされ最新メッセージプレビューが見える
TEAMS_URL = (
    f"https://teams.microsoft.com/v2/#/l/message/{TEAMS_CHAT_ID}/{TEAMS_ANCHOR_MSG}"
    f"?context=%7B%22contextType%22%3A%22chat%22%7D"
)

# メッセージ検出パターン
MSG_BLOCK_PATTERN = re.compile(
    r"日付[：:].*?当月休暇実績時間[：:][\d.]+\s*h?",
    re.DOTALL,
)
HOURS_PATTERN = re.compile(r"残業予定時間[：:]\s*([\d.]+)\s*h?")
DATE_PATTERN = re.compile(r"日付[：:]\s*(\d{4})/(\d{1,2})/(\d{1,2})")


def clean_overtime_message(raw: str) -> str:
    """Teamsページ本文から7フィールドを抽出し200字以内に収める"""
    # \xa0（ノーブレークスペース）や改行を通常スペースに正規化
    s = raw.replace('\xa0', ' ').replace('\n', ' ').replace('\r', ' ')

    # テキスト値フィールド用: 次のフィールドラベルで停止
    NEXT = (
        r"(?=\s*(?:"
        r"日付[：:]|残業予定時間[：:]|残業実績時間[：:]|残業理由[：:]"
        r"|当月深夜勤務累計[：:]|当月残業実績累計[^：:]*[：:]|当月休暇実績時間[：:]"
        r"))"
    )
    patterns = [
        r"日付[：:]\s*\d{4}/\d{1,2}/\d{1,2}(?:[（(][^）)]+[）)])?",
        r"残業予定時間[：:]\s*[\d.]+\s*h?",
        r"残業実績時間[：:]\s*[\d.]+\s*h?",
        rf"残業理由[：:]\s*.*?{NEXT}",
        r"当月深夜勤務累計[：:]\s*[\d.]+\s*h?",
        rf"当月残業実績累計[^：:]*[：:]\s*.*?{NEXT}",
        r"当月休暇実績時間[：:]\s*[\d.]+\s*h?",
    ]
    lines = []
    for pat in patterns:
        m = re.search(pat, s)
        if m:
            lines.append(m.group(0).strip())
    return "\n".join(lines)[:200]


def message_is_today(message: str) -> bool:
    """メッセージの日付が今日かどうかチェック"""
    m = DATE_PATTERN.search(message)
    if not m:
        return False
    try:
        msg_date = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        return msg_date == today_jst()
    except ValueError:
        return False


def log(msg: str):
    timestamp = now_jst().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    try:
        print(line)
    except UnicodeEncodeError:
        print(line.encode(errors="replace").decode(errors="replace"))
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_state() -> dict:
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"last_report_date": "", "last_message_hash": ""}


def save_state(state: dict):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def notify(title: str, msg: str):
    ps = f"""
Add-Type -AssemblyName System.Windows.Forms
$n = New-Object System.Windows.Forms.NotifyIcon
$n.Icon = [System.Drawing.SystemIcons]::Application
$n.Visible = $true
$n.BalloonTipTitle = "{title}"
$n.BalloonTipText = "{msg}"
$n.ShowBalloonTip(6000)
Start-Sleep -Seconds 7
$n.Visible = $false
$n.Dispose()
"""
    subprocess.Popen(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
        creationflags=0x08000000,
    )


async def get_teams_message() -> str | None:
    """Edge永続プロファイルでTeamsチャットから残業申請メッセージを取得。
    送信者（MY_NAME）かつ本日日付のメッセージを対象とする。
    """
    log("Teamsメッセージを確認中 (Edge)...")

    if not SESSION_DIR.exists():
        log("[ERROR] sessionフォルダがありません。setup.pyを実行してください。")
        return None

    async with async_playwright() as p:
        # headless=False: headlessモード・最小化状態だとTeamsのチャット本文が描画されない
        ctx = await p.chromium.launch_persistent_context(
            str(SESSION_DIR),
            channel="msedge",
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--window-size=1920,1080",
                # Edgeのアカウント検出/プロファイル切替プロンプトを無効化
                "--disable-features=msEdgeProfileSwitcher,AccountConsistencyService,IdentityConsistency,EdgeProfileSwitcherFREDialog",
                "--no-default-browser-check",
                "--no-first-run",
                "--disable-sync",
            ],
        )
        page = await ctx.new_page()
        try:
            # Teams SPAは "load" イベントが発火しないことがあるので
            # wait_until="domcontentloaded" で navigate を完了させる
            # ERR_FAILED 等の一時エラーに備えて最大3回リトライ
            nav_ok = False
            for nav_attempt in range(3):
                try:
                    await page.goto(
                        TEAMS_URL,
                        timeout=60_000,
                        wait_until="domcontentloaded",
                    )
                    nav_ok = True
                    break
                except Exception as e:
                    log(f"Teamsナビゲーション試行{nav_attempt + 1}/3失敗: {str(e)[:120]}")
                    if nav_attempt < 2:
                        # Teamsホームを経由してリトライ
                        try:
                            await page.goto(
                                "https://teams.microsoft.com/v2/",
                                timeout=60_000,
                                wait_until="domcontentloaded",
                            )
                        except Exception:
                            pass
                        await page.wait_for_timeout(3_000)
            if not nav_ok:
                log("[ERROR] Teamsへのナビゲーションに3回失敗しました。")
                log("Edgeプロセスが残っている可能性があります。Task Managerで msedge.exe を全て終了してから再試行してください。")
                return None

            # ── URL確認・タブ切り替え処理 ──────────────────────────────────────
            await page.wait_for_timeout(2_000)
            log(f"ナビゲーション後URL: {page.url[:100]}")

            # v2/#/l/message URL がランチャー経由になった場合や別タブが開いた場合の対応
            if "dl/launcher" in page.url or "launcher.html" in page.url:
                # ランチャーURLパラメータから直接 Teams v2 URLを構築して再ナビゲート
                from urllib.parse import urlparse, parse_qs, unquote
                parsed_launcher = urlparse(page.url)
                lparams = parse_qs(parsed_launcher.query)
                if "url" in lparams:
                    web_path = unquote(lparams["url"][0])
                    direct_url = (
                        "https://teams.microsoft.com"
                        + web_path.replace("/_#/", "/v2/#/")
                    )
                    log(f"ランチャーをスキップ → {direct_url[:100]}")
                    await page.goto(direct_url, timeout=60_000)
                    await page.wait_for_load_state("domcontentloaded", timeout=30_000)
                    await page.wait_for_timeout(2_000)

            elif page.url in ("about:blank", TEAMS_URL):
                # 別タブにTeamsが開いた場合は切り替え
                for p in ctx.pages:
                    if ("teams.microsoft.com" in p.url
                            and "launcher" not in p.url
                            and p.url != "about:blank"):
                        page = p
                        log(f"別タブのTeamsに切り替え: {page.url[:80]}")
                        break

            # Teams SPA が描画されるまで待機（最大90秒）
            try:
                await page.wait_for_function(
                    "() => document.body.innerText.length > 200",
                    timeout=90_000,
                )
            except Exception:
                log("Teams描画タイムアウト")

            await page.wait_for_timeout(3_000)

            # セッション確認（URL で判定）
            if "login.microsoftonline" in page.url:
                log("Teamsセッション切れ。setup.pyを再実行してください。")
                notify("残業報告エラー ⚠", "Teamsのセッションが切れています\nsetup.pyを実行してください")
                return None

            # ── キャッシュ対策: いったんTeamsホームに遷移し、再度チャットへ遷移 ──
            # Teams SPAは永続セッションだとキャッシュが残ったままで最新メッセージを
            # 取得しないことがある。ホーム→対象チャットの二段階遷移で強制的に再取得させる
            log("Teamsを再ナビゲートして最新メッセージを取得します")
            try:
                await page.goto(
                    "https://teams.microsoft.com/v2/",
                    timeout=60_000,
                    wait_until="domcontentloaded",
                )
                await page.wait_for_timeout(4_000)
                await page.goto(
                    TEAMS_URL,
                    timeout=60_000,
                    wait_until="domcontentloaded",
                )
                await page.wait_for_function(
                    "() => document.body.innerText.length > 200",
                    timeout=60_000,
                )
            except Exception as e:
                log(f"再ナビゲート時のタイムアウト: {e}")

            # WebSocket同期と最新メッセージ受信を待つ（長めに待機）
            await page.wait_for_timeout(10_000)

            # ── Teams再サインイン要求バナーの検知 ─────────────────────────────
            # Teamsはセッション切れ時、ページ上部に赤いバナーで再サインインを要求する。
            # この状態では古いキャッシュのみ表示され、サーバから最新メッセージが
            # 取れない。バナーは body.innerText に含まれないことがあるため、
            # ドキュメント全体（HTMLおよびDocumentElement）をスキャンする。
            reauth_required = await page.evaluate("""
                () => {
                    const keywords = [
                        "もう一度サインイン",
                        "再度サインイン",
                        "サインイン要求",
                        "セッションの有効期限",
                        "Sign in again",
                        "session has expired",
                    ];
                    // documentElement.innerText (body外も含む)
                    const allText = (document.documentElement.innerText || "") + " "
                                  + (document.documentElement.textContent || "");
                    for (const k of keywords) {
                        if (allText.includes(k)) return { found: true, keyword: k };
                    }
                    // outerHTMLでもチェック（aria-label等の属性もカバー）
                    const html = (document.documentElement.outerHTML || "").substring(0, 50000);
                    for (const k of keywords) {
                        if (html.includes(k)) return { found: true, keyword: k + " (in HTML)" };
                    }
                    return { found: false };
                }
            """)
            if reauth_required.get("found"):
                log(f"[ERROR] Teams再サインイン要求を検知: 「{reauth_required.get('keyword')}」")
                log("Teamsのセッション切れにより最新メッセージが取得できません。setup.pyを再実行してください。")
                notify(
                    "残業報告エラー ⚠ Teams再サインイン必要",
                    "Teamsのセッションが期限切れです\nsetup.pyを実行してサインインし直してください",
                )
                return None

            d = today_jst()
            today_str       = d.strftime("%Y/%m/%d")           # 2026/05/25（ゼロ埋め）
            today_str_short = f"{d.year}/{d.month}/{d.day}"    # 2026/5/25（ゼロなし）

            # まず「最新メッセージへジャンプ」ボタンが表示されていればクリック
            # Teams は古いメッセージから開いた場合、下部に「↓ 新しいメッセージ」を表示する
            jumped = await page.evaluate("""
                () => {
                    // チャットペイン内のボタンを探す（サイドバーの「新しいチャット」を除外）
                    const labels = [
                        '最新', 'latest', 'Jump', 'jump',
                        '新しいメッセージへ', 'newer messages', '↓', '新しい投稿',
                    ];
                    // メイン領域内のボタンに限定
                    const mainRoots = document.querySelectorAll(
                        '[role="main"], [data-tid="message-pane"], [id*="message-pane"], [class*="message-pane"]'
                    );
                    for (const root of mainRoots) {
                        const btns = root.querySelectorAll('button, [role="button"]');
                        for (const btn of btns) {
                            const txt = (btn.textContent || '').trim();
                            const lbl = btn.getAttribute('aria-label') || '';
                            for (const k of labels) {
                                if ((txt && txt.includes(k)) || (lbl && lbl.includes(k))) {
                                    btn.click();
                                    return { clicked: true, text: txt.substring(0, 30), label: lbl.substring(0, 30) };
                                }
                            }
                        }
                    }
                    return { clicked: false };
                }
            """)
            if jumped.get("clicked"):
                log(f"「最新へジャンプ」ボタンをクリック: text={jumped.get('text')}, label={jumped.get('label')}")
                await page.wait_for_timeout(3_000)

            # メッセージスレッドを最下部へスクロール（最大8回 / 増加が止まったら終了）
            # サイドバーではなくメッセージペインをスクロール対象とする
            prev_scroll_h = 0
            no_growth = 0
            for attempt in range(8):
                # 既にターゲットが描画されていれば早期終了
                has_target = await page.evaluate(f"""
                    () => {{
                        const b = document.body.innerText || "";
                        return b.includes("日付：{today_str}")
                            || b.includes("日付:{today_str}")
                            || b.includes("日付：{today_str_short}")
                            || b.includes("日付:{today_str_short}");
                    }}
                """)
                if has_target:
                    log(f"スクロール{attempt + 1}回目: ターゲット検出済み、スクロール終了")
                    break

                # メッセージスレッドのコンテナを優先して探す
                scroll_info = await page.evaluate("""
                    () => {
                        const threadSels = [
                            '[data-tid="message-pane-list"]',
                            '[data-tid="messageListContainer"]',
                            '[role="log"]',
                            'div[id*="message-pane"]',
                            'div[class*="message-pane-list"]',
                            'div[class*="messagePaneList"]',
                        ];
                        for (const sel of threadSels) {
                            const el = document.querySelector(sel);
                            if (el && el.scrollHeight > el.clientHeight + 50) {
                                el.scrollTop = el.scrollHeight;
                                const rect = el.getBoundingClientRect();
                                return {
                                    sel: sel,
                                    cx: Math.round(rect.left + rect.width / 2) || 960,
                                    cy: Math.round(rect.top + rect.height / 2) || 540,
                                    scrollH: el.scrollHeight,
                                    clientH: el.clientHeight,
                                };
                            }
                        }
                        // フォールバック: スクロール可能な div を全探索（メインエリア内）
                        const mainEls = document.querySelectorAll(
                            '[role="main"] div, [data-tid*="message"] div'
                        );
                        for (const el of mainEls) {
                            if (el.scrollHeight > el.clientHeight + 100
                                && el.clientHeight > 200) {
                                el.scrollTop = el.scrollHeight;
                                const rect = el.getBoundingClientRect();
                                return {
                                    sel: 'fallback-scrollable',
                                    cx: Math.round(rect.left + rect.width / 2) || 960,
                                    cy: Math.round(rect.top + rect.height / 2) || 540,
                                    scrollH: el.scrollHeight,
                                    clientH: el.clientHeight,
                                };
                            }
                        }
                        return { sel: 'none', cx: 960, cy: 540, scrollH: 0, clientH: 0 };
                    }
                """)
                scroll_h = scroll_info['scrollH']
                grew = scroll_h > prev_scroll_h
                log(f"スクロール{attempt + 1}回目: sel={scroll_info['sel']}, "
                    f"scrollH={scroll_h}, clientH={scroll_info.get('clientH', '?')}"
                    f"{' (増加)' if grew else ''}")

                if grew:
                    no_growth = 0
                else:
                    no_growth += 1
                prev_scroll_h = scroll_h

                # マウスをメッセージスレッド中央に移動してホイール + Endキー
                await page.mouse.move(scroll_info["cx"], scroll_info["cy"])
                await page.mouse.wheel(0, 15000)
                await page.wait_for_timeout(500)
                await page.keyboard.press("End")
                await page.wait_for_timeout(2_500)

                # 2回連続で増加なし & 適切なコンテナが見つからない場合は打ち切り
                if no_growth >= 2 and scroll_info['sel'] == 'none':
                    log("スクロール可能なメッセージペインが見つからず、終了")
                    break

            # スクリーンショット保存（診断用）
            screenshot_path = BASE_DIR / "teams_debug.png"
            await page.screenshot(path=str(screenshot_path))
            log(f"スクリーンショット保存: {screenshot_path}")

            # デバッグ: ページ内キーワードの出現状況をログ出力
            debug_info = await page.evaluate(f"""
                () => {{
                    const today  = "{today_str}";
                    const todayS = "{today_str_short}";
                    const sender = "{MY_NAME}";
                    const body   = document.body.innerText || "";
                    const patterns = [
                        "日付：" + today, "日付:" + today,
                        "日付：" + todayS, "日付:" + todayS,
                    ];
                    let context = "(なし)";
                    for (const pat of patterns) {{
                        const idx = body.indexOf(pat);
                        if (idx >= 0) {{
                            context = body.substring(Math.max(0, idx - 10), idx + 80).replace(/\\n/g, "|");
                            break;
                        }}
                    }}
                    // 日付フィールドが見つからなくても「日付」の周辺テキストをログ
                    let dateAnyCtx = "(なし)";
                    const dateIdx = body.indexOf("日付");
                    if (dateIdx >= 0) {{
                        dateAnyCtx = body.substring(Math.max(0, dateIdx - 5), dateIdx + 60).replace(/\\n/g, "|");
                    }}
                    return {{
                        hasDateField: patterns.some(p => body.includes(p)),
                        hasSender:    body.includes(sender),
                        hasOT:        body.includes("残業予定時間"),
                        context:      context,
                        dateAnyCtx:   dateAnyCtx,
                        bodyLen:      body.length,
                    }};
                }}
            """)
            log(f"デバッグ: bodyLen={debug_info['bodyLen']}, "
                f"日付フィールドあり={debug_info['hasDateField']}, "
                f"送信者({MY_NAME})あり={debug_info['hasSender']}, "
                f"残業予定時間あり={debug_info['hasOT']}")
            log(f"デバッグ: 日付フィールド周辺={debug_info['context']}")
            log(f"デバッグ: 日付(任意)周辺={debug_info['dateAnyCtx']}")

            # ── 優先パス: 送信者(MY_NAME)の最新の残業申請メッセージを本文から直接抽出 ──
            # [TAISHO]残業連絡 は複数の人が投稿するグループチャット。
            # Teams DOMは送信者ヘッダと本文を別要素にするため、要素単位の検索では
            # 自分のメッセージを特定しづらい。
            # 戦略: body.innerText 内のすべての「日付：YYYY/MM/DD」を見つけ、
            #       その直前300文字以内に MY_NAME があるものを「自分の投稿」とみなす。
            #       全フィールドが揃った（=サイドバー要約ではなく本文）ものを優先し、
            #       その中で最新（=最も後ろの位置）を返す。
            my_latest = await page.evaluate(f"""
                () => {{
                    const sender = "{MY_NAME}";
                    const body   = document.body.innerText || "";
                    const re     = /日付[：:]\\s*(\\d{{4}})\\/(\\d{{1,2}})\\/(\\d{{1,2}})/g;
                    const nextDateRe = /日付[：:]\\s*\\d{{4}}\\/\\d{{1,2}}\\/\\d{{1,2}}/;
                    const matches = [];
                    let m;
                    while ((m = re.exec(body)) !== null) {{
                        // 直前300文字以内に送信者名があるかチェック（=自分の投稿）
                        const before = body.substring(Math.max(0, m.index - 300), m.index);
                        if (!before.includes(sender)) continue;
                        // 抽出範囲は「日付」の位置から開始（手前は含めず別メッセージとの混在を防ぐ）
                        // 終了位置は「次の 日付：YYYY/MM/DD まで」または1000文字
                        const after = body.substring(m.index + m[0].length);
                        const nd = after.match(nextDateRe);
                        let endOffset = 1000;
                        if (nd && nd.index < 1000) {{
                            endOffset = m[0].length + nd.index;
                        }}
                        const text = body.substring(m.index, m.index + endOffset);
                        const fields = ["残業予定時間", "残業実績時間", "当月休暇実績時間"];
                        const fieldCount = fields.filter(k => text.includes(k)).length;
                        matches.push({{
                            idx: m.index,
                            date: m[1] + "/" + m[2] + "/" + m[3],
                            text: text,
                            fieldCount: fieldCount,
                        }});
                    }}
                    if (!matches.length) return null;
                    // フィールド数が多い順、同じならば後方位置順（=最新）
                    matches.sort((a, b) => {{
                        if (a.fieldCount !== b.fieldCount) return b.fieldCount - a.fieldCount;
                        return b.idx - a.idx;
                    }});
                    return matches[0];
                }}
            """)
            if my_latest:
                msg_text = my_latest["text"]
                msg_date = my_latest["date"]
                field_count = my_latest["fieldCount"]
                log(f"自分の最新残業申請を検出: 日付={msg_date}, フィールド数={field_count}/3")
                if "残業予定時間" in msg_text and "残業実績時間" in msg_text:
                    safe_msg = msg_text.replace('\xa0', ' ')
                    log(f"rawメッセージ(自分): {safe_msg[:120].replace(chr(10), '|')}...")
                    clean = clean_overtime_message(safe_msg)
                    log(f"整形後(自分): {clean.replace(chr(10), ' | ')}")
                    if message_is_today(clean):
                        return clean
                    else:
                        log(f"自分の最新メッセージは {msg_date} で本日ではないため却下します")
                        # 本日分は無いと確定できるのでここで終了（DOM検索フォールバックは不要）
                        return None
                else:
                    log("自分の最新メッセージに必要なフィールドが揃わず、DOM検索にフォールバック")
            else:
                log(f"自分（{MY_NAME}）の残業申請メッセージが本文中に見つかりませんでした")

            # DOM から送信者（MY_NAME）＋残業フィールドを含む最小メッセージ要素を取得
            # 日付チェックは message_is_today() に委ねる（フォーマット差異を吸収）
            raw_texts = await page.evaluate(f"""
                () => {{
                    const sender   = "{MY_NAME}";
                    const required = ["日付", "残業予定時間", "当月休暇実績時間"];
                    const results  = [];
                    for (const el of document.querySelectorAll('*')) {{
                        const t = (el.innerText || "");
                        if (t.includes(sender)
                            && required.every(k => t.includes(k))) {{
                            results.push({{ len: t.length, text: t }});
                        }}
                    }}
                    if (!results.length) return [];
                    results.sort((a, b) => a.len - b.len);
                    const minLen = results[0].len;
                    return results
                        .filter(r => r.len <= minLen + 200)
                        .map(r => r.text);
                }}
            """)

            if not raw_texts:
                log("残業申請メッセージが見つかりませんでした（DOM検索: 送信者+残業フィールド）")
                return None

            log(f"DOM候補: {len(raw_texts)}件")
            # 全7フィールドを含むものを優先
            full_candidates = [
                t for t in raw_texts
                if "当月深夜勤務累計" in t and "当月残業実績累計" in t
            ]
            target = (full_candidates[-1] if full_candidates else raw_texts[-1]).strip()
            # \xa0（ノーブレークスペース）などを通常スペースに正規化してからログ出力
            target = target.replace('\xa0', ' ').replace('​', '').replace('‍', '')
            log(f"rawメッセージ: {target[:120].replace(chr(10), '|')}...")
            clean = clean_overtime_message(target)
            log(f"整形後: {clean.replace(chr(10), ' | ')}")

            # 最終確認: 抽出した日付が本日であること
            if not message_is_today(clean):
                log("本日分の残業申請メッセージがありません（Teamsに今日の投稿が見当たりません）")
                # 検出された最新メッセージが2日以上前のものなら、
                # Teamsセッション切れ（「もう一度サインイン」バナー表示）の可能性を警告
                m = DATE_PATTERN.search(clean)
                if m:
                    try:
                        msg_date = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                        days_old = (today_jst() - msg_date).days
                        if days_old >= 2:
                            log(f"[警告] 検出された最新メッセージは{days_old}日前のものです。")
                            log("Teamsで「もう一度サインインする必要があります」バナーが出ている可能性があります。")
                            log("→ その場合は setup.py を実行してTeamsに再サインインしてください。")
                            notify(
                                "残業報告 ⚠ 最新メッセージ未取得",
                                f"検出された最新メッセージが{days_old}日前です\n"
                                "Teams再サインインが必要な可能性 → setup.py 実行",
                            )
                    except ValueError:
                        pass
                return None

            return clean

        except Exception as e:
            log(f"Teams取得エラー: {e}\n{traceback.format_exc()}")
            return None
        finally:
            await ctx.close()


async def _submit_on_page(page, hours: str, reason: str) -> bool:
    """MyTim フォームへの入力・送信（ページオブジェクト渡し）"""
    try:
        await page.wait_for_timeout(2_000)
        clicked = await page.evaluate("""
            () => {
                const btn = [...document.querySelectorAll('*')].find(
                    el => el.childElementCount === 0
                      && el.textContent.trim() === '標準労働時間超過報告'
                );
                if (btn) { btn.click(); return true; }
                return false;
            }
        """)
        if not clicked:
            screenshot_path = BASE_DIR / "mytim_debug.png"
            await page.screenshot(path=str(screenshot_path), full_page=True)
            log(f"「標準労働時間超過報告」ボタンが見つかりません。スクリーンショット: {screenshot_path}")
            return False
        await page.wait_for_timeout(2_000)

        await page.wait_for_selector("text=標準労働時間超過報告", timeout=10_000)
        log("フォームが開きました")

        time_input = page.locator("input").last
        await time_input.clear()
        await time_input.fill(hours)
        log(f"報告時間: {hours}h")

        reason_input = page.locator("textarea").last
        await reason_input.clear()
        await reason_input.fill(reason)
        log(f"報告理由: {reason[:30]}...")

        submit_btn = page.locator("button:has-text('報告')").last
        await submit_btn.click()
        await page.wait_for_timeout(3_000)

        await page.screenshot(path=str(BASE_DIR / "mytim_after_submit.png"), full_page=True)
        log("[OK] 残業超過報告 送信完了")
        notify("残業報告完了 ✓", f"標準労働時間超過報告を提出しました\n残業予定時間: {hours}h")
        return True

    except Exception as e:
        log(f"MyTim報告エラー: {e}\n{traceback.format_exc()}")
        notify("残業報告エラー", str(e)[:80])
        return False


async def reauth_and_submit(hours: str, reason: str) -> bool:
    """セッション切れ時に Edge で再認証してから残業報告を送信する"""
    log("MyTimセッション切れ。Edgeで再認証を試みます...")
    SESSION_DIR.mkdir(exist_ok=True)

    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            str(SESSION_DIR),
            channel="msedge",
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
        try:
            await page.goto(MYTIM_URL, timeout=60_000)

            if urlparse(page.url).netloc != MYTIM_HOST:
                log("Edge でサインインしてください（ブラウザが開いています）...")
                try:
                    await page.wait_for_function(
                        f"() => location.hostname === '{MYTIM_HOST}'",
                        timeout=300_000,
                    )
                except Exception:
                    notify("残業報告エラー ⚠", "ログインがタイムアウトしました\nsetup.pyを実行してください")
                    return False

            try:
                await page.wait_for_load_state("networkidle", timeout=15_000)
            except Exception:
                pass
            await page.wait_for_timeout(3_000)

            result = await _submit_on_page(page, hours, reason)
            await ctx.storage_state(path=str(STORAGE_FILE))
            return result

        except Exception as e:
            log(f"再認証エラー: {e}\n{traceback.format_exc()}")
            notify("残業報告エラー", str(e)[:80])
            return False
        finally:
            await ctx.close()


async def submit_overtime_report(hours: str, reason: str) -> bool:
    """MyTim の標準労働時間超過報告フォームを送信する"""
    if not STORAGE_FILE.exists():
        # storage.json がなければ直接 Edge 再認証へ
        return await reauth_and_submit(hours, reason)

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
            log("MyTimに接続中...")
            await page.goto(MYTIM_URL, timeout=60_000)
            await page.wait_for_load_state("networkidle", timeout=20_000)

            if urlparse(page.url).netloc != MYTIM_HOST:
                # セッション切れ → Edge で自動再認証
                await ctx.close()
                await browser.close()
                return await reauth_and_submit(hours, reason)

            return await _submit_on_page(page, hours, reason)

        except Exception as e:
            log(f"MyTim報告エラー: {e}\n{traceback.format_exc()}")
            notify("残業報告エラー", str(e)[:80])
            return False
        finally:
            try:
                await ctx.close()
                await browser.close()
            except Exception:
                pass


async def run(force: bool = False, dry_run: bool = False):
    now = now_jst()

    if not force:
        if now.weekday() >= 5:
            log("土日のためスキップ")
            return
        if not (15 <= now.hour < 18):
            log(f"対象時間外（{now.strftime('%H:%M')}）のためスキップ")
            return

    state = load_state()
    today = today_jst().isoformat()

    if not dry_run and state.get("last_report_date") == today:
        log("本日の残業報告は送信済みです。スキップします。")
        return

    message = await get_teams_message()
    if not message:
        return

    import hashlib
    msg_hash = hashlib.md5(message.encode()).hexdigest()
    if not dry_run and msg_hash == state.get("last_message_hash"):
        log("新しいメッセージはありません。スキップします。")
        return

    match = HOURS_PATTERN.search(message)
    if not match:
        log("「残業予定時間：」が見つかりませんでした")
        return
    hours = match.group(1)

    if dry_run:
        log(f"[DRY-RUN] 送信予定 残業予定時間: {hours}h")
        log(f"[DRY-RUN] 送信予定 報告理由:\n{message}")
        log("[DRY-RUN] MyTimへの送信はスキップしました")
        return

    log(f"新しい残業申請を検出: {hours}h")

    success = await submit_overtime_report(hours, message)

    if success:
        state["last_report_date"] = today
        state["last_message_hash"] = msg_hash
        save_state(state)


if __name__ == "__main__":
    import sys
    force = "--force" in sys.argv
    dry_run = "--dry-run" in sys.argv
    mode = "（ドライラン）" if dry_run else "（強制実行）" if force else ""
    log(f"=== 残業報告チェック開始{mode} ===")
    try:
        asyncio.run(run(force=force, dry_run=dry_run))
    except Exception as e:
        log(f"[FATAL] {e}\n{traceback.format_exc()}")
    log("=== 残業報告チェック終了 ===")
