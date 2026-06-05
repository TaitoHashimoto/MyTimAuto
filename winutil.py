"""Windows GUI ユーティリティ
- 画面解像度の取得（論理ピクセル / Edgeのウィンドウ座標系と一致）
- ウィンドウの中央座標計算
- 起動済み Edge ウィンドウを CDP 経由で移動

各スクリプト（setup.py / overtime.py / punch.py）から共通利用する。
"""
from __future__ import annotations

import ctypes


# 画面外位置（ユーザーに見せない時の配置）
OFFSCREEN_X = -32000
OFFSCREEN_Y = -32000


def get_screen_size() -> tuple[int, int]:
    """プライマリディスプレイの解像度 (width, height) を論理ピクセルで返す。
    Edge の --window-position / --window-size も論理ピクセルを使うため整合する。
    """
    u = ctypes.windll.user32
    return u.GetSystemMetrics(0), u.GetSystemMetrics(1)


def fit_window_size(target_w: int = 1280, target_h: int = 900,
                    ratio: float = 0.85) -> tuple[int, int]:
    """画面サイズに収まるよう、目標サイズ or 画面の指定割合のうち
    小さい方を採用してウィンドウサイズを決定する。"""
    sw, sh = get_screen_size()
    w = min(target_w, int(sw * ratio))
    h = min(target_h, int(sh * ratio))
    return w, h


def center_window_position(window_w: int, window_h: int) -> tuple[int, int]:
    """指定サイズのウィンドウを画面中央に配置するための (x, y) 座標を返す。"""
    sw, sh = get_screen_size()
    return max(0, (sw - window_w) // 2), max(0, (sh - window_h) // 2)


def offscreen_args(window_w: int | None = None, window_h: int | None = None) -> list[str]:
    """Edge をユーザーに見えない位置で起動するためのコマンドライン引数。"""
    if window_w is None or window_h is None:
        window_w, window_h = fit_window_size()
    return [
        f"--window-position={OFFSCREEN_X},{OFFSCREEN_Y}",
        f"--window-size={window_w},{window_h}",
    ]


def centered_args(window_w: int | None = None, window_h: int | None = None) -> list[str]:
    """Edge を画面中央で起動するためのコマンドライン引数。"""
    if window_w is None or window_h is None:
        window_w, window_h = fit_window_size()
    cx, cy = center_window_position(window_w, window_h)
    return [
        f"--window-position={cx},{cy}",
        f"--window-size={window_w},{window_h}",
    ]


async def move_browser_window_to_center(page) -> bool:
    """CDP 経由で Edge ウィンドウを画面中央に移動する。
    画面外で起動した Edge を、ユーザー操作（サインイン等）が必要だと
    判明したタイミングで画面に出すために使用する。
    成功時 True、失敗時 False を返す（ベストエフォート）。
    """
    try:
        window_w, window_h = fit_window_size()
        cx, cy = center_window_position(window_w, window_h)
        session = await page.context.new_cdp_session(page)
        info = await session.send("Browser.getWindowForTarget")
        await session.send(
            "Browser.setWindowBounds",
            {
                "windowId": info["windowId"],
                "bounds": {
                    "left": cx,
                    "top": cy,
                    "width": window_w,
                    "height": window_h,
                    "windowState": "normal",
                },
            },
        )
        try:
            await session.detach()
        except Exception:
            pass
        return True
    except Exception:
        return False
