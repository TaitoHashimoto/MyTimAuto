@echo off
chcp 65001 > nul
echo ================================================
echo   MyTim 自動打刻 - インストール
echo ================================================
echo.
echo [1/3] Playwright をインストール中...
python -m pip install playwright
if %errorlevel% neq 0 (
    echo エラー: pip install が失敗しました
    pause
    exit /b 1
)
echo.
echo [2/3] Playwright 用 Edge を準備中...
python -m playwright install msedge
if %errorlevel% neq 0 (
    echo エラー: playwright install msedge が失敗しました
    pause
    exit /b 1
)
echo.
echo [3/3] 初期セットアップを開始します（config.py を作成してから実行してください）...
if not exist config.py (
    echo.
    echo *** config.py が見つかりません ***
    echo config.example.py をコピーして config.py を作成し、値を設定してから
    echo もう一度 install.bat を実行するか、python setup.py を直接実行してください。
    echo.
    pause
    exit /b 1
)
python setup.py
echo.
pause
