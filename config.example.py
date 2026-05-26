"""
MyTimAuto 設定ファイル（テンプレート）

このファイルを config.py にコピーして、自分の環境に合わせて値を書き換えてください。
config.py は .gitignore で除外されるため、個人情報を含めても安全です。

    cp config.example.py config.py
    # その後、config.py を編集
"""

# ──────────────────────────────────────────────────────────────────────
# MyTim
# ──────────────────────────────────────────────────────────────────────
MYTIM_URL = "https://whm.accenture.com/mytim/secure/punchClock"
MYTIM_HOST = "whm.accenture.com"

# ──────────────────────────────────────────────────────────────────────
# Teams 残業申請チャット
# ──────────────────────────────────────────────────────────────────────
# 対象のグループチャットID（"19:xxxxxxxxxxxx@thread.v2" 形式）
#   取得方法: Teamsチャットの右上「...」→「チャットへのリンクをコピー」
#   その URL の "/l/message/<このID>/<msgId>" 部分が TEAMS_CHAT_ID
TEAMS_CHAT_ID = "19:xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx@thread.v2"

# チャット内の任意のメッセージID（チャットを開くアンカーとして使用）
# 古いものでもOK。コピーしたURLの "/<msgId>?context=..." 部分。
TEAMS_ANCHOR_MSG = "0000000000000"

# ──────────────────────────────────────────────────────────────────────
# 自分のアカウント情報（Teams投稿のフィルタ用）
# ──────────────────────────────────────────────────────────────────────
# Teams 上の表示名（送信者名そのまま、姓名カンマ区切りなど Teams の表記に合わせる）
#   例: "Yamada, Taro"
MY_NAME = "Lastname, Firstname"

# 自分のメールアドレスのローカルパート（@より前）
MY_EMAIL = "firstname.lastname"
