# MyTimAuto

Accenture 社内勤怠システム **MyTim** の打刻と **標準労働時間超過報告（残業申請）** を自動化する Python ツールセットです。Playwright + Microsoft Edge を使い、デスクトップアイコンやタスクスケジューラから運用できます。

> ⚠️ 本ツールは Accenture 社内向けの個人利用を想定しています。利用は自己責任でお願いします。

---

## 主な機能

| スクリプト         | 機能                                                                 |
| ------------------ | -------------------------------------------------------------------- |
| `setup.py`         | 初期セットアップ（MyTim / Teams へのSSOサインインとセッション保存）。デスクトップに `MyTim_出勤.lnk` / `MyTim_休憩.lnk` / `MyTim_退勤.lnk` を作成 |
| `punch.py`         | 出勤 / 休憩 / 退勤の打刻。状態（未出勤・勤務中・休憩中）に応じてボタン押下可否を自動判定 |
| `overtime.py`      | Teams の残業申請チャットから「自分の本日分のメッセージ」を検出し、MyTim の「標準労働時間超過報告」フォームへ自動送信 |

---

## 動作環境

- Windows 10 / 11
- Python 3.10 以降（zoneinfo 互換であれば 3.9 でも可、ただし開発・確認は 3.14 で実施）
- Microsoft Edge（Accenture コンプライアンス上 Chrome 不可）
- 社内ネットワーク接続（MyTim / Teams にアクセスできる状態）

---

## セットアップ

### 1. 依存ライブラリのインストール

```powershell
pip install playwright
playwright install msedge
```

### 2. 設定ファイルの作成

`config.example.py` をコピーして `config.py` を作り、自分の環境に合わせて値を埋めます。
`config.py` は `.gitignore` 対象なので、個人情報を含めてもリポジトリには反映されません。

```powershell
Copy-Item config.example.py config.py
notepad config.py
```

設定する値:

| 変数               | 内容                                                                                       |
| ------------------ | ------------------------------------------------------------------------------------------ |
| `MYTIM_URL`        | MyTim 打刻画面のURL                                                                        |
| `MYTIM_HOST`       | MyTim のホスト名（セッション判定に使用）                                                  |
| `TEAMS_CHAT_ID`    | 残業申請を投稿するグループチャットのID（`19:xxxx@thread.v2`）                            |
| `TEAMS_ANCHOR_MSG` | チャットを開く際のアンカーメッセージID（任意のメッセージ）                                |
| `MY_NAME`          | Teams 上の自分の表示名（例: `"Yamada, Taro"` ）。残業申請の送信者フィルタに使用          |
| `MY_EMAIL`         | メールアドレスのローカルパート                                                            |

**Teams チャットIDの取得方法**

Teams のチャット画面 → 右上「︙」→「チャットへのリンクをコピー」  
コピーしたURL `https://teams.microsoft.com/l/message/19:xxxxxxx@thread.v2/000000?context=...` の：

- `19:xxxxxxx@thread.v2` → `TEAMS_CHAT_ID`
- `000000` → `TEAMS_ANCHOR_MSG`

### 3. 初回サインイン

```powershell
python setup.py
```

Edge が立ち上がるので、MyTim → Teams の順に SSO サインインを完了させます。  
「Edge プロファイルの切り替え」または「もう一度サインインする必要があります」が出た場合は、画面の案内に従ってサインインしてください。  
完了したらターミナルで ENTER。

完了するとデスクトップに `MyTim_出勤.lnk` / `MyTim_休憩.lnk` / `MyTim_退勤.lnk` が作成されます。

---

## 使い方

### 出勤・休憩・退勤の打刻

デスクトップショートカットをダブルクリックするだけで打刻できます。

コマンドで実行する場合:

```powershell
python punch.py 出勤   # または: python punch.py 1
python punch.py 休憩   # または: python punch.py 2
python punch.py 退勤   # または: python punch.py 3
```

### 残業申請の自動送信

Teams に所定のフォーマットで残業申請を投稿しておけば、`overtime.py` を実行することで MyTim の「標準労働時間超過報告」フォームへ自動送信されます。

**Teams への投稿フォーマット例**:

```
日付：YYYY/MM/DD(曜日)
残業予定時間：X.Xh
残業実績時間：X.Xh
残業理由：...
当月深夜勤務累計：X.Xh
当月残業実績累計(休暇込み)：X.Xh
当月休暇実績時間：X.Xh
```

**ドライラン（送信せず検出結果を確認）**:

```powershell
python overtime.py --force --dry-run
```

**本番実行**:

```powershell
python overtime.py --force
```

### タスクスケジューラで自動実行

`overtime_task.xml` を Windows タスクスケジューラにインポートすると、平日 15:00〜18:00 の間 15分おきに `overtime.py` が実行されます。Teams に当日分の残業申請を投稿しておけば、最寄りの実行タイミングで MyTim へ自動送信されます。

```powershell
schtasks /Create /XML overtime_task.xml /TN "MyTim_Overtime"
```

---

## ファイル構成

```
MyTimAuto/
├── README.md
├── .gitignore
├── config.example.py        # 設定テンプレート（コピー元）
├── config.py                # ローカル設定（.gitignore対象）
├── setup.py                 # 初期セットアップ
├── punch.py                 # 打刻スクリプト
├── overtime.py              # 残業申請自動送信
├── overtime.vbs             # overtime.py のラッパー（タスクスケジューラ用）
├── overtime_task.xml        # タスクスケジューラ設定
├── 出勤.vbs / 休憩.vbs / 退勤.vbs  # 各打刻のVBSラッパー
├── install.bat              # セットアップ補助バッチ
├── requirements.txt         # Python依存ライブラリ
├── test_edge.py             # Edge + Teams 接続診断
└── test_session.py          # MyTim セッション確認診断
```

---

## トラブルシューティング

| 症状                                                       | 対処                                                                 |
| ---------------------------------------------------------- | -------------------------------------------------------------------- |
| Teams で「もう一度サインインする必要があります」バナー      | `python setup.py` を再実行                                          |
| `net::ERR_FAILED` でTeams に接続できない                    | Edge プロセスを全終了 → `Get-Process msedge \| Stop-Process -Force`  |
| 残業申請が検出されない                                      | `overtime.log` を確認。Teams の投稿フォーマットが正しいかチェック  |
| MyTim 打刻ボタンが見つからない                              | 現在の状態（未出勤/勤務中/休憩中）と操作の組み合わせを確認          |

---

## ライセンス

個人利用のサンプル実装です。コピー・改変はご自由にどうぞ。
