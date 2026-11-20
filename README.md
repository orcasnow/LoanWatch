# LoanWatch

名古屋グランパスから期限付き移籍している選手について、直近の試合での出場状況を集めて Excel と CSV の一覧にする Windows 用アプリです。選手ごとに、最新試合が先発・途中出場・ベンチ入りのみ・ベンチ外・試合なしのどれだったかを、日付、対戦相手、試合結果、出場時間とあわせて出力します。

現在のバージョンは 4.2 です（変更履歴は [version.py](version.py)）。

## 取得できる内容

各選手の情報は、所属先に応じて次のサイトから取得します。どのサイトを使うかは選手ごとに `players.json` で指定します。

| 対象 | 取得元 |
| --- | --- |
| Jリーグのクラブ（リーグ戦） | Football-LAB のメンバー表。出場時間などはＪリーグ公式の出場履歴で補う |
| Jリーグのクラブ（Football-LAB が未更新のとき） | Ｊリーグ公式の出場履歴 |
| 海外クラブ | ESPN、Transfermarkt |
| 海外のリザーブチーム（例：アルヘンティノス・ジュニアーズII） | BeSoccer |

トップチームのリーグ戦のほかに、U-21 Ｊリーグ・ルヴァンカップ・天皇杯に今季の記録がある選手は、大会ごとに別の行を出します。J1〜J3 の行には、今季のリーグ戦の通算（出場数、先発数、出場時間、ゴール、警告・退場）も入ります。

これとは別に、指定した U-21 クラブの出場記録を試合ごとに積み上げ、選手ごとの台帳として Excel の別シート「U-21最新試合」に出力します。台帳は `u21_history.db`（SQLite）に保存されます。実行のたびにまだ取り込んでいない試合だけを取得するので、実行の間隔が空いても試合の取りこぼしはありません。

各サイトでの判定方法の違い（ベンチ外を判別できる大会とできない大会など）は、操作手順書の「大会による判定精度の違い」にまとめています。

## 使い方（exe 版）

[Releases](https://github.com/orcasnow/LoanWatch/releases) から最新版の `LoanWatch.exe` をダウンロードします。リポジトリの `players.json` と `u21_clubs.json` を同じフォルダに置き、`LoanWatch.exe` を起動します。Python のインストールは不要です。

画面の操作、選手の追加・編集、出力ファイルの見方は [LoanWatch 操作手順書](LoanWatch%20操作手順書.md)（[PDF](LoanWatch%20操作手順書.pdf)）にあります。

## ソースから実行する

Python 3.15 で動作を確認しています。Python 3.12・3.13 で動かす場合は、pandas 2.3 系に固定した [requirements-313.txt](requirements-313.txt) を使ってください。

```powershell
py -3.15 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m playwright install chromium
```

GUI で実行する場合は次のとおりです。

```powershell
python LoanWatchGui.py      # 実行設定とログの画面
python PlayersEditor.py     # players.json の編集画面だけを開く
```

コマンドラインで実行する場合は次のとおりです。

```powershell
python LoanWatchMain.py                          # players.json の全選手を取得
python LoanWatchMain.py --only 貴田遼河 --debug   # 1人だけ試す
python LoanWatchMain.py --source espn            # 取得元を絞る
python LoanWatchMain.py --u21-club 33186 33189   # U-21シートの対象クラブ
python LoanWatchMain.py --out loan_2026-11-21    # 出力ファイル名(拡張子なし)
```

主なオプションは次のとおりです。全オプションは `--help` で確認できます。

| オプション | 内容 |
| --- | --- |
| `--year` | 対象シーズン（既定：2026） |
| `--only 名前 ...` | 指定した選手だけ取得する |
| `--source` | 指定した取得元の選手だけ取得する（`espn` / `jleague` / `football-lab` / `transfermarkt` / `besoccer` / `manual`） |
| `--no-fallback` | Football-LAB が未更新でもＪリーグ公式に切り替えない |
| `--no-besoccer` | BeSoccer からリザーブチームの行を取得しない |
| `--u21-club` / `--no-u21-team` | U-21 シートの対象クラブを指定する／U-21 シートを作らない |
| `--no-excel` / `--no-csv` | Excel／CSV を出力しない |
| `--stale-days` | この日数より古い試合しか取れなかった選手を、最後に一覧で表示する（既定：14） |
| `--headed` / `--save-html` / `--debug` | 調査用。ブラウザ画面を表示する／取得した HTML を保存する／表の構造を表示する |

## 設定ファイル

`players.json` には対象選手を登録します。選手の追加・編集は GUI の「選手設定を編集」からも行えます。Jリーグの選手は次のように書きます。

```json
{
  "players": [
    {
      "name_ja": "河面旺成",
      "name_en": "Akinari KAWAZURA",
      "club": "水戸ホーリーホック",
      "source": "football-lab",
      "url": "https://www.football-lab.jp/mito/report",
      "jleague_url": "https://www.jleague.jp/player/1603777/"
    }
  ]
}
```

`source` には取得元を指定します。ESPN の選手は `espn_id` などを、リザーブチームの行も取得したい選手は `besoccer` を追加で書きます。書き方は操作手順書の5章と6章を参照してください。

`u21_clubs.json` には U-21 Ｊリーグの設定を書きます。対象にできるクラブの一覧、既定で取得するクラブ、開幕月、日程ページの URL です。順位決定戦が別の大会として扱われた場合も、`schedules` に日程ページを1件足せば同じ台帳に取り込めます。

## 出力

既定では `loan_players_latest.xlsx` と `loan_players_latest.csv` を作成します。U-21 シートを作った場合は `loan_players_latest_u21.csv` も作成します。Excel では、「チーム最新試合日付」の列からその試合のページに移動できます。

## exe のビルド

PyInstaller でビルドします。ビルドに使う venv と出力先は、Dropbox などの同期フォルダの外に置いてください。同期ソフトがファイルをロックするため、ビルドが WinError 32 で失敗します。手順は [requirements.txt](requirements.txt) の冒頭に書いてあります。

```powershell
$b = "$env:LOCALAPPDATA\LoanWatch-build"
py -3.15 -m venv "$b\venv"
& "$b\venv\Scripts\python.exe" -m pip install -r requirements.txt
& "$b\venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean `
    --distpath "$b\dist" --workpath "$b\build" .\LoanWatch.spec
& "$b\dist\LoanWatch.exe" --loanwatch-cli --playwright-smoke-test
```

exe には `%LOCALAPPDATA%\ms-playwright` の Chromium が同梱されるので、事前に `playwright install chromium` を実行しておいてください。

## ファイル構成

| ファイル | 内容 |
| --- | --- |
| `LoanWatchMain.py` | 取得と判定の本体（コマンドライン版） |
| `LoanWatchGui.py` | GUI。exe の起動点でもある |
| `PlayersEditor.py` | `players.json` の編集画面 |
| `u21_clubs.py` | `u21_clubs.json` の読み込み |
| `version.py` | バージョン番号と変更履歴 |
| `LoanWatch.spec` | PyInstaller の設定 |
| `LoanWatch_Colab.ipynb` | Google Colab で実行するためのノートブック |
| `espn_status.py`、`jleague_search.py`、`Transfermarkt-checker.py` | 各サイトの取得方法を調べるときに使った補助スクリプト |

## 注意事項

- Windows 専用です。
- 各サイトのページを取得して表を読み取っています。サイトの構成が変わると、正しく取得できなくなることがあります。
- 取得先のサイトに負荷をかけないよう、アクセスの間隔を空けています（`--sleep`、既定 2 秒）。短い間隔で繰り返し実行するのは控えてください。BeSoccer などは連続してアクセスするとボット判定でブロックされます。
- 出場履歴に載らない試合（ベンチ入りのみ・ベンチ外）は、取得元によっては判別できません。その場合は備考欄に「ベンチ外は判別不可」と出ます。

## ライセンス

Copyright 2026 Yukinobu Yoshihara

[Apache License 2.0](LICENSE) で公開しています。
