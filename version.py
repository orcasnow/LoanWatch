"""アプリのバージョン

ここだけを書き換える。LoanWatchMain.py(CUI)・LoanWatchGui.py(GUI)・
PlayersEditor.py(選手設定画面)が、この値を読んで表示する。

  3.1    ルヴァンカップ対応(Jリーグ公式の出場履歴から、天皇杯と同様に
         選手ごとの行を追加)
  3.0    BeSoccerの取得をインストール済みGoogle Chromeへ変更(永続プロファイル)。
         Chromeが無い場合は同梱Chromiumへフォールバック。--debug時に診断情報を保存。
         画面と実行ログにバージョンを表示。
  2.2    BeSoccer対応(アルヘンティノス・ジュニアーズIIなどリザーブチームの行)、
         ボット判定時のキャッシュ利用、Google Colab版の追加
  2.1    GUI版
"""

APP_VERSION = "3.1"
