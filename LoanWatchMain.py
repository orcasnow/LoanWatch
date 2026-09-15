"""LoanWatch - 名古屋グランパス レンタル移籍選手の最新試合 出場ステータス一覧

出力: 選手ごと1行 (先発 / 途中出場 / ベンチ入りのみ / ベンチ外 / 試合なし)
      ただしトップチームのほかに U-21 Ｊリーグ・ルヴァンカップ・天皇杯にも
      今季の記録がある選手は、大会ごとに1行ずつ出力する。
      players.json に besoccer 設定を書いた選手は、ESPN・Football-LAB に
      載らないチーム(アルヘンティノス・ジュニアーズIIなどのリザーブチーム)の
      行も BeSoccer から取得して追加する。

準備:
    pip install playwright requests beautifulsoup4 lxml html5lib pandas openpyxl
    playwright install chromium

実行:
    python LoanWatchMain.py
    python LoanWatchMain.py --year 2026 --debug
    python LoanWatchMain.py --headed --save-html          # 画面表示＋HTML保存
    python LoanWatchMain.py --only 貴田遼河 --debug        # 1人だけ試す
    python LoanWatchMain.py --source espn --debug         # ESPN対象の2人だけ
    python LoanWatchMain.py --no-besoccer                 # BeSoccerの追加行を出さない
    python LoanWatchMain.py --u21-club 33189 33186        # U-21シートの対象クラブ
    python LoanWatchMain.py --no-u21-team                 # U-21シートを作らない

GUI版:
    python LoanWatchGui.py            # 実行設定・ログ画面＋「選手設定を編集」ボタン
    python PlayersEditor.py           # players.json 編集画面を単体起動

U-21 Ｊリーグ(チーム単位・台帳方式):
    レンタル選手とは別に、指定したU-21クラブの出場記録を積み上げ、選手ごとに
    1行でExcelの別シート「U-21最新試合」へ出力する。行は
      ・すでに台帳にある選手が次の試合に関与 → その行を最新の内容で上書き
      ・台帳に無い選手                      → 行を追記
      ・その試合に関与しなかった選手        → 前回までの内容のまま残す
    という動き方をする。台帳は u21_history.db(SQLite)に貯め、実行のたびに
    未取得の試合だけを取りに行く。実行間隔を空けても、その間の試合はまとめて
    反映される。
    対象クラブ・開幕月・日程の取得元は u21_clubs.json で管理し、クラブは
    GUIのプルダウンまたは --u21-club で選ぶ(既定は名古屋グランパス)。
    メンバー表に載らないベンチ外の選手は、この方式では出力されない。

出力列:
    日付 / 経過 を先頭寄りに置き、いつの試合かがひと目で分かるようにしている。
    リーグ名 はその試合の大会名(明治安田Ｊ１ / U-21 Ｊリーグ / ルヴァンカップ /
    天皇杯 / ESPNは英語表記)。
    ベンチ外・ベンチ入りのみは大会名を特定できないため、Jリーグ公式の
    所属リーグで代用し、備考に「リーグ名は所属リーグから推定」と入れる。
    対戦相手 は相手クラブ名のみ、試合結果 は対象クラブ視点の勝敗・スコア。
    警告/退場 は「警告数/退場数」で表示する。
    --stale-days で指定した日数より古い試合しか取れなかった選手は末尾に列挙する。
"""

import argparse
import asyncio
import os
import re
import sqlite3
import sys
import time
import traceback
import hashlib
import json
import platform
import unicodedata
from importlib import metadata as importlib_metadata
from datetime import date, datetime
from io import StringIO
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

import u21_clubs
from version import APP_VERSION


def configure_system_cert_store():
    if getattr(sys, "frozen", False):
        return
    try:
        import truststore
    except ImportError:
        return
    truststore.inject_into_ssl()


configure_system_cert_store()

import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup


def configure_playwright_browsers():
    if getattr(sys, "frozen", False):
        bundle_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(bundle_dir / "ms-playwright"))


configure_playwright_browsers()

from playwright.async_api import async_playwright


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# =========================================================
# 定数
# =========================================================
STARTER, SUBBED_IN, BENCH, OUT_SQUAD, NO_MATCH = \
    "先発", "途中出場", "ベンチ入りのみ", "ベンチ外", "試合なし"

# 出場時間の計算に使う試合の長さ。アディショナルタイムは扱わない
# (Jリーグ公式の出場時間も同じ数え方)。
FULL_MATCH_MINUTES, HALF_MINUTES = 90, 45

# 行テキストから状態を判定するためのキーワード
KW_BENCH     = ["ベンチ入り", "ベンチ", "控え", "出場なし", "未出場"]
KW_OUT_SQUAD = ["招集外", "メンバー外", "登録外", "負傷", "けが", "出場停止", "累積警告"]
KW_SUB_IN    = ["途中出場", "交代出場", "交代IN"]
# 出場時間が入っている列名の候補
KW_MIN_COL   = ["出場時間", "出場", "時間", "分"]
# 交代INが入っている列名の候補
KW_SUBIN_COL = ["交代IN", "投入", "途中出場"]
# 中止・延期を示す本文中の文言
RE_CANCELLED = re.compile(r"(中止|延期|試合データがありません|データはありません)")

COLUMNS = ["選手", "クラブ", "リーグ名", "日付", "経過", "チーム最新試合日付", "ステータス",
           "出場時間", "対戦相手", "試合結果", "ゴール数", "警告/退場", "備考", "URL"]
RE_GAME_URL = re.compile(r"/soccer/match/_/gameId/(\d+)")

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Referer": "https://www.google.com/",
}

# 選手マスタは players.json で管理する。
# source: "espn" | "transfermarkt" | "football-lab" | "besoccer" | "manual"
# espn_id は ESPN 選手URLの /id/<数字>/、jleague_url は出場時間補完用。
# u21_club_id は、所属クラブが U-21 Ｊリーグに参加している選手だけに書く。
# 日程ページ https://www.jleague.jp/u-21/match/ でクラブを絞ったときの
# URL に付く &club=<ここ> の数字(例: ファジアーノ岡山 = 33189)。
# 書いておくと試合ページのメンバー表まで見に行くので、出場しなかった試合も
# ベンチ入りのみ/ベンチ外として判別できる。未記入なら出場履歴だけを見る。
# espn_team_id を書いておくと、チームの結果ページ
#   (https://www.espn.com/soccer/team/results/_/id/<espn_team_id>)
# と突き合わせ、選手ページの Last 5 Matches がチームの最新試合より
# 古い場合に備考へ「要確認」を出す(選手ページの反映遅れ検知用)。
#
# besoccer は source と関係なく足せる任意項目で、ESPN・Football-LAB に
# 載らないチーム(リザーブチームなど)の行を追加するために使う:
#   "besoccer": {
#     "url": "https://www.besoccer.com/player/matches/ryoga-kida-3237854",
#     "teams": [
#       {"logo":  "torneo-proyeccion",              ← リーグロゴのファイル名
#        "team":  "Argentinos Juniors II",          ← BeSoccer上のチーム名
#        "club":  "AAアルヘンティノス・ジュニアーズII", ← 出力の クラブ 列(任意)
#        "league": "トルネオ・プロジェクシオン"}       ← 出力の リーグ名 列(任意)
#     ]
#   }
# logo は選手ページの試合一覧で行頭に出るロゴ画像
#   https://cdn.resfu.com/media/img/league_logos/<logo>.png?size=60x&lossy=1
# のファイル名部分。これで大会(=どのチームでの出場か)を見分ける。
DEFAULT_PLAYERS_CONFIG = app_dir() / "players.json"

CFG = {}  # main() で argparse の結果を格納


def load_players(config_path):
    with Path(config_path).open("r", encoding="utf-8") as f:
        config = json.load(f)

    players = config.get("players") if isinstance(config, dict) else config
    if not isinstance(players, list):
        raise ValueError("選手設定ファイルは players 配列、または選手配列そのものを指定してください")

    required = {"name_ja", "club", "source", "url"}
    for i, player in enumerate(players, start=1):
        if not isinstance(player, dict):
            raise ValueError(f"選手設定の {i} 件目がオブジェクトではありません")
        missing = sorted(required - set(player))
        if missing:
            raise ValueError(f"選手設定の {i} 件目に必須項目がありません: {', '.join(missing)}")
        validate_u21_club_id(player, i)
        validate_besoccer(player, i)
    return players


def validate_u21_club_id(player, i):
    """任意項目 u21_club_id が数字かどうかを確かめる"""
    value = str(player.get("u21_club_id", "")).strip()
    if value and not value.isdigit():
        raise ValueError(f"選手設定の {i} 件目の u21_club_id は数字で書いてください "
                         f"(U-21日程ページの &club= の値): {value!r}")


def validate_besoccer(player, i):
    """任意項目 besoccer の形を確かめる(綴り違いを実行前に気付けるように)"""
    conf = player.get("besoccer")
    if conf is None:
        return
    where = f"選手設定の {i} 件目の besoccer"
    if not isinstance(conf, dict):
        raise ValueError(f"{where} はオブジェクトで書いてください")
    if not str(conf.get("url", "")).strip():
        raise ValueError(f"{where} に url がありません")
    teams = conf.get("teams")
    if not isinstance(teams, list) or not teams:
        raise ValueError(f"{where} の teams は1件以上の配列で書いてください")
    for j, team in enumerate(teams, start=1):
        if not isinstance(team, dict):
            raise ValueError(f"{where}.teams の {j} 件目がオブジェクトではありません")
        lacking = [k for k in ("logo", "team") if not str(team.get(k, "")).strip()]
        if lacking:
            raise ValueError(f"{where}.teams の {j} 件目に必須項目がありません: "
                             f"{', '.join(lacking)}")


# =========================================================
# 例外・ログ
# =========================================================
class ScrapeError(Exception):
    """人が読んで原因がわかるメッセージを持つ例外

    kind は呼び出し側が分岐に使う識別子。
      "no_matches" … そのシーズンの試合がまだ1件も無い(中止・延期など)
    """

    def __init__(self, message, hint="", detail="", kind=""):
        super().__init__(message)
        self.message, self.hint, self.detail = message, hint, detail
        self.kind = kind

    def summary(self):
        return self.message + (f" / 対処: {self.hint}" if self.hint else "")


def log(msg, level="INFO"):
    print(f"  [{level}] {msg}")


def debug(msg):
    if CFG.get("debug"):
        log(msg, "DEBUG")


# =========================================================
# 共通ユーティリティ
# =========================================================
def norm(s):
    """名前照合用。全角半角を揃え、空白を除去する"""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(s)))


def row_text(row):
    return " ".join(str(v) for v in row.values if pd.notna(v))


def _to_minutes(text):
    """'90'' や '4.785'' のような表記を整数に変換する(桁区切りドットに対応)"""
    m = re.search(r"\d{1,3}(?:[.,]\d{3})*", str(text))
    if not m:
        return None
    return int(m.group(0).replace(".", "").replace(",", ""))


def parse_minutes(row, columns):
    """行から出場時間(分)を取り出す。見つからなければ None"""
    for col, val in zip(columns, row.values):
        if any(k in str(col) for k in KW_MIN_COL):
            minutes = _to_minutes(val)
            if minutes is not None:
                return minutes
    m = re.search(r"(\d{1,3})\s*(?:分|')", row_text(row))
    return int(m.group(1)) if m else None


def has_sub_in(row, columns):
    """交代IN列に値が入っているか"""
    for col, val in zip(columns, row.values):
        if any(k in str(col) for k in KW_SUBIN_COL):
            if str(val).strip() not in ("nan", "-", "", "None"):
                return True
    return any(k in row_text(row) for k in KW_SUB_IN)


def judge_status(row, columns):
    """(ステータス, 出場時間) を判定する"""
    text = row_text(row)
    minutes = parse_minutes(row, columns)
    if any(k in text for k in KW_OUT_SQUAD):
        return OUT_SQUAD, None
    if minutes and minutes > 0:
        return (SUBBED_IN if has_sub_in(row, columns) else STARTER), minutes
    if any(k in text for k in KW_BENCH):
        return BENCH, None
    return BENCH, None


def days_ago(date_str):
    """'2026-03-12' から経過日数を出す。解釈できなければ None"""
    try:
        d = datetime.strptime(str(date_str), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None
    return (date.today() - d).days


def format_elapsed(date_str):
    n = days_ago(date_str)
    if n is None:
        return ""
    if n == 0:
        return "本日"
    if n == 1:
        return "昨日"
    return f"{n}日前"


def result(player, status, match="", date="", minutes=None, url=None, note="",
           match_result="", goals="", cards="", league="",
           team_latest_date="", team_latest_url=""):
    return {"選手": player["name_ja"], "クラブ": player["club"],
            "リーグ名": league,
            "日付": date, "経過": format_elapsed(date),
            "チーム最新試合日付": team_latest_date,
            "ステータス": status,
            "出場時間": f"{minutes}分" if minutes else "",
            "対戦相手": match,
            "試合結果": match_result,
            "ゴール数": goals,
            "警告/退場": cards,
            "備考": note, "URL": url or player.get("url", ""),
            "_team_latest_url": team_latest_url}


def blank(player, status, note=""):
    return result(player, status, note=note)


def describe(html, url):
    """パース失敗時に何が返ってきたかを説明する"""
    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    classes = sorted({" ".join(t.get("class", [])) or "(class無)" for t in tables})
    lines = [f"URL: {url}",
             f"HTML長: {len(html)}文字 / "
             f"title: {soup.title.get_text(strip=True) if soup.title else '(なし)'}",
             f"table: {len(tables)}個 class={classes if classes else 'なし'}"]
    if not tables:
        lines.append("→ tableが0個。JavaScript描画かブロックページの可能性")
    return "\n     ".join(lines)


def save_html(html, name):
    if not CFG.get("save_html"):
        return
    fn = re.sub(r"\W+", "_", name)[:80] + ".html"
    with open(fn, "w", encoding="utf-8") as f:
        f.write(html)
    log(f"HTML保存: {fn}")


def read_tables_any(html):
    """pandas.read_html を parser の有無に左右されず実行する

    既定の read_html は lxml で失敗すると html5lib に切り替えるため、
    html5lib 未導入の環境では ImportError になる。lxml で正規化してから
    読み直すことで、追加パッケージ無しでも動くようにする。
    """
    try:
        return pd.read_html(StringIO(html), flavor="lxml")
    except ValueError:
        pass
    except ImportError:
        pass
    except Exception as e:
        debug(f"read_html(lxml)が失敗: {type(e).__name__}: {e}")
    # 壊れたHTMLでも BeautifulSoup で組み直せば読めることがある
    try:
        fixed = str(BeautifulSoup(html, "lxml"))
        return pd.read_html(StringIO(fixed), flavor="lxml")
    except Exception as e:
        debug(f"整形後も read_html に失敗: {type(e).__name__}: {e}")
        return []


def tables_from(html, selector="table.items"):
    """指定セレクタのテーブルを DataFrame のリストにする"""
    soup = BeautifulSoup(html, "lxml")
    nodes = soup.select(selector) or soup.find_all("table")
    frames = []
    for t in nodes:
        frames += read_tables_any(str(t))
    return frames


def clean_columns(df):
    df.columns = [str(c) for c in df.columns]
    return df.loc[:, [c for c in df.columns if not c.startswith("Unnamed")]]


def find_col(columns, keywords):
    return next((c for c in columns if any(k in str(c) for k in keywords)), None)


# =========================================================
# Transfermarkt (Playwright でレンダリング)
# =========================================================
async def _render(url, headless=True, wait_ms=6000, wait_selector="table.items",
                  locale="ja-JP", scroll=False):
    """PlaywrightでページをレンダリングしてHTMLを返す。

    サイト固有のUser-AgentやClient Hintsの偽装は行わず、Playwrightが
    起動したブラウザ自身の値をそのまま利用する。
    """
    headers = {"Accept-Language": f"{locale},en;q=0.9"}
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=headless)
        ctx = await browser.new_context(
            locale=locale, timezone_id="Asia/Tokyo",
            viewport={"width": 1440, "height": 1000},
            extra_http_headers=headers)
        page = await ctx.new_page()
        try:
            resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            status = resp.status if resp else 0
            if status == 403:
                raise ScrapeError("HTTP 403 アクセス拒否",
                                  "IPがボット判定された可能性。回線・実行環境を変える", url)
            if status == 406:
                raise ScrapeError("HTTP 406 ボット判定でブロック",
                                  "時間をおいて再実行する。連続実行は控える", url)
            if status >= 400:
                raise ScrapeError(f"HTTP {status}", "URL・選手IDを確認", url)

            # Cookie同意バナーがコンテンツを覆うことがあるので閉じる
            for sel in ("button:has-text('同意')", "#onetrust-accept-btn-handler",
                        "button:has-text('I Accept')", "button:has-text('Accept')"):
                try:
                    await page.click(sel, timeout=1500)
                    debug(f"Cookie同意をクリック: {sel}")
                    break
                except Exception:
                    pass

            if wait_selector:
                try:
                    await page.wait_for_selector(wait_selector, timeout=wait_ms)
                except Exception:
                    log(f"待機セレクタが現れませんでした: {wait_selector}", "WARN")
            if scroll:
                await page.mouse.wheel(0, 3000)
                await page.wait_for_timeout(2000)
            return await page.content()
        finally:
            await browser.close()


def render(url, wait_selector="table.items", locale="ja-JP", scroll=False):
    """_render を同期的に呼ぶラッパ。

    DNSや接続の一時的な失敗は fetch() と同じようにリトライする。
    HTTP 403/406 などサイト側が意図して弾いた場合(ScrapeError)は繰り返さない。
    """
    last = None
    for attempt in range(1, CFG["retry"] + 1):
        try:
            return asyncio.run(_render(url, headless=not CFG["headed"],
                                       wait_selector=wait_selector,
                                       locale=locale, scroll=scroll))
        except ScrapeError:
            raise
        except Exception as e:
            last = e
            if "Executable doesn't exist" in str(e):
                # Playwrightのブラウザ本体がなければリトライしても解決しない。
                command = f'"{sys.executable}" -m playwright install chromium'
                raise ScrapeError(
                    "Playwright Chromiumが見つかりません",
                    f"次のコマンドを実行: {command}",
                    url,
                ) from e
            if attempt < CFG["retry"]:
                log(f"ブラウザ取得に失敗({type(e).__name__}) → {attempt}回目リトライ", "WARN")
                time.sleep(CFG["sleep"] * 2)
    raise ScrapeError(f"ブラウザ取得に失敗: {type(last).__name__}: {last}",
                      f'"{sys.executable}" -m playwright install chromium を実行済みか確認', url)


async def _playwright_smoke_test():
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        page = await browser.new_page()
        try:
            await page.set_content("<html><title>LoanWatch</title><body>OK</body></html>")
            title = await page.title()
            body = await page.text_content("body")
            return title == "LoanWatch" and body == "OK"
        finally:
            await browser.close()


def playwright_smoke_test():
    return asyncio.run(_playwright_smoke_test())


def tm_season_links(html, base_url, year):
    """概要ページから、表示シーズンが year の大会別 試合ログリンクを集める

    Transfermarkt は表示上の 2026 が URL 上 saison/2025 になることがある。
    行に書かれたシーズン文字列で判定することでこのずれを吸収する。
    """
    soup = BeautifulSoup(html, "lxml")
    table = soup.select_one("table.items")
    if table is None:
        raise ScrapeError("成績テーブルが見つからない",
                          "URL末尾に /plus/1 が付いているか確認", describe(html, base_url))

    y4, y2 = str(year), str(year)[-2:]
    links, seasons_seen = [], []
    for tr in table.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        if not cells:
            continue
        season = cells[0].get_text(" ", strip=True)
        if season:
            seasons_seen.append(season)
        # 「2026」「25/26」どちらの表記にも対応する
        if y4 not in season and y2 not in season.replace("/", ""):
            continue
        comp = cells[2].get_text(" ", strip=True) if len(cells) > 2 else ""
        for a in tr.select("a[href]"):
            href = a["href"]
            if "leistungsdatendetails" in href and "wettbewerb" in href:
                links.append((season, comp, urljoin(base_url, href)))
                break

    if not links:
        raise ScrapeError(f"{year}年のシーズン行が0件",
                          f"--year を変更する。ページ内のシーズン表記: "
                          f"{sorted(set(s for s in seasons_seen if s))}", base_url)
    return links


def tm_match_rows(html, url):
    """試合ログページを DataFrame にする。日付列が無ければ None(=集計表)"""
    frames = tables_from(html)
    if not frames:
        return None
    df = clean_columns(pd.concat(frames, ignore_index=True))
    debug(f"列: {list(df.columns)}")
    if CFG.get("debug"):
        print(df.head(3).to_string())

    date_col = find_col(df.columns, ["日付", "日時", "Datum", "Date"])
    if date_col is None:
        return None
    df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    return df.dropna(subset=["_date"])


def scrape_transfermarkt(player):
    overview = render(player["url"])
    save_html(overview, player["name_ja"] + "_overview")

    links = tm_season_links(overview, player["url"], CFG["year"])
    log(f"{CFG['year']}年の大会 {len(links)}件: {', '.join(c for _, c, _ in links)}")

    collected = []
    for _season, comp, url in links:
        try:
            html = render(url)
        except ScrapeError as e:
            log(f"{comp}: {e.summary()}", "WARN")
            continue
        save_html(html, f"{player['name_ja']}_{comp}")
        df = tm_match_rows(html, url)
        if df is None:
            log(f"{comp}: 日付列なし(集計表が返った)", "WARN")
            continue
        if df.empty:
            log(f"{comp}: 試合行なし", "WARN")
            continue
        df["_comp"] = comp
        df["_url"] = url
        collected.append(df)
        time.sleep(CFG["sleep"])

    if not collected:
        return blank(player, NO_MATCH, f"{CFG['year']}年の試合行を取得できない")

    all_df = pd.concat(collected, ignore_index=True).sort_values("_date")
    latest = all_df.iloc[-1]
    status, minutes = judge_status(latest, all_df.columns)

    opp_col = find_col(all_df.columns, ["対戦", "相手", "Gegner"])
    label = str(latest[opp_col]).strip() if opp_col else ""
    return result(player, status, match=label,
                  date=latest["_date"].strftime("%Y-%m-%d"),
                  minutes=minutes, url=latest["_url"])


# =========================================================
# Football-Lab (requests)
# =========================================================
SESSION = requests.Session()
SESSION.headers.update(HEADERS)


def session_get(url):
    try:
        return SESSION.get(url, timeout=CFG["timeout"])
    except requests.exceptions.SSLError:
        log("SSL証明書検証に失敗したため、検証を緩和して再取得します", "WARN")
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        return SESSION.get(url, timeout=CFG["timeout"], verify=False)


def fetch(url):
    last = None
    for attempt in range(1, CFG["retry"] + 1):
        try:
            r = session_get(url)
        except requests.Timeout:
            last = ScrapeError(f"タイムアウト({CFG['timeout']}秒)", "--timeout を延ばす", url)
        except requests.ConnectionError as e:
            last = ScrapeError("接続失敗(DNS/ネットワーク)", "URLと通信状態を確認", f"{url}\n{e}")
        else:
            if r.status_code == 404:
                raise ScrapeError("HTTP 404 ページなし", "URLの綴りを確認", url)
            if r.status_code == 429:
                last = ScrapeError("HTTP 429 レート制限", "--sleep を増やす", url)
            elif r.status_code == 403:
                last = ScrapeError("HTTP 403 アクセス拒否", "実行環境・回線を変える", url)
            elif r.status_code >= 500:
                last = ScrapeError(f"HTTP {r.status_code} サーバ側エラー", "時間をおいて再実行", url)
            else:
                # Content-Type に charset があればそれが正解。apparent_encoding
                # (charset_normalizerの推定)で上書きすると、まれに ptcp154 などの
                # 全バイトに文字が割り当たった文字コードが選ばれ、例外も出ないまま
                # 日本語が化ける。宣言が無いサイトのときだけ推測に頼る。
                if "charset" not in r.headers.get("Content-Type", "").lower():
                    r.encoding = r.apparent_encoding or "utf-8"
                html = r.text
                if len(html) < 2000:
                    raise ScrapeError(f"応答が異常に短い({len(html)}文字)",
                                      "--save-html で中身を確認", html[:300])
                time.sleep(CFG["sleep"])
                return html
        if attempt < CFG["retry"]:
            log(f"{last.message} → {attempt}回目リトライ", "WARN")
            time.sleep(CFG["sleep"] * 2)
    raise last


def fl_latest_match(report_url):
    """レポート一覧から、指定年のうち日付が最も新しい試合を返す

    戻り値: (対戦相手, 試合URL, 日付文字列, 今季初戦の日付)
    今季初戦の日付は、全シーズン分が入っているJリーグ公式の出場履歴から
    「今季の」U-21 Ｊリーグの記録だけを取り出すときの基準に使う。
    """
    html = fetch(report_url)
    save_html(html, urlparse(report_url).path + "_report")
    soup = BeautifulSoup(html, "lxml")
    team_path = urlparse(report_url).path.rstrip("/")

    cands, years_seen = [], set()
    for a in soup.select("a[href]"):
        href = a["href"]
        if team_path not in href or "year=" not in href:
            continue
        q = parse_qs(urlparse(href).query)
        try:
            y = int(q.get("year", [0])[0])
            mo = int(q.get("month", [0])[0])
            d = int(q.get("date", [0])[0])
        except (ValueError, IndexError):
            continue
        years_seen.add(y)
        if y == CFG["year"]:
            cands.append(((y, mo, d), a.get_text(" ", strip=True),
                          urljoin(report_url, href)))

    if not cands:
        raise ScrapeError(
            f"{CFG['year']}年の試合リンクが0件",
            f"今季未消化なら正常。ページ内の年: {sorted(years_seen) or '(なし)'}",
            report_url, kind="no_matches")

    cands.sort(key=lambda x: x[0])
    (y, mo, d), label, url = cands[-1]
    debug(f"{CFG['year']}年の試合リンク {len(cands)}件、最新 {y}-{mo:02d}-{d:02d}")
    try:
        season_start = date(*cands[0][0])
    except ValueError:
        season_start = None
    return fl_opponent_from_label(label), url, f"{y}-{mo:02d}-{d:02d}", season_start


def fl_opponent_from_label(label):
    """Football-LABのリンク文言から日付・節を除き、対戦相手だけを残す。"""
    text = str(label).strip()
    text = re.sub(r"^\d{1,2}/\d{1,2}(?:\s+\d{1,2}\.\d{1,2})?", "", text).strip()
    return re.sub(r"^\d+節\s*", "", text).strip()


def fl_match_card(html):
    """個別レポートのページ見出しから対戦相手を取り出す。

    Football-LABの一覧ではリンク文字列が日付だけのクラブがある一方、
    個別ページの title / パンくずには ``... | 8月8日 vs 相手 | ...`` が
    一貫して含まれる。ここを優先すれば一覧側の表示形式に依存しない。
    """
    soup = BeautifulSoup(html, "lxml")
    candidates = []
    if soup.title:
        candidates.append(soup.title.get_text(" ", strip=True))
    for node in soup.select("meta[property='og:title'], meta[name='title'], h1"):
        candidates.append(node.get("content", "") or node.get_text(" ", strip=True))

    for text in candidates:
        m = re.search(r"\bvs\.?\s*(.+?)(?:\s*[|｜]\s*|$)", text, re.I)
        if m:
            opponent = m.group(1).strip()
            # title末尾のサイト名等が混ざった場合は除外する
            opponent = re.sub(r"\s+(?:データによって.*|Football\s*LAB.*)$", "", opponent,
                              flags=re.I).strip()
            if opponent:
                return {"opponent": opponent}
    return {"opponent": ""}


RE_MIN_ANY = re.compile(r"(\d{1,3})\s*(?:分|')")


def fl_member_status(html, name_ja):
    """メンバー表から出場区分を読む

    実ページの構造:
        | Starting Member |
        | MF | 3  | <img alt="←OUT"> |  | 萩野 滉大 |
        | Substitutes     |
        | MF | 14 | <img alt="→IN">  |  | 鈴木 陽人 |
        | MF | 20 |                   |  | 金本 信利 |

    戻り値:
        None                     メンバー表そのものが無い(試合前など)
        section=None             表はあるが名前が無い = ベンチ外
        section='starting'/'sub' + has_in / has_out / minute
    """
    soup = BeautifulSoup(html, "lxml")
    target = norm(name_ja)
    sections = []
    for table in soup.find_all("table"):
        head = table.get_text(" ", strip=True)[:80]
        if "Starting Member" in head:
            sections.append(("starting", table))
        elif "Substitutes" in head:
            sections.append(("sub", table))
    if not sections:
        return None

    for section, table in sections:
        for tr in table.find_all("tr"):
            cells = tr.find_all(["th", "td"])
            if not cells:
                continue
            text = " ".join(c.get_text(" ", strip=True) for c in cells)
            if target not in norm(text):
                continue
            # 交代矢印は img の alt / title に入っている
            marks = " ".join(f"{img.get('alt', '')} {img.get('title', '')}"
                             for img in tr.find_all("img"))
            mark_text = f"{marks} {text}"
            has_in = "→IN" in mark_text.replace(" ", "")
            has_out = "←OUT" in mark_text.replace(" ", "")
            m = RE_MIN_ANY.search(mark_text)
            minute = int(m.group(1)) if m else None
            debug(f"メンバー表ヒット: 区分={section} IN={has_in} OUT={has_out} "
                  f"分={minute} | {text[:50]}")
            return {"section": section, "has_in": has_in,
                    "has_out": has_out, "minute": minute}
    return {"section": None, "has_in": False, "has_out": False, "minute": None}


def _footballlab_core(player, latest):
    """Football-LAB から 最新試合 の出場区分 を取る

    メンバー表は試合当日から載る(後日更新されるのはタイムライン・シュート・CBP)。
    ベンチ外かベンチ入りかは、この表に名前があるかどうかでしか判別できない。
    戻り値の _fl キーに判定材料を入れ、呼び出し側で出場時間を補完する。
    latest は呼び出し側で取得済みの fl_latest_match() の戻り値
    (label, match_url, date, season_start) — チーム最新試合日付列と同じ取得を
    二重に行わないよう、ここでは受け取るだけにする。
    """
    label, match_url, date = latest[:3]
    html = fetch(match_url)
    save_html(html, f"{player['name_ja']}_fl_match")
    soup = BeautifulSoup(html, "lxml")
    card = fl_match_card(html)
    if card["opponent"]:
        label = card["opponent"]
    else:
        debug("個別レポートから対戦相手を取得できず、一覧リンクの値を使用")

    if RE_CANCELLED.search(soup.get_text(" ", strip=True)):
        row = result(player, NO_MATCH, match=label, date=date, url=match_url,
                     note="ページに中止・延期の記載あり")
        row["_fl"] = None
        return row

    member = fl_member_status(html, player["name_ja"])
    if member is None:
        row = result(player, NO_MATCH, match=label, date=date, url=match_url,
                     note="メンバー表が未掲載")
        row["_fl"] = None
        return row

    if member["section"] is None:
        # メンバー表はあるが名前が無い = ベンチ外(FLでしか判別できない)
        row = result(player, OUT_SQUAD, match=label, date=date, url=match_url)
        row["_fl"] = member
        return row

    if member["section"] == "starting":
        status = STARTER
    elif member["has_in"]:
        status = SUBBED_IN
    else:
        status = BENCH

    row = result(player, status, match=label, date=date, url=match_url)
    row["_fl"] = member
    return row


# =========================================================
# ESPN (Playwright + 埋め込みJSON)
#
# ESPNのページには window['__espnfitt__'] に描画元のJSONが埋まっている。
# 試合ページの実構造(2026年8月に確認):
#   page.content.gamepackage.lineUps[i].starters / substitutes / unused
#   page.content.gamepackage.lineUps[i].playersMap[<id>]
#       stats.subIns      交代出場した回数
#       plyrPlys          出来事(iconType: subIn / subOut, minute: "85'")
#       subbedOutPlayer   入れ替わった選手のID
#   frmtnPlc は途中出場でも 1〜11 を取るため先発判定には使えない。
# =========================================================
def extract_fitt(html):
    """ページに埋まっているESPNの状態JSONを取り出す。無ければ None"""
    m = re.search(r"window\[.__espnfitt__.\]\s*=\s*(\{.*?\});?\s*</script>", html, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError as e:
        debug(f"__espnfitt__ のJSON解析に失敗: {e}")
        return None


def find_dicts(node, predicate, path=""):
    """predicate(dict) が真になる辞書を再帰的に集める"""
    found = []
    if isinstance(node, dict):
        if predicate(node):
            found.append((path, node))
        for k, v in node.items():
            found += find_dicts(v, predicate, f"{path}.{k}")
    elif isinstance(node, list):
        for i, v in enumerate(node):
            found += find_dicts(v, predicate, f"{path}[{i}]")
    return found


def ids_in(container):
    """starters / substitutes / unused の中身を選手IDの集合にする"""
    out = set()
    if isinstance(container, dict):
        container = container.values()
    if not isinstance(container, (list, tuple, set)):
        return out
    for item in container:
        if isinstance(item, (str, int)):
            out.add(str(item))
        elif isinstance(item, dict):
            for k in ("id", "playerId", "athleteId"):
                if item.get(k):
                    out.add(str(item[k]))
                    break
    return out


def find_play(rec, icon_types):
    """plyrPlys から指定 iconType の出来事を探し、分を返す"""
    for play in rec.get("plyrPlys") or []:
        if not isinstance(play, dict):
            continue
        icon = str(play.get("iconType", "")).lower()
        if any(t.lower() in icon for t in icon_types):
            m = re.match(r"^(\d{1,3})", str(play.get("minute", "")).strip())
            return int(m.group(1)) if m else None
    return None


def stat_int(rec, key):
    try:
        return int(str((rec.get("stats") or {}).get(key, "")).strip())
    except (ValueError, TypeError):
        return 0


def resolve_name(player_id, players_map):
    """選手IDを表示名に変換する"""
    rec = (players_map or {}).get(str(player_id))
    if isinstance(rec, dict):
        for k in ("shrtNm", "name", "lstName"):
            if rec.get(k):
                num = rec.get("nmbr")
                return f"#{num} {rec[k]}" if num else str(rec[k])
    return f"ID {player_id}"


def espn_parse_date(text):
    """'Thu 3/12' を date にする。年が無いので当年基準で推定する"""
    m = re.search(r"\b(\d{1,2})/(\d{1,2})\b", text)
    if not m:
        return None
    today = date.today()
    mo, d = int(m.group(1)), int(m.group(2))
    try:
        guess = date(today.year, mo, d)
    except ValueError:
        return None
    return date(today.year - 1, mo, d) if guess > today else guess


def espn_latest_match(html, base_url):
    """Last 5 Matches の全行から、日付が最も新しい試合を返す

    ESPNの表は昇順のことがあるため最上位行を無条件には採用しない。
    """
    soup = BeautifulSoup(html, "lxml")
    section = None
    for tag in soup.find_all(string=re.compile(r"Last\s*5\s*Matches", re.I)):
        section = tag.find_parent(
            lambda t: t.name in ("section", "div") and t.find("table"))
        if section:
            break
    if section is None:
        debug("Last 5 Matches の見出しが見つからない。ページ全体から探す")
        section = soup

    table = section.find("table")
    if table is None:
        return None

    rows = []
    for tr in table.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        link = tr.find("a", href=RE_GAME_URL)
        if not cells or link is None:
            continue
        texts = [c.get_text(" ", strip=True) for c in cells]
        row_text_ = " | ".join(texts)
        url = urljoin(base_url, link["href"])
        app = next((t for t in texts
                    if re.fullmatch(r"(Started|Sub|Subbed On|Bench|-)", t, re.I)), "")
        opponent_cell = next((t for t in texts if re.match(r"^(?:vs\.?|@)\s+", t, re.I)), "")
        opponent = re.sub(r"^(?:vs\.?|@)\s+", "", opponent_cell, flags=re.I).strip()
        score_cell = next((t for t in texts
                           if re.fullmatch(r"[WDL]\s*\d+\s*-\s*\d+", t, re.I)), "")
        match_result = espn_format_result(score_cell, opponent_cell.startswith("@"))
        rows.append({"row": row_text_, "url": url,
                     "game_id": RE_GAME_URL.search(url).group(1),
                     "result": link.get_text(" ", strip=True),
                     "app": app, "date": espn_parse_date(row_text_),
                     "opponent": opponent, "match_result": match_result})
    if not rows:
        return None

    for i, r in enumerate(rows):
        debug(f"行{i}: date={r['date']} app={r['app'] or '-'} | {r['row'][:70]}")

    dated = [r for r in rows if r["date"]]
    if dated:
        latest = max(dated, key=lambda r: r["date"])
        if latest is not rows[0]:
            log("最上位行より新しい行があったので、そちらを採用しました", "WARN")
    else:
        log("日付を解釈できなかったので最上位行を使います", "WARN")
        latest = rows[0]
    return latest


def espn_format_result(score_cell, is_away):
    """ESPNの ``W 3-0`` を対象クラブ視点の ``勝 3-0`` に整形する。"""
    m = re.fullmatch(r"([WDL])\s*(\d+)\s*-\s*(\d+)", score_cell, re.I)
    if not m:
        return ""
    outcome, scored, conceded = m.groups()
    # ESPNのスコアはホーム-アウェー順。@ は対象クラブがアウェーを表す。
    if is_away:
        scored, conceded = conceded, scored
    outcome_ja = {"W": "勝", "D": "分", "L": "敗"}[outcome.upper()]
    return f"{outcome_ja} {scored}-{conceded}"


def espn_find_record(fitt, espn_id):
    """lineUps[*].playersMap[<id>] を探す"""
    try:
        lineups = fitt["page"]["content"]["gamepackage"]["lineUps"]
        for i, lu in enumerate(lineups):
            pmap = lu.get("playersMap") or {}
            if espn_id in pmap:
                return (f".page.content.gamepackage.lineUps[{i}].playersMap.{espn_id}",
                        pmap[espn_id], lu)
    except (KeyError, TypeError, IndexError):
        debug("既知のパスに playersMap が無い。再帰探索に切り替える")
    hits = find_dicts(fitt, lambda d: isinstance(d.get("playersMap"), dict)
                      and espn_id in d["playersMap"])
    if hits:
        path, lu = hits[0]
        return f"{path}.playersMap.{espn_id}", lu["playersMap"][espn_id], lu
    return None, None, None


def espn_league_name(fitt):
    """試合ページのJSONから大会名を取り出す

    ``page.content.gamepackage.league`` に "Argentine Liga Profesional de Fútbol"
    のような正式名称が入っている。無ければ短縮表記(gmStrp.nte)、
    リーグID("arg.1")の順に落とす。
    """
    try:
        gp = fitt["page"]["content"]["gamepackage"]
    except (KeyError, TypeError):
        debug("試合ページのJSONに gamepackage が無い")
        return ""
    for value in (gp.get("league"), (gp.get("gmStrp") or {}).get("nte"), gp.get("lg")):
        if value:
            return str(value).strip()
    return ""


def espn_judge(html, espn_id):
    """試合ページのJSONから出場状況を判定する"""
    fitt = extract_fitt(html)
    if fitt is None:
        raise ScrapeError("__espnfitt__ が見つからない",
                          "--headed で描画を確認する。ページ構造の変更かもしれない")

    league = espn_league_name(fitt)
    path, rec, lineup = espn_find_record(fitt, espn_id)
    if rec is None:
        debug("JSONに選手レコードが無い(メンバー入りしていない)")
        return {"status": OUT_SQUAD, "minutes": None, "entry": None,
                "exit": None, "swap": None, "goals": "", "cards": "",
                "league": league, "via": "JSON(レコード無し)"}

    debug(f"採用: {path} / keys={list(rec)}")
    lineup = lineup or {}
    pmap = lineup.get("playersMap") or {}
    starters = ids_in(lineup.get("starters"))
    subs = ids_in(lineup.get("substitutes"))
    unused = ids_in(lineup.get("unused"))
    sub_in_min = find_play(rec, ["subin"])
    sub_out_min = find_play(rec, ["subout"])
    sub_ins = stat_int(rec, "subIns")
    appearances = stat_int(rec, "appearances")
    debug(f"starters={len(starters)} substitutes={len(subs)} unused={len(unused)} "
          f"subIns={sub_ins} appearances={appearances} "
          f"subIn分={sub_in_min} subOut分={sub_out_min}")

    entry = exit_min = None
    if espn_id in starters:
        status, entry, exit_min = STARTER, 0, sub_out_min
        minutes = sub_out_min if sub_out_min else 90
    elif sub_in_min is not None or sub_ins > 0:
        status, entry = SUBBED_IN, sub_in_min
        minutes = (90 - sub_in_min) if sub_in_min is not None else None
    elif espn_id in unused or espn_id in subs:
        status, minutes = BENCH, None
    elif appearances > 0:
        status, minutes = STARTER, 90
    else:
        status, minutes = BENCH, None

    swap = resolve_name(rec["subbedOutPlayer"], pmap) if rec.get("subbedOutPlayer") else None
    where = ("starters" if espn_id in starters else
             "substitutes" if espn_id in subs else
             "unused" if espn_id in unused else "リスト外")
    return {"status": status, "minutes": minutes, "entry": entry, "exit": exit_min,
            "swap": swap, "team": rec.get("tmName", ""), "league": league,
            "goals": str(stat_int(rec, "totalGoals")),
            "cards": f"{stat_int(rec, 'yellowCards')}/{stat_int(rec, 'redCards')}",
            "via": f"{where}, subIns={sub_ins}"}


def espn_team_latest_result(team_id):
    """チームの結果ページ(埋め込みJSON)から、完了済み試合の最新日付を返す

    選手ページの Last 5 Matches は反映が遅れることがあるため、チーム全体の
    結果ページと突き合わせて「これが本当に最新の試合か」を確認する材料にする。
    戻り値: (日付文字列 "YYYY-MM-DD" または None, 試合URL または None)
    """
    url = f"https://www.espn.com/soccer/team/results/_/id/{team_id}"
    # ESPNはAWS WAFのJS認証チャレンジがあり requests では弾かれるため、
    # 選手ページ・試合ページと同様にブラウザ経由(render)で取得する。
    html = render(url, wait_selector="body", locale="en-US", scroll=True)
    fitt = extract_fitt(html)
    if fitt is None:
        debug("チーム結果ページに __espnfitt__ が無い")
        return None, None
    try:
        events = fitt["page"]["content"]["results"]["events"]
    except (KeyError, TypeError):
        debug("チーム結果ページのJSONに results.events が無い")
        return None, None
    completed = [e for e in events if e.get("completed") and e.get("date")]
    if not completed:
        return None, None
    latest = max(completed, key=lambda e: e["date"])
    m = re.match(r"(\d{4}-\d{2}-\d{2})", latest["date"])
    return (m.group(1) if m else None), urljoin(url, latest.get("link", ""))


def scrape_espn(player):
    espn_id = str(player["espn_id"])
    html = render(player["url"], wait_selector="table", locale="en-US", scroll=True)
    save_html(html, f"{player['name_ja']}_espn_player")

    match = espn_latest_match(html, player["url"])
    if match is None:
        raise ScrapeError("Last 5 Matches から試合リンクを取得できない",
                          "--save-html でHTMLを保存し表構造を確認する", player["url"])

    date_str = match["date"].strftime("%Y-%m-%d") if match["date"] else ""
    log(f"最新試合: {date_str or '日付不明'} ({format_elapsed(date_str) or '経過不明'}) "
        f"{match['result']} gameId={match['game_id']}")

    match_html = render(match["url"], wait_selector="body", locale="en-US", scroll=True)
    save_html(match_html, f"{player['name_ja']}_espn_match")
    res = espn_judge(match_html, espn_id)

    # 出場時間の表示を組み立てる
    if res["status"] == SUBBED_IN and res["entry"]:
        minutes_text = f"{res['entry']}分から出場(約{res['minutes']}分)"
    elif res["status"] == STARTER and res["exit"]:
        minutes_text = f"先発〜{res['exit']}分(約{res['exit']}分)"
    elif res["minutes"]:
        minutes_text = f"約{res['minutes']}分"
    else:
        minutes_text = ""

    notes = []
    if res.get("swap"):
        notes.append(f"{res['swap']}と交代")

    # ESPNは移籍の反映が遅れることがある。想定と違うクラブの記録を
    # 拾っていないか確認できるよう、レコード上の所属を残す
    espn_team = res.get("team") or ""
    if espn_team:
        log(f"ESPN上の所属: {espn_team}")
        expected = player.get("espn_team")
        if expected and norm(expected).lower() not in norm(espn_team).lower():
            notes.append(f"要確認 ESPN所属={espn_team}")
            log(f"想定({expected})と異なる所属です。移籍未反映の可能性", "WARN")
    # 一覧のAPP列と食い違ったら記録に残す
    expected_app = {"sub": SUBBED_IN, "started": STARTER}.get((match["app"] or "").lower())
    if expected_app and expected_app != res["status"]:
        notes.append(f"一覧のAPP列は{match['app']}")
        log(f"APP列 '{match['app']}' と判定 '{res['status']}' が食い違います", "WARN")

    # 選手ページの Last 5 Matches がチームの最新試合を反映しているか確認する
    team_id = player.get("espn_team_id")
    team_date = team_url = ""
    if team_id and date_str:
        try:
            team_date, team_url = espn_team_latest_result(team_id)
        except ScrapeError as e:
            log(f"チーム最新試合の確認に失敗: {e.summary()}", "WARN")
            team_date = team_url = ""
        if team_date and team_date > date_str:
            notes.append("要確認 チームの最新試合と不一致")
            log(f"選手ページ({date_str})より新しいチーム試合({team_date})があります。"
                f"未反映の可能性", "WARN")
        elif team_date:
            debug(f"チーム最新試合({team_date})は選手ページ({date_str})以前。反映済みとみなす")

    row = result(player, res["status"], match=match["opponent"], date=date_str,
                 match_result=match["match_result"], goals=res["goals"],
                 cards=res["cards"], league=res.get("league", ""),
                 url=match["url"], note=" / ".join(notes),
                 team_latest_date=team_date or "", team_latest_url=team_url or "")
    row["出場時間"] = minutes_text
    debug(f"判定経路: {res['via']}")
    return row


def jl_latest_row(player):
    """Jリーグ公式の出場履歴から最新行を取り出す(出場時間の補完用)"""
    df, _ = jl_history_df(player)
    if df.empty:
        return None, None, None
    latest = df.iloc[-1]
    min_col = find_col(df.columns, [JL_COL_MIN])
    app_col = find_col(df.columns, [JL_COL_APP])
    return (latest["_date"],
            _to_minutes(latest[min_col]) if min_col else None,
            str(latest[app_col]).strip() if app_col else "")


def describe_minutes(status, minutes, fl_minute):
    """出場時間の表示を組み立てる

    先発なら「何分まで」、途中出場なら「何分から」が分かるようにする。
    どちらも取れない場合は空にし、備考で後日更新である旨を伝える。
    """
    if status == STARTER:
        if minutes and minutes < 90:
            return f"先発〜{minutes}分で交代({minutes}分)"
        if minutes:
            return f"{minutes}分"
        if fl_minute:
            return f"先発〜{fl_minute}分で交代"
        return ""
    if status == SUBBED_IN:
        if minutes:
            _, label = minutes_to_entry(minutes)
            return f"{label}から出場({minutes}分)" if label else f"{minutes}分"
        if fl_minute:
            return f"{fl_minute}分から出場"
        return ""
    return ""


def _with_team_latest(row, team_latest_date, team_latest_url):
    """チーム最新試合日付/URLを行に付与する(FL/公式どちらの行でも共通)"""
    row["チーム最新試合日付"] = team_latest_date
    row["_team_latest_url"] = team_latest_url
    return row


def scrape_footballlab(player):
    """Football-LAB で出場区分を決め、足りない出場時間をJリーグ公式で補う

    ベンチ外とベンチ入りの区別は Football-LAB のメンバー表でしか付かない。
    一方、交代の分数は後日更新のため、当日はJリーグ公式の出場試合時間を使う。

    戻り値は行のリスト。U-21 Ｊリーグはトップチームとは別日程で開催され、
    Football-LAB のレポートにも載らないため、Jリーグ公式の出場履歴に今季の
    記録があれば2行目として返す。
    """
    fallback = None if CFG.get("no_fallback") else player.get("jleague_url")

    # チームの最新試合(=このシーズンでチーム公式が消化した最新日程)を
    # 先に取得しておく。以降どの経路で行が確定しても同じ値を付与できる。
    try:
        latest = fl_latest_match(player["url"])
    except ScrapeError as e:
        if e.kind == "no_matches":
            # 今季まだ1試合も消化していない。Jリーグ公式に落とすと
            # 昨季以前の試合を最新として拾ってしまうので切り替えない
            log(f"{CFG['year']}年の試合が0件のため 試合なし で確定します")
            return [blank(player, NO_MATCH,
                          f"{CFG['year']}年の試合がまだない(中止・延期など)")]
        if not fallback:
            raise
        log(f"Football-LAB失敗({e.message}) → Jリーグ公式のみで判定します", "WARN")
        jl = scrape_jleague(player, note_prefix="FL取得失敗のため ")
        jl["備考"] += " / ベンチ外は判別不可"
        return with_other_competitions([jl], player, None)
    row = jl_fill_league(_footballlab_row(player, latest), player)
    return with_other_competitions([row], player, latest[3])


def with_other_competitions(rows, player, season_start):
    """トップチームのリーグ戦以外(U-21 Ｊリーグ・ルヴァンカップ・天皇杯)の行を足す"""
    if CFG.get("no_fallback") or not player.get("jleague_url"):
        return rows
    rows = with_u21_row(rows, player, season_start)
    rows = with_levain_row(rows, player, season_start)
    return with_emperor_row(rows, player, season_start)


def with_cup_row(rows, player, season_start, build_row, label, fallback=None):
    """カップ戦(メンバー表を読む大会)の行を足して返す

    メンバー表を読むのでベンチ外も判別できる。取得に失敗した場合、
    fallback(出場履歴からの組み立て)があればそちらへ落とす。
    """
    try:
        cup = build_row(player, season_start)
    except ScrapeError as e:
        log(f"{label}の確認に失敗: {e.summary()}", "WARN")
        cup = None
        if fallback is None:
            return rows
    if cup is None and fallback is not None:
        # メンバー表を読めなくても、出場した試合なら出場履歴に残っている
        try:
            df, _ = jl_history_df(player)
            cup = fallback(player, df, season_start)
        except ScrapeError as e:
            log(f"{label}(出場履歴)の確認にも失敗: {e.summary()}", "WARN")
            return rows
        if cup is not None:
            debug(f"{label}はメンバー表を読めなかったため出場履歴から作成")
    if cup is None:
        return rows
    if any(r["日付"] == cup["日付"] and r["対戦相手"] == cup["対戦相手"] for r in rows):
        debug(f"{label}の試合は1行目と同じなので追加しない")
        return rows
    log(f"{label}の出場記録を追加: {cup['日付']} {cup['ステータス']}")
    rows.append(cup)
    return rows


def with_emperor_row(rows, player, season_start):
    """天皇杯の最新試合があれば行を足して返す

    天皇杯はJFA主催のため、Jリーグ公式の出場履歴にも Football-LAB にも載らない。
    クラブページの試合ログから今季の天皇杯を探し、試合ページのメンバー表で
    出場区分を決める。
    """
    return with_cup_row(rows, player, season_start, jl_emperor_row, EMPEROR)


def with_jl_history_row(rows, player, season_start, build_row, label):
    """Jリーグ公式の出場履歴から作った大会別の行を足して返す

    出場履歴は1回の実行中キャッシュされるので、大会ごとに呼んでも
    選手ページへのアクセスは増えない。
    """
    try:
        df, _ = jl_history_df(player)
    except ScrapeError as e:
        log(f"{label}の確認に失敗: {e.summary()}", "WARN")
        return rows
    extra = build_row(player, df, season_start)
    if extra is None:
        return rows
    # Football-LAB が取れずJリーグ公式に落ちた場合、1行目自体が
    # その大会の試合ということがある。その時は二重に出さない。
    if any(r["日付"] == extra["日付"] and r["対戦相手"] == extra["対戦相手"] for r in rows):
        debug(f"{label}の試合は1行目と同じなので追加しない")
        return rows
    log(f"{label}の出場履歴を追加: {extra['日付']} {extra['ステータス']}")
    rows.append(extra)
    return rows


def with_u21_row(rows, player, season_start):
    """U-21 Ｊリーグの最新試合があれば行を足して返す

    日程ページから今季の試合を探し、試合ページのメンバー表で出場区分を
    決めるのでベンチ外も判別できる。試合ページを読めなかったときだけ
    出場履歴に落とす(その場合はベンチ外を判別できない)。
    """
    return with_cup_row(rows, player, season_start,
                        jl_u21_cup_row, U21, fallback=jl_u21_row)


def with_levain_row(rows, player, season_start):
    """ルヴァンカップの最新試合があれば行を足して返す

    ルヴァンカップは Football-LAB がデータを持っていない。天皇杯と同じく
    クラブページの試合ログ→試合ページのメンバー表を読むので、ベンチ外も
    判別できる。試合ページを読めなかったときだけ出場履歴に落とす
    (その場合はベンチ外を判別できない)。
    """
    return with_cup_row(rows, player, season_start,
                        jl_levain_cup_row, LEVAIN, fallback=jl_levain_row)


def _footballlab_row(player, latest):
    """Football-LAB(+Jリーグ公式)で、トップチームの最新試合の行を1つ作る"""
    fallback = None if CFG.get("no_fallback") else player.get("jleague_url")
    team_latest_date, team_latest_url = latest[2], latest[1]

    try:
        row = _footballlab_core(player, latest)
        member = row.pop("_fl", None)
    except ScrapeError as e:
        if not fallback:
            raise
        log(f"Football-LAB失敗({e.message}) → Jリーグ公式のみで判定します", "WARN")
        jl = scrape_jleague(player, note_prefix="FL取得失敗のため ")
        jl["備考"] += " / ベンチ外は判別不可"
        return _with_team_latest(jl, team_latest_date, team_latest_url)

    # メンバー表が無い場合のみ、Jリーグ公式に丸ごと切り替える
    if member is None:
        if row["ステータス"] == NO_MATCH and fallback:
            log(f"Football-LAB未掲載({row['備考'] or '理由不明'}) "
                f"→ Jリーグ公式に切り替えます", "WARN")
            try:
                jl = scrape_jleague(player, note_prefix="FLメンバー表未掲載のため ")
                jl["備考"] += " / ベンチ外は判別不可"
                return _with_team_latest(jl, team_latest_date, team_latest_url)
            except ScrapeError as e:
                log(f"Jリーグ公式も失敗: {e.summary()}", "WARN")
                row["備考"] = f"{row['備考']} / 公式も失敗".strip(" /")
        return _with_team_latest(row, team_latest_date, team_latest_url)

    status = row["ステータス"]
    # この2区分は個人成績が無いのが正常。公式の出場履歴による補完を
    # 試みないことで、行なし・取得失敗に起因する不要な警告を防ぐ。
    if status in (OUT_SQUAD, BENCH):
        return _with_team_latest(row, team_latest_date, team_latest_url)

    # 同日の公式出場履歴で、個人成績と試合結果を補完する。Football-LABは
    # 出場区分の判定に強く、公式はこの3項目に強いため役割を分ける。
    official_row = official_df = None
    if fallback and row.get("日付"):
        try:
            official_row, official_df, _ = jl_match_record(player, row["日付"])
            if official_row is not None:
                row.update(jl_extra_fields(official_row, official_df.columns))
            else:
                debug("同日のJリーグ公式出場履歴は無し（ベンチ外等では正常）")
        except ScrapeError as e:
            log(f"公式の成績補完に失敗: {e.summary()}", "WARN")
    fl_minute = member.get("minute")
    # 先発かつ交代(←OUT)マークが無ければ最後まで出場したとみなす。
    # マークがあれば fl_minute に交代分数(90分超のロスタイムも含む)が入る。
    fl_full = status == STARTER and not member.get("has_out")
    minutes, notes = None, []

    if fl_full:
        debug("Football-LABに交代(←OUT)マークが無いためフル出場とみなす")
    elif fl_minute:
        debug(f"Football-LABの矢印から分を取得: {fl_minute}分")
    elif official_row is not None:
        jl_minutes = _to_minutes(jl_value(official_row, official_df.columns, JL_COL_MIN))
        jl_app = jl_value(official_row, official_df.columns, JL_COL_APP)
        if jl_minutes:
            minutes = jl_minutes
            notes.append("出場時間はJリーグ公式")
            if jl_app == "先発" and status != STARTER:
                notes.append(f"公式の出場列は{jl_app}")
            elif jl_app == "サブ" and status == STARTER:
                notes.append(f"公式の出場列は{jl_app}")
        else:
            notes.append("出場時間は後日更新")
    elif fallback:
        try:
            jl_date, jl_minutes, jl_app = jl_latest_row(player)
        except ScrapeError as e:
            log(f"出場時間の補完に失敗: {e.summary()}", "WARN")
            jl_date = jl_minutes = jl_app = None
        if jl_minutes:
            minutes = jl_minutes
            if jl_date and jl_date.strftime("%Y-%m-%d") != row["日付"]:
                notes.append(f"出場時間は公式の{jl_date:%-m/%-d}の行")
                log(f"FLの試合日({row['日付']})と公式の最新行({jl_date})が不一致",
                    "WARN")
            else:
                notes.append("出場時間はJリーグ公式")
            if jl_app == "先発" and status != STARTER:
                notes.append(f"公式の出場列は{jl_app}")
            elif jl_app == "サブ" and status == STARTER:
                notes.append(f"公式の出場列は{jl_app}")
        else:
            notes.append("出場時間は後日更新")
    else:
        notes.append("出場時間は後日更新")

    row["出場時間"] = "フル出場" if fl_full else describe_minutes(status, minutes, fl_minute)
    if notes:
        row["備考"] = " / ".join([p for p in [row["備考"]] + notes if p])
    return _with_team_latest(row, team_latest_date, team_latest_url)


# =========================================================
# Jリーグ公式 (Football-LAB のフォールバック)
#
# 選手ページの「出場履歴」テーブル(2026年8月に確認):
#   年月日 | 大会 | 対戦相手 | 勝敗 | 出場 | ゴール数 | シュート数 | 出場試合時間 | 警告/退場
#   26/8/9 | 明治安田Ｊ３ | 滋賀 | 勝 3-0 | サブ | 0 | 2 | 30分 | 0/0
#
# 「出場」列: 先発 / サブ。サブでも出場試合時間があれば途中出場。
# 出場時刻は 90 - 出場試合時間 で逆算できる(公式はアディショナルを数えない)。
# 出場履歴には帯同しなかった試合が載らないため、ベンチ外は判別できない。
# =========================================================
JL_COL_DATE = "年月日"
JL_COL_APP = "出場"
JL_COL_MIN = "出場試合時間"
JL_COL_GOALS = "ゴール数"
JL_COL_RESULT = "勝敗"
JL_COL_CARDS = "警告/退場"
JL_COL_COMP = "大会"
JL_COL_OPP = "対戦相手"
JL_COL_SHOTS = "シュート数"

# 選手ページに埋め込まれたJSON(playerHistory)を、上の表と同じ列名に整形する。
# 表のHTMLには直近数試合しか描かれないが、JSONには全試合分が入っている。
JL_HISTORY_COLUMNS = [JL_COL_DATE, JL_COL_COMP, JL_COL_OPP, JL_COL_RESULT,
                      JL_COL_APP, JL_COL_GOALS, JL_COL_SHOTS, JL_COL_MIN,
                      JL_COL_CARDS]
JL_APPEARANCE = {"start": "先発", "sub": "サブ"}
JL_OUTCOME = {"win": "勝", "draw": "分", "loss": "負"}
# 大会名 "U-21 Ｊリーグ" の判定用(NFKC正規化してから照合する)
RE_U21 = re.compile(r"U-?21", re.I)
# ルヴァンカップの判定用。埋め込みJSONから読めたときは大会ID(gameKindId)が
# 4 で安定しているのでそれを使い、HTMLの表に落ちて大会IDが無い場合に備えて
# 名称でも照合する(ナビスコカップ時代の表記も拾う)。
LEVAIN = "ルヴァンカップ"
LEVAIN_KIND = "4"
RE_LEVAIN = re.compile(r"ルヴァン|ナビスコ")
# 選手ページが持つ「現在の所属リーグ」。埋め込みJSONの入れ子の深さによって
# 引用符のエスケープ(\")が付いたり付かなかったりするため \\* で吸収する。
#   ..."gameKindName":"２０２６／２７明治安田Ｊ１リーグ","gameKindId":2...
RE_JL_GAME_KIND = re.compile(
    r'gameKindName\\*"\s*:\s*\\*"(.*?)\\*"\s*,\s*\\*"gameKindId\\*"\s*:\s*(\d+)')

# 同じ選手ページを何度も取りに行かないよう、1回の実行中は結果を使い回す
# {選手ページURL: {"df": 出場履歴, "league": 所属リーグ名, "slug": クラブslug}}
_JL_HISTORY_CACHE = {}
# {クラブslug: クラブページの試合ログ} — 天皇杯の有無を調べるのに使う
_JL_CLUB_CACHE = {}


def jl_parse_date(text):
    """'26/8/9' を date にする"""
    m = re.match(r"\s*(\d{2,4})/(\d{1,2})/(\d{1,2})", str(text))
    if not m:
        return None
    y, mo, d = (int(x) for x in m.groups())
    if y < 100:
        y += 2000
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def jl_history_records(html):
    """ページに埋め込まれた出場履歴JSON(playerHistory)を取り出す

    Jリーグ公式は Next.js のストリーミングデータとして全試合分を埋め込む一方、
    HTMLの表には直近の数試合しか描画しない。JSONを読めば U-21 Ｊリーグを含む
    全大会の履歴が、大会名付きで取れる。見つからなければ None。
    """
    i = html.find("playerHistory")
    if i < 0:
        debug("ページに playerHistory が無い。HTMLの表から読む")
        return None
    start = html.find("[", i)
    if start < 0:
        return None
    depth, end = 0, -1
    for j in range(start, len(html)):
        if html[j] == "[":
            depth += 1
        elif html[j] == "]":
            depth -= 1
            if depth == 0:
                end = j + 1
                break
    if end < 0:
        return None
    raw = html[start:end]
    # 埋め込みはJSの文字列リテラル内にあるため、多くの場合 \" でエスケープ
    # されている。素のまま読めなければ、エスケープを戻してから読み直す。
    for text in (raw, raw.replace('\\"', '"').replace("\\\\", "\\")):
        try:
            records = json.loads(text)
        except ValueError:
            continue
        if isinstance(records, list) and all(isinstance(r, dict) for r in records):
            return records
    debug("playerHistory を JSON として解釈できなかった")
    return None


def jl_result_text(rec):
    """JSONの result を '勝 3-0' / '負 1-1(PK4-5)' の表記にする"""
    r = rec.get("result") or {}
    mark = JL_OUTCOME.get(r.get("outcome"), "")
    if r.get("playerScore") is None:
        return mark
    text = f"{mark} {r['playerScore']}-{r['opponentScore']}".strip()
    pk = r.get("penalty") or {}
    if pk:
        text += f"(PK{pk.get('playerScore')}-{pk.get('opponentScore')})"
    return text


def jl_records_to_frame(records):
    """playerHistory を、HTMLの表と同じ列名の DataFrame にする"""
    rows = []
    for rec in records:
        minutes = rec.get("minutes")
        rows.append({
            JL_COL_DATE: rec.get("date", ""),
            JL_COL_COMP: rec.get("competition", ""),
            JL_COL_OPP: rec.get("opponent", ""),
            JL_COL_RESULT: jl_result_text(rec),
            JL_COL_APP: JL_APPEARANCE.get(rec.get("appearance"),
                                          rec.get("appearance") or ""),
            JL_COL_GOALS: "" if rec.get("goals") is None else str(rec["goals"]),
            JL_COL_SHOTS: "" if rec.get("shots") is None else str(rec["shots"]),
            JL_COL_MIN: f"{minutes}分" if minutes else "",
            JL_COL_CARDS: rec.get("cards", ""),
            # 大会ID。所属リーグ(gameKindId)と突き合わせて大会名を引くのに使う
            "_kind": str(rec.get("gameKindId", "")),
        })
    return pd.DataFrame(rows, columns=JL_HISTORY_COLUMNS + ["_kind"])


def jl_history_table(html):
    """出場履歴のテーブルを DataFrame で返す

    埋め込みJSONを優先する。HTMLの表は直近の数試合しか描画されないため、
    表だけを読むと少し前の試合や U-21 Ｊリーグの記録を取りこぼす。
    """
    records = jl_history_records(html)
    if records:
        debug(f"出場履歴JSONから{len(records)}行を取得")
        return jl_records_to_frame(records)

    frames = read_tables_any(html)
    for df in frames:
        cols = [str(c) for c in df.columns]
        if any(JL_COL_DATE in c for c in cols) and any(JL_COL_MIN in c for c in cols):
            df.columns = cols
            return df
    if CFG.get("debug"):
        for i, df in enumerate(frames):
            log(f"table[{i}] cols={[str(c) for c in df.columns][:9]}", "DEBUG")
    return None


def jl_league_from_page(html, df):
    """選手ページから、その選手が今所属しているリーグの大会名を返す

    ページには gameKindName("２０２６／２７明治安田Ｊ１リーグ")と、出場履歴と
    共通の大会ID(gameKindId)が入っている。同じIDの出場履歴があれば、そちらの
    表記("明治安田Ｊ１")に揃えて他の行と見た目を合わせる。
    """
    m = RE_JL_GAME_KIND.search(html)
    if not m:
        debug("選手ページに gameKindName が無い")
        return ""
    name, kind = m.group(1).strip(), m.group(2)
    comp_col = find_col(df.columns, [JL_COL_COMP]) if df is not None else None
    if comp_col is not None and "_kind" in df.columns:
        same = df[df["_kind"] == kind]
        if not same.empty:
            return str(same.iloc[-1][comp_col]).strip()
    return name


def jl_page_value(player, key):
    """選手ページから読んだ値(league / slug)。未取得なら選手ページを取りに行く"""
    url = player.get("jleague_url") or player.get("url")
    if url not in _JL_HISTORY_CACHE:
        jl_history_df(player)
    return (_JL_HISTORY_CACHE.get(url) or {}).get(key, "")


def jl_current_league(player):
    """選手の現在の所属リーグ名"""
    return jl_page_value(player, "league")


def jl_fill_league(row, player):
    """リーグ名が空の行に、Jリーグ公式の所属リーグ名を補う

    ベンチ外・ベンチ入りのみの試合は公式の出場履歴に載らないため、その試合の
    大会名を直接は特定できない。代わりに所属リーグを当てる。カップ戦だった
    可能性は残るので、推定であることを備考に残す。
    """
    if row.get("リーグ名") or row["ステータス"] == NO_MATCH:
        return row
    if CFG.get("no_fallback") or not player.get("jleague_url"):
        return row
    try:
        league = jl_current_league(player)
    except ScrapeError as e:
        log(f"所属リーグの取得に失敗: {e.summary()}", "WARN")
        return row
    if not league:
        return row
    debug(f"リーグ名を所属リーグで補完: {league}")
    row["リーグ名"] = league
    row["備考"] = " / ".join([p for p in [row["備考"], "リーグ名は所属リーグから推定"] if p])
    return row


def jl_history_df(player):
    """Jリーグ公式の出場履歴を (DataFrame, URL) で返す。日付昇順・_date列付き

    出場時間の補完・同日行の照合・U-21 の抽出で同じページを何度も取りに
    行かないよう、1回の実行中は取得結果を使い回す。
    """
    url = player.get("jleague_url") or player.get("url")
    cached = _JL_HISTORY_CACHE.get(url)
    if cached is not None:
        return cached["df"].copy(), url

    html = fetch(url)
    save_html(html, f"{player['name_ja']}_jleague")
    df = jl_history_table(html)
    if df is None:
        # サーバ側で描画されていない場合はブラウザで取り直す
        debug("Jリーグ公式は表を画面側で描画するためブラウザで取得します")
        html = render(url, wait_selector="table", locale="ja-JP", scroll=True)
        save_html(html, f"{player['name_ja']}_jleague_rendered")
        df = jl_history_table(html)
    if df is None:
        raise ScrapeError("出場履歴テーブルが見つからない",
                          "--save-html でHTMLを保存し列名を確認する", describe(html, url))

    date_col = find_col(df.columns, [JL_COL_DATE])
    df["_date"] = df[date_col].map(jl_parse_date)
    df = df.dropna(subset=["_date"]).sort_values("_date")
    _JL_HISTORY_CACHE[url] = {"df": df,
                              "league": jl_league_from_page(html, df),
                              "slug": jl_club_slug(html)}
    return df.copy(), url


def jl_match_record(player, target_date=None):
    """Jリーグ公式の出場履歴から、指定日（なければ最新）の1行を返す。"""
    df, url = jl_history_df(player)
    if df.empty:
        return None, df, url
    if target_date:
        target = datetime.strptime(target_date, "%Y-%m-%d").date()
        exact = df[df["_date"] == target]
        if exact.empty:
            return None, df, url
        return exact.iloc[-1], df, url
    return df.iloc[-1], df, url


def jl_value(row, columns, label):
    col = find_col(columns, [label])
    if col is None:
        return ""
    value = str(row[col]).strip()
    return "" if value.lower() in ("nan", "none") else value


def jl_extra_fields(row, columns):
    """公式の1行をExcel用のリーグ名・ゴール・結果・警告/退場列へ変換する。"""
    return {"リーグ名": jl_value(row, columns, JL_COL_COMP),
            "試合結果": jl_value(row, columns, JL_COL_RESULT),
            "ゴール数": jl_value(row, columns, JL_COL_GOALS),
            "警告/退場": jl_value(row, columns, JL_COL_CARDS)}


def jl_status_from_row(row, columns):
    """公式の出場履歴1行から、出場区分と出場時間の表示を組み立てる

    出場履歴には帯同しなかった試合が載らないため、ベンチ外は判別できない。
    """
    minutes = _to_minutes(jl_value(row, columns, JL_COL_MIN))
    app = jl_value(row, columns, JL_COL_APP)
    if "先発" in app:
        if minutes and minutes < 90:
            return STARTER, f"先発〜{minutes}分で交代"
        return STARTER, (f"{minutes}分" if minutes else "")
    if minutes:
        _, label = minutes_to_entry(minutes)
        return SUBBED_IN, (f"{label}から出場({minutes}分)" if label else f"{minutes}分")
    return BENCH, ""


def jl_latest_comp_row(player, df, season_start, is_target, label):
    """出場履歴から、ある大会の今季最新試合を1行にする。無ければ None

    トップチームの行は Football-LAB のレポート(=リーグ戦)から作るため、
    別日程で行われる大会はJリーグ公式の出場履歴からしか拾えない。1行に
    まとめるとどちらかの試合が見えなくなるので、大会ごとに別の行として出す。

    is_target(大会名, 大会ID) が True を返す行だけを対象にする。大会名は
    NFKC正規化済み、大会IDは埋め込みJSONから読めた場合だけ入る。
    """
    if df.empty:
        return None
    comp_col = find_col(df.columns, [JL_COL_COMP])
    if comp_col is None:
        debug(f"出場履歴に大会列が無いため{label}の判定をスキップ")
        return None
    kinds = df["_kind"] if "_kind" in df.columns else pd.Series("", index=df.index)
    target = df[[bool(is_target(norm(comp), str(kind)))
                 for comp, kind in zip(df[comp_col], kinds)]]
    # 出場履歴は過去シーズン分も含むため、今季の記録だけに絞る
    target = target[target["_date"] >= (season_start or date(CFG["year"], 1, 1))]
    if target.empty:
        return None

    latest = target.iloc[-1]
    status, minutes_text = jl_status_from_row(latest, df.columns)
    row = result(player, status, match=jl_value(latest, df.columns, JL_COL_OPP),
                 date=latest["_date"].strftime("%Y-%m-%d"),
                 url=player.get("jleague_url") or player.get("url"),
                 note=f"Jリーグ公式({label}) / ベンチ外は判別不可")
    row.update(jl_extra_fields(latest, df.columns))
    row["出場時間"] = minutes_text
    return row


def jl_u21_row(player, df, season_start):
    """出場履歴から U-21 Ｊリーグの今季最新試合を1行にする。無ければ None"""
    return jl_latest_comp_row(player, df, season_start,
                              lambda comp, kind: RE_U21.search(comp),
                              "U-21 Ｊリーグ")


def jl_levain_row(player, df, season_start):
    """出場履歴から ルヴァンカップ の今季最新試合を1行にする。無ければ None

    ルヴァンカップは Football-LAB がデータを持っていないため、Jリーグ公式の
    出場履歴からしか拾えない。
    """
    return jl_latest_comp_row(
        player, df, season_start,
        lambda comp, kind: kind == LEVAIN_KIND or RE_LEVAIN.search(comp),
        LEVAIN)


# =========================================================
# 天皇杯 (Jリーグ公式のクラブページ + 試合ページ)
#
# 天皇杯はJFA主催のため、選手ページの出場履歴には載らない(2026年8月に確認)。
# 代わりにクラブページへ全大会の試合ログが埋まっているので、そこから試合IDを
# 取り出し、試合ページのメンバー表を読んで出場区分を決める。
#
#   クラブページ https://www.jleague.jp/club/<slug>/
#     gameLogInClubByYear[<年>-<大会群>] = [{gameId, gameDate, leagueName, ...}]
#   試合ページ   https://www.jleague.jp/match/emperor/<年>/<gameIdの下6桁>/
#     homeTeam/awayTeam.players   先発11人(playerId付き)
#     allMemberIds                両チームのメンバー入り全員 → ベンチ外の判別
#     m-lineup-list-item の subs  交代の分数("68’")
#     eventType goal/yellow-card  得点・警告(playerId付き)
#
# メンバー表はJリーグ公式が画面側で描画するため、DOMではなく埋め込みJSONを読む。
# =========================================================
EMPEROR = "天皇杯"
JL_CLUB_URL = "https://www.jleague.jp/club/{slug}/"
RE_JL_CLUB_SLUG = re.compile(r'clubId\\*"\s*:\s*\\*"([A-Za-z0-9_-]+)\\*"')
RE_JL_PLAYER_ID = re.compile(r"/player/(\d+)")
RE_JL_ALL_MEMBERS = re.compile(r'"allMemberIds"\s*:\s*(\[[^\]]*\])')
# 交代イベントの時刻。"46*’" の * はハーフタイム等の区切りでの交代を表す
RE_JL_EVENT_TIME = re.compile(r"(\d+)\s*(\*?)")
# "m-lineup-list-item" と "subs" の間には "ref" のほかに属性が入ることがある。
#   通常          ..."m-lineup-list-item","ref":"$undefined","subs":[...]
#   警告を受けた選手 ..."m-lineup-list-item","ref":"$undefined","card":"yellow","subs":[...]
# 属性を決め打ちすると、警告を受けた選手の交代情報だけを取りこぼす。
RE_JL_LINEUP_SUBS = re.compile(
    r'"m-lineup-list-item"(?:,"[A-Za-z]+":"[^"]*")*?,"subs":(\[[^\]]*\])'
    r'.{0,400}?"memberId":"(\d+)"', re.S)
RE_JL_MATCH_EVENT = re.compile(
    r'"eventType":"(goal|yellow-card|red-card)",'
    r'"legacyPlayerPhotoLookup":\{[^}]*?"playerId":"(\d+)"')


def jl_unescape(text):
    """RSCストリーム(JSの文字列リテラル)内のエスケープを戻す"""
    return text.replace('\\"', '"').replace("\\\\", "\\")


def jl_json_after(text, key, open_ch="{"):
    """key の直後にある JSON を括弧の対応で切り出して読む。読めなければ None"""
    i = text.find(key)
    if i < 0:
        return None
    start = text.find(open_ch, i + len(key))
    if start < 0:
        return None
    close_ch = "}" if open_ch == "{" else "]"
    depth = 0
    for j in range(start, len(text)):
        if text[j] == open_ch:
            depth += 1
        elif text[j] == close_ch:
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(text[start:j + 1])
                except ValueError:
                    return None
    return None


def jl_club_slug(html):
    """選手ページからクラブのslug(例: gifu)を取り出す"""
    m = RE_JL_CLUB_SLUG.search(html)
    return m.group(1) if m else ""


def jl_score_text(score):
    """クラブページの scoreResult を '勝 3-1' / '負 2-2(PK4-5)' にする"""
    if not isinstance(score, dict):
        return ""
    text = f"{score.get('label', '')} {score.get('score')}-{score.get('lost')}".strip()
    pk = score.get("penalty") or {}
    if pk:
        text += f"(PK{pk.get('score')}-{pk.get('lost')})"
    return text


def jl_club_games(slug, name_ja=""):
    """クラブページの試合ログ(全大会・全シーズン)を返す"""
    if slug in _JL_CLUB_CACHE:
        return _JL_CLUB_CACHE[slug]
    url = JL_CLUB_URL.format(slug=slug)
    html = fetch(url)
    save_html(html, f"{name_ja or slug}_club")
    by_year = jl_json_after(jl_unescape(html), '"gameLogInClubByYear"') or {}
    games = []
    for entries in by_year.values():
        for g in entries:
            games.append({"date": jl_parse_date(g.get("gameDate")),
                          "league": str(g.get("leagueName") or ""),
                          "opponent": str(g.get("opponentTeamShortName") or ""),
                          "result": jl_score_text(g.get("scoreResult")),
                          "game_id": str(g.get("gameId") or "")})
    if not games:
        debug(f"クラブページから試合ログを取得できない: {url}")
    _JL_CLUB_CACHE[slug] = games
    return games


def jl_match_url(game_id, path):
    """gameId 2026081917 → https://www.jleague.jp/match/<path>/2026/081917/

    path は大会ごとのURL断片(天皇杯 emperor / ルヴァンカップ leaguecup)。
    """
    return f"https://www.jleague.jp/match/{path}/{game_id[:4]}/{game_id[4:]}/"


def jl_match_members(html):
    """試合ページから (先発ID, メンバーID全員, 交代情報, 得点・警告) を取り出す"""
    plain = jl_unescape(html)

    starters = set()
    for side in ('"homeTeam":{"teamName"', '"awayTeam":{"teamName"'):
        i = plain.find(side)
        if i < 0:
            continue
        for p in jl_json_after(plain[i:], '"players"', "[") or []:
            pid = (p.get("legacyPlayerPhotoLookup") or {}).get("playerId")
            if pid:
                starters.add(str(pid))

    members = set()
    m = RE_JL_ALL_MEMBERS.search(plain)
    if m:
        try:
            members = {str(x) for x in json.loads(m.group(1))}
        except ValueError:
            debug("allMemberIds を JSON として読めなかった")

    subs = {}
    for mm in RE_JL_LINEUP_SUBS.finditer(plain):
        try:
            subs[mm.group(2)] = json.loads(mm.group(1))
        except ValueError:
            continue

    events = {}
    for kind, pid in RE_JL_MATCH_EVENT.findall(plain):
        events.setdefault(pid, {"goal": 0, "yellow-card": 0, "red-card": 0})[kind] += 1

    debug(f"試合ページ: 先発{len(starters)}人 メンバー{len(members)}人 "
          f"交代{len(subs)}人分 得点・警告{len(events)}人分")
    return starters, members, subs, events


def jl_event_minute(time_text):
    """交代イベントの時刻表記を分に直す

    "70’" は70分。"46*’" の * はハーフタイムなど区切りでの交代を表し、
    Jリーグ公式はこれを45分として数える(46*→45、延長の 106*→105 も同じ)。
    """
    m = RE_JL_EVENT_TIME.search(str(time_text))
    if not m:
        return None
    minute = int(m.group(1))
    return minute - 1 if m.group(2) else minute


def jl_entry_label(entry):
    """出場した時刻を "後半25分" のような表記にする"""
    if entry is None:
        return ""
    if entry <= 0:
        return "開始"
    if entry < HALF_MINUTES:
        return f"前半{entry}分"
    if entry == HALF_MINUTES:
        return "後半開始"
    return f"後半{entry - HALF_MINUTES}分"


def jl_member_status(player_id, starters, members, subs):
    """メンバー表から (出場区分, 出場時間の表示, 出場分数) を決める

    メンバー表(allMemberIds)に名前が無ければベンチ外と判定できる。
    出場履歴しか見ない大会との精度の差はここから来る。

    分数はJリーグ公式の出場時間と一致する数え方にしている。
      先発     … 交代で退いた分。退いていなければ90(フル出場)
      途中出場 … (退いた分 または 90) − 入った分
    途中出場の選手がさらに交代で退くことがあるため、in だけでなく out も見る。
    見ていないと、その選手だけ試合終了まで出場した扱いになる。
    """
    events = subs.get(player_id) or []

    def minute(kind):
        for e in events:
            if e.get("type") == kind:
                return jl_event_minute(e.get("time", ""))
        return None

    if player_id in starters:
        out = minute("out")
        if not out:
            return STARTER, "フル出場", FULL_MATCH_MINUTES
        return STARTER, f"先発〜{out}分で交代({out}分)", out
    if player_id in members:
        entry = minute("in")
        if entry is None:
            return BENCH, "", 0
        out = minute("out")
        minutes = max((out or FULL_MATCH_MINUTES) - entry, 0)
        label = jl_entry_label(entry)
        if not label:
            return SUBBED_IN, f"{minutes}分", minutes
        if out:
            return SUBBED_IN, f"{label}〜{out}分で交代({minutes}分)", minutes
        return SUBBED_IN, f"{label}から出場({minutes}分)", minutes
    return OUT_SQUAD, "", 0


def jl_cup_row(player, season_start, league, path):
    """今季のカップ戦の最新試合から、その選手の出場状況を1行にする。無ければ None

    クラブページの試合ログで対象大会の試合を見つけ、その試合ページの
    メンバー表を読む。メンバー表を見るのでベンチ外も判別できる。

    league … クラブページの試合ログに出る大会名(天皇杯 / ルヴァンカップ)
    path   … 試合ページのURL断片(emperor / leaguecup)
    """
    m = RE_JL_PLAYER_ID.search(player.get("jleague_url") or "")
    slug = jl_page_value(player, "slug")
    if not m or not slug:
        debug(f"選手IDまたはクラブslugが取れないため{league}をスキップ")
        return None
    player_id = m.group(1)

    # クラブページの試合ログは過去シーズン分も含むので、今季開幕以降に絞る
    floor = season_start or date(CFG["year"], 1, 1)
    cup = [g for g in jl_club_games(slug, player["name_ja"])
           if league in g["league"] and g["date"] and g["date"] >= floor and g["game_id"]]
    if not cup:
        debug(f"{CFG['year']}シーズンの{league}の試合は無し")
        return None

    latest = max(cup, key=lambda g: g["date"])
    match_url = jl_match_url(latest["game_id"], path)
    html = fetch(match_url)
    save_html(html, f"{player['name_ja']}_{path}")
    starters, members, subs, events = jl_match_members(html)
    if not members:
        raise ScrapeError(f"{league}の試合ページからメンバー表を取得できない",
                          "--save-html でHTMLを保存し構造を確認する", match_url)

    status, minutes_text, _minutes = jl_member_status(
        player_id, starters, members, subs)
    played = status in (STARTER, SUBBED_IN)
    ev = events.get(player_id) or {}
    row = result(player, status, match=latest["opponent"],
                 date=latest["date"].strftime("%Y-%m-%d"),
                 url=match_url, league=league, match_result=latest["result"],
                 goals=str(ev.get("goal", 0)) if played else "",
                 cards=f"{ev.get('yellow-card', 0)}/{ev.get('red-card', 0)}" if played else "",
                 note=f"Jリーグ公式({league})")
    row["出場時間"] = minutes_text
    return row


def jl_emperor_row(player, season_start):
    """今季の天皇杯の最新試合から、その選手の出場状況を1行にする。無ければ None"""
    return jl_cup_row(player, season_start, EMPEROR, "emperor")


def jl_levain_cup_row(player, season_start):
    """今季のルヴァンカップの最新試合から、その選手の出場状況を1行にする"""
    return jl_cup_row(player, season_start, LEVAIN, "leaguecup")


# =========================================================
# U-21 Ｊリーグ (日程検索ページ + 試合ページ)
#
# U-21 Ｊリーグはクラブページの試合ログに載らないため、天皇杯のように
# 「出場していない試合」を見つける入口が長らく無かった。日程検索ページなら
# 期間を指定して今季の全試合を取れるので、そこからクラブの試合を探す。
#   https://www.jleague.jp/u-21/match/search-list/
#       ?startdate=2026-03-01&enddate=2026-09-03&period=custom
# &club=<クラブID> を付けると、そのクラブの試合だけに絞れる。
# U-21 Ｊリーグに参加しているクラブは一部だけなので、players.json に
# u21_club_id を書いた選手だけを対象にする(未記入なら一覧を取りに行かない)。
#
# なお絞り込まずに取ると、ReactのSuspenseで一部のカードが遅延配信され、
# requests では中身が <template id="P:n"> のままの試合が混じる
# (2026年9月時点で5件中2件)。数試合だけ静かに欠けるとベンチ外の判定を誤るため、
# リンク数とパースできた数が食い違ったときはブラウザ描画で取り直す。
# 1試合分のDOM:
#   <div class="m-schedule__wrapper">
#     <a class="m-schedule__link" href="/match/u-21/2026/082309/">
#       <div class="m-schedule__team-home">
#         <span class="m-schedule__team-name" data-media="pc">U-21 ファジアーノ岡山</span>
#         <span class="m-schedule__team-name" data-media="sp">U-21岡山</span>
#       <p class="m-schedule__score">0</p> 試合終了 <p class="m-schedule__score">1</p>
#       <div class="m-schedule__team-away"> ... U-21 ガンバ大阪
# 試合IDは href の 6桁(月日+連番)。日付もここから決まる。
# =========================================================
U21 = "U-21 Ｊリーグ"
# 日程の取得元・開幕月は u21_clubs.json で変えられる(順位決定戦が別大会・
# 別URLになった場合に、設定を足すだけで同じ台帳へ取り込めるようにするため)。
_U21_CONFIG = {}
# {(開始日, 終了日, クラブID): 試合リスト} — 同じクラブの選手で使い回す
_JL_U21_CACHE = {}


def u21_config():
    """u21_clubs.json の内容を返す(1回の実行中は読み直さない)"""
    if not _U21_CONFIG:
        _U21_CONFIG.update(u21_clubs.load_config(CFG.get("u21_clubs_config")))
    return _U21_CONFIG


def u21_season_start():
    """今季の日程検索の開始日。開幕月の1日"""
    return date(CFG["year"], u21_config()["season_start_month"], 1)


def u21_match_href_re(path):
    """試合ページURL(/match/<path>/<年>/<月日連番>/)を拾う正規表現"""
    return re.compile(rf"/match/{re.escape(path)}/(\d{{4}})/(\d{{6}})/")


def jl_u21_games(club_id):
    """U-21 Ｊリーグの、そのクラブの今季の消化済み試合を返す(日付の古い順)

    検索期間は開幕月の1日から当日まで。Football-LAB から得る season_start は
    レポートページに出ている範囲までしか遡らないため、それを開始日にすると
    開幕直後の試合を取りこぼすことがある。
    終了日を当日にしているので、シーズンが年をまたいでも(2026-27は2027年3月
    まで)そのまま拾える。
    """
    opening = u21_season_start()
    today = date.today()
    if today < opening:
        debug(f"{U21}は{opening.year}年{opening.month}月開幕のためまだ試合が無い")
        return []
    start, end = opening.strftime("%Y-%m-%d"), today.strftime("%Y-%m-%d")
    key = (start, end, club_id)
    if key in _JL_U21_CACHE:
        return _JL_U21_CACHE[key]

    # 同じ試合が複数の取得元に出ても、試合IDで1件にまとまる
    found = {}
    for schedule in u21_config()["schedules"]:
        for game in jl_u21_schedule_games(schedule, club_id, start, end):
            found.setdefault(game["game_id"], game)
    games = sorted(found.values(), key=lambda g: (g["date"], g["game_id"]))
    debug(f"{U21}の日程({start}〜{end} club={club_id}): 消化済み{len(games)}試合")
    _JL_U21_CACHE[key] = games
    return games


def jl_u21_schedule_games(schedule, club_id, start, end):
    """取得元1件ぶんの日程を読む。失敗しても他の取得元を止めない"""
    url = schedule["url"].format(start=start, end=end, club=club_id)
    href_re = u21_match_href_re(schedule["path"])
    try:
        html = fetch(url)
    except ScrapeError as e:
        log(f"{schedule['name']}の日程を取得できません({e.message})", "WARN")
        return []
    games = jl_u21_parse_games(html, url, href_re)
    # 遅延配信で中身が届いていない試合があれば、ブラウザで取り直す
    linked = {m[1] for m in href_re.findall(html)}
    if len(games) < len(linked):
        debug(f"{schedule['name']}の日程: リンク{len(linked)}件に対し"
              f"パースできたのは{len(games)}件。ブラウザで取り直します")
        try:
            html = render(url, wait_selector="div.m-schedule__wrapper",
                          locale="ja-JP", scroll=True)
            games = jl_u21_parse_games(html, url, href_re)
        except ScrapeError as e:
            log(f"{schedule['name']}の日程をブラウザで取得できません({e.message})。"
                f"一部の試合が欠けたまま判定します", "WARN")
    save_html(html, f"u21_search_{schedule['path']}_{club_id}")
    return games


def jl_u21_parse_games(html, base_url, href_re):
    """日程検索ページから、結果の出ている試合だけを取り出す"""
    soup = BeautifulSoup(html, "lxml")
    games = []
    for card in soup.select("div.m-schedule__wrapper"):
        link = next((a for a in card.select("a[href]")
                     if href_re.search(a["href"])), None)
        if link is None:
            continue
        m = href_re.search(link["href"])
        year, tail = m.group(1), m.group(2)
        try:
            when = date(int(year), int(tail[:2]), int(tail[2:4]))
        except ValueError:
            debug(f"U-21の試合IDから日付にできない: {year}/{tail}")
            continue
        scores = [s.get_text(strip=True) for s in card.select("p.m-schedule__score")]
        if len(scores) < 2 or not all(s.isdigit() for s in scores[:2]):
            continue  # 未消化(キックオフ時刻だけの表示)
        names = {}
        for side in ("home", "away"):
            box = card.select_one(f"div.m-schedule__team-{side}")
            if box is None:
                continue
            names[side] = {
                media: (box.select_one(f"span.m-schedule__team-name[data-media={media}]")
                        or box).get_text(strip=True)
                for media in ("pc", "sp")}
        if len(names) < 2:
            continue
        games.append({
            "date": when,
            "game_id": f"{year}{tail}",
            "url": urljoin(base_url, link["href"]),
            "home": names["home"]["pc"], "home_short": names["home"]["sp"],
            "away": names["away"]["pc"], "away_short": names["away"]["sp"],
            "home_goals": int(scores[0]), "away_goals": int(scores[1]),
        })
    return games


def jl_u21_cup_row(player, season_start):
    """U-21 Ｊリーグの最新試合をメンバー表から1行にする。無ければ None

    出場履歴と違い、日程から試合を探して試合ページのメンバー表を読むので、
    出場しなかった試合もベンチ入りのみ/ベンチ外として判別できる。
    """
    club_id = str(player.get("u21_club_id", "")).strip()
    if not club_id:
        # 所属クラブが U-21 Ｊリーグに参加していない(=大半の選手)。
        # 一覧を取りに行かず、出場履歴からの組み立てに任せる。
        debug(f"u21_club_id が未設定のため{U21}の日程は見ない")
        return None
    m = RE_JL_PLAYER_ID.search(player.get("jleague_url") or "")
    club = norm(player.get("club", ""))
    if not m or not club:
        debug(f"選手IDまたはクラブ名が取れないため{U21}をスキップ")
        return None
    player_id = m.group(1)

    mine = [g for g in jl_u21_games(club_id)
            if club in norm(g["home"]) or club in norm(g["away"])]
    if not mine:
        debug(f"{CFG['year']}シーズンの{U21}に {player['club']} の試合は無し")
        return None

    latest = max(mine, key=lambda g: g["date"])
    is_home = club in norm(latest["home"])
    html = fetch(latest["url"])
    save_html(html, f"{player['name_ja']}_u21")
    starters, members, subs, events = jl_match_members(html)
    if not members:
        raise ScrapeError(f"{U21}の試合ページからメンバー表を取得できない",
                          "--save-html でHTMLを保存し構造を確認する", latest["url"])

    status, minutes_text, _minutes = jl_member_status(
        player_id, starters, members, subs)
    played = status in (STARTER, SUBBED_IN)
    ev = events.get(player_id) or {}
    scored, conceded = ((latest["home_goals"], latest["away_goals"]) if is_home
                        else (latest["away_goals"], latest["home_goals"]))
    outcome = "勝" if scored > conceded else "負" if scored < conceded else "分"
    row = result(player, status,
                 match=latest["away_short"] if is_home else latest["home_short"],
                 date=latest["date"].strftime("%Y-%m-%d"),
                 url=latest["url"], league=U21,
                 match_result=f"{outcome} {scored}-{conceded}",
                 goals=str(ev.get("goal", 0)) if played else "",
                 cards=f"{ev.get('yellow-card', 0)}/{ev.get('red-card', 0)}" if played else "",
                 note=f"Jリーグ公式({U21})")
    row["出場時間"] = minutes_text
    return row


# =========================================================
# U-21 Ｊリーグ チーム単位 (台帳方式)
#
# レンタル選手の追跡(選手 → 大会)とは向きが逆で、クラブ → 試合 →
# メンバー全員 の順にたどる。players.json への選手登録は不要。
#
#   ① 日程検索(club指定)から今季の消化済み試合の一覧を得る
#   ② 台帳(SQLite)に未記録の試合だけ、試合ページを取りに行く
#   ③ メンバー表の全員を「1試合1選手1レコード」で台帳へ記録する
#   ④ 台帳から、選手ごとの最新1試合＋今季の通算を組み立ててシートにする
#
# ②の差分取得により、実行間隔に左右されない。2週間空けても、その間の全試合が
# まとめて反映される。順位決定戦などが後から日程に足されても、台帳に無い
# 試合IDとして現れるだけなので自動的に取り込まれる。
# 台帳はJリーグ公式から再生成できる(全試合を取り直せば同じ内容になる)ため、
# 壊れても失うのは取得の手間だけ。
#
# 日程は先に全部公開され、結果が付いた試合だけが「消化済み」になる。試合当日は
# 内容が後から補正されることがあるので、新しい方から数試合は毎回取り直す。
#
# 試合ページはRSCのストリーミング参照で書かれているが、参照先の実体も
# 同じHTML内にあるため requests だけで全員分取れる。控え1人分の実体:
#   1e0:["$","$Ld7","1657670",{"href":"/player/1657670", ...
#        {"memberId":"1657670", ... "teamNameKey":"nagasaki" ...
#        ..."m-lineup-list-item__position", ... "children":"GK 13"}
#        ..."m-lineup-list-item__name", ... "children":"黒川 偉心"}
#
# ベンチ外はメンバー表に載らないため、この方式では出せない(仕様)。
# =========================================================
MAIN_SHEET = "レンタル選手"
U21_SHEET = "U-21最新試合"
U21_COLUMNS = ["背番号", "ポジション", "選手", "クラブ", "リーグ名",
               "最新試合関与試合", "経過", "ステータス", "出場時間",
               "対戦相手", "試合結果", "ゴール数", "警告/退場",
               "メンバー入り", "先発", "途中出場", "ベンチ入りのみ",
               "通算出場時間(分)", "備考", "URL"]
# 出力の並び順。名鑑と同じ GK→DF→MF→FW、その中は背番号順
U21_POS_ORDER = {"GK": 0, "DF": 1, "MF": 2, "FW": 3}
# 新しい方からこの数だけは毎回取り直す(試合当日の内容が後から補正されるため)
U21_REFETCH_RECENT = 2

# メンバー1人分のかたまりは '"memberId":"<選手ID>"' で始まる。そこで区切って
# から中を読むので、ある選手の項目を隣の選手の項目と取り違える余地がない。
RE_U21_ITEM_ID = re.compile(r'^(\d+)"')
RE_U21_TEAM_KEY = re.compile(r'"teamNameKey":"([A-Za-z0-9_-]+)"')
RE_U21_ITEM_POS = re.compile(
    r'm-lineup-list-item__position"[^}]{0,200}?"children":"([^"]*)"')
RE_U21_ITEM_NAME = re.compile(
    r'm-lineup-list-item__name"[^}]{0,200}?"children":"([^"]*)"')
# "GK 13" → ポジション と 背番号
RE_U21_POS_NO = re.compile(r"^([A-Za-z]+)\s*(\d*)$")


def jl_match_team_keys(plain):
    """試合ページから {'home': 'nagasaki', 'away': 'nagoya'} を作る

    メンバー表の項目にはクラブ名が入っていないが teamNameKey は入っている。
    先発の1人目から拾えば、どちらのクラブの項目かを確実に見分けられる。
    """
    keys = {}
    for side, marker in (("home", '"homeTeam":{"teamName"'),
                         ("away", '"awayTeam":{"teamName"')):
        i = plain.find(marker)
        if i < 0:
            continue
        for p in jl_json_after(plain[i:], '"players"', "[") or []:
            key = (p.get("legacyPlayerPhotoLookup") or {}).get("teamNameKey")
            if key:
                keys[side] = str(key)
                break
    return keys


def u21_lineup_items(plain):
    """メンバー表を [{id, team_key, pos, no, name}] にする(表示順のまま)"""
    items, seen = [], set()
    for part in plain.split('"memberId":"')[1:]:
        m = RE_U21_ITEM_ID.match(part)
        if not m or m.group(1) in seen:
            continue
        team = RE_U21_TEAM_KEY.search(part)
        pos = RE_U21_ITEM_POS.search(part)
        name = RE_U21_ITEM_NAME.search(part)
        if not (team and pos and name):
            # allMemberIds には監督など、メンバー表に描画されない項目も混ざる
            continue
        seen.add(m.group(1))
        pn = RE_U21_POS_NO.match(pos.group(1).strip())
        items.append({
            "id": m.group(1), "team_key": team.group(1),
            "pos": pn.group(1) if pn else pos.group(1).strip(),
            "no": pn.group(2) if pn else "",
            "name": name.group(1),
        })
    return items


def u21_team_side(game, club_name):
    """対象クラブがホームかアウェーかを返す。判別できなければ None"""
    want = norm(club_name)
    if not want:
        return None
    for side in ("home", "away"):
        if norm(game[side]) == want:
            return side
    for side in ("home", "away"):
        if want in norm(game[side]) or norm(game[side]) in want:
            return side
    return None


def u21_result_text(game, side):
    """対象クラブ視点の '勝 3-1' を作る"""
    scored, conceded = ((game["home_goals"], game["away_goals"]) if side == "home"
                        else (game["away_goals"], game["home_goals"]))
    outcome = "勝" if scored > conceded else "負" if scored < conceded else "分"
    return f"{outcome} {scored}-{conceded}"


# ---------------------------------------------------------
# 台帳 (SQLite)
#
# 置き場所は exe と同じフォルダ。Dropboxフォルダで動かすため WAL は使わない
# (-wal / -shm が別ファイルとして同期され、DBが壊れることがある)。既定の
# journal_mode のまま、処理の間だけ開いて閉じる。
# ---------------------------------------------------------
U21_DB_NAME = "u21_history.db"
U21_DB_VERSION = 1
U21_DB_SCHEMA = """
CREATE TABLE IF NOT EXISTS game(
  season INTEGER NOT NULL, club_id TEXT NOT NULL, game_id TEXT NOT NULL,
  date TEXT NOT NULL,
  side TEXT NOT NULL,
  opponent TEXT NOT NULL, result TEXT NOT NULL,
  url TEXT NOT NULL, fetched_at TEXT NOT NULL,
  PRIMARY KEY(season, club_id, game_id));

CREATE TABLE IF NOT EXISTS appearance(
  season INTEGER NOT NULL, club_id TEXT NOT NULL, game_id TEXT NOT NULL,
  player_id TEXT NOT NULL,
  name TEXT NOT NULL, no TEXT, pos TEXT,
  status TEXT NOT NULL,
  minutes_text TEXT, minutes INTEGER NOT NULL DEFAULT 0,
  goals INTEGER NOT NULL DEFAULT 0,
  yellow INTEGER NOT NULL DEFAULT 0, red INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY(season, club_id, game_id, player_id));

CREATE INDEX IF NOT EXISTS idx_app_player
  ON appearance(season, club_id, player_id, game_id);
"""


def u21_db_path():
    return app_dir() / U21_DB_NAME


def u21_db_open(path):
    """台帳を開いて、必要なら表を作る

    途中で失敗したら必ず閉じる。開いたままにすると、Windowsでは呼び出し側が
    壊れたファイルを退避しようとしても「使用中」で動かせなくなる。
    """
    con = sqlite3.connect(path)
    try:
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA journal_mode=DELETE")
        version = con.execute("PRAGMA user_version").fetchone()[0]
        if version > U21_DB_VERSION:
            raise ScrapeError(
                f"U-21台帳が新しい形式です(台帳 v{version} / このアプリ v{U21_DB_VERSION})",
                "LoanWatch を新しい版に更新してください", str(path))
        con.executescript(U21_DB_SCHEMA)
        con.execute(f"PRAGMA user_version={U21_DB_VERSION}")
        con.commit()
    except BaseException:
        con.close()
        raise
    return con


def u21_db_connect():
    """台帳を開く。壊れていたら退避して作り直す

    台帳はJリーグ公式から再生成できるので、壊れた場合に失うのは取得の手間だけ。
    退避したファイルは消さずに残す(中身を見たいときのため)。
    """
    path = u21_db_path()
    try:
        return u21_db_open(path)
    except sqlite3.DatabaseError as e:
        broken = path.with_name(f"{path.stem}_破損_{datetime.now():%Y%m%d_%H%M%S}.db")
        log(f"U-21台帳を開けません({type(e).__name__}: {e}) → "
            f"{broken.name} へ退避して作り直します", "WARN")
        try:
            path.replace(broken)
        except OSError as move_error:
            raise ScrapeError("U-21台帳を退避できない",
                              f"{path} を手で削除してください", str(move_error))
        return u21_db_open(path)


def u21_db_recorded_games(con, season, club_id):
    """台帳に入っている試合IDの集合"""
    rows = con.execute(
        "SELECT game_id FROM game WHERE season=? AND club_id=?",
        (season, club_id)).fetchall()
    return {r["game_id"] for r in rows}


def u21_db_save_game(con, season, club_id, game, side, records):
    """1試合ぶんを台帳へ書く(取り直しのときは同じ試合の記録を入れ替える)"""
    with con:  # 例外が出れば丸ごと巻き戻る
        con.execute("DELETE FROM appearance WHERE season=? AND club_id=? AND game_id=?",
                    (season, club_id, game["game_id"]))
        con.execute(
            "INSERT OR REPLACE INTO game"
            "(season, club_id, game_id, date, side, opponent, result, url, fetched_at)"
            " VALUES(?,?,?,?,?,?,?,?,?)",
            (season, club_id, game["game_id"], game["date"].strftime("%Y-%m-%d"),
             side, game["away_short"] if side == "home" else game["home_short"],
             u21_result_text(game, side), game["url"],
             datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        con.executemany(
            "INSERT INTO appearance"
            "(season, club_id, game_id, player_id, name, no, pos,"
            " status, minutes_text, minutes, goals, yellow, red)"
            " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
            [(season, club_id, game["game_id"], r["player_id"], r["name"],
              r["no"], r["pos"], r["status"], r["minutes_text"], r["minutes"],
              r["goals"], r["yellow"], r["red"]) for r in records])


def u21_db_appearances(con, season, club_id):
    """その季・そのクラブの出場記録を、日付の古い順で返す"""
    return con.execute(
        "SELECT a.*, g.date, g.opponent, g.result, g.url"
        "  FROM appearance a"
        "  JOIN game g ON g.season=a.season AND g.club_id=a.club_id"
        "             AND g.game_id=a.game_id"
        " WHERE a.season=? AND a.club_id=?"
        " ORDER BY g.date, a.game_id",
        (season, club_id)).fetchall()


def u21_match_records(html, plain, game, side):
    """試合ページから、対象クラブのメンバー全員ぶんの記録を作る"""
    starters, members, subs, events = jl_match_members(html)
    if not members:
        raise ScrapeError(f"{U21}の試合ページからメンバー表を取得できない",
                          "--save-html でHTMLを保存し構造を確認する", game["url"])
    team_key = jl_match_team_keys(plain).get(side)
    if not team_key:
        raise ScrapeError(f"{U21}の試合ページからチームの識別子を取得できない",
                          "--save-html でHTMLを保存し構造を確認する", game["url"])

    records = []
    for item in u21_lineup_items(plain):
        if item["team_key"] != team_key:
            continue
        status, minutes_text, minutes = jl_member_status(
            item["id"], starters, members, subs)
        ev = events.get(item["id"]) or {}
        records.append({
            "player_id": item["id"], "name": item["name"],
            "no": item["no"], "pos": item["pos"],
            "status": status, "minutes_text": minutes_text, "minutes": minutes,
            "goals": ev.get("goal", 0),
            "yellow": ev.get("yellow-card", 0), "red": ev.get("red-card", 0),
        })
    if not records:
        raise ScrapeError(f"{U21}のメンバー表に対象クラブの選手が1人もいない",
                          f"チーム識別子 {team_key} の項目を確認", game["url"])
    return records


def u21_collect_games(con, season, club, games):
    """未記録の試合と、新しい方の数試合を取り直して台帳へ入れる

    1試合の取得に失敗しても台帳に入れない(入れると以後ずっと欠けたままになる)。
    戻り値は (取り込んだ試合数, 失敗した試合のリスト)。
    """
    recorded = u21_db_recorded_games(con, season, club["id"])
    recent = {g["game_id"] for g in games[-U21_REFETCH_RECENT:]}
    todo = [g for g in games if g["game_id"] not in recorded or g["game_id"] in recent]
    if not todo:
        debug(f"{club['name']}: 台帳は最新です(記録済み{len(recorded)}試合)")
        return 0, []

    log(f"{club['name']}: 台帳に{len(recorded)}試合 / 今回取得する試合 {len(todo)}件")
    done, failed = 0, []
    for game in todo:
        side = u21_team_side(game, club["name"])
        if side is None:
            failed.append((game, f"対戦カードに '{club['name']}' が無い"
                                 f"({game['home']} / {game['away']})"))
            continue
        try:
            html = fetch(game["url"])
            save_html(html, f"u21_match_{game['game_id']}")
            records = u21_match_records(html, jl_unescape(html), game, side)
            u21_db_save_game(con, season, club["id"], game, side, records)
        except ScrapeError as e:
            failed.append((game, e.summary()))
            continue
        except Exception as e:
            failed.append((game, f"{type(e).__name__}: {e}"))
            continue
        done += 1
        log(f"  {game['date']} {game['home']} "
            f"{game['home_goals']}-{game['away_goals']} {game['away']}"
            f" → メンバー{len(records)}人")
    for game, reason in failed:
        log(f"{club['name']}: {game['date']} の試合を取り込めません({reason})", "WARN")
    return done, failed


def u21_ledger_rows(con, season, club):
    """台帳から、選手ごとに1行(最新の関与試合＋今季の通算)を作る"""
    per_player = {}
    for rec in u21_db_appearances(con, season, club["id"]):
        acc = per_player.setdefault(rec["player_id"], {
            "メンバー入り": 0, "先発": 0, "途中出場": 0, "ベンチ入りのみ": 0,
            "通算出場時間(分)": 0})
        acc["メンバー入り"] += 1
        if rec["status"] in acc:
            acc[rec["status"]] += 1
        acc["通算出場時間(分)"] += rec["minutes"] or 0
        acc["latest"] = rec  # 日付の古い順なので、最後に入ったものが最新

    rows = []
    for acc in per_player.values():
        rec = acc["latest"]
        played = rec["status"] in (STARTER, SUBBED_IN)
        no = rec["no"]
        rows.append({
            # 背番号は試合ごとに変わりうるので、最新の関与試合の値を使う
            "背番号": int(no) if str(no).isdigit() else (no or ""),
            "ポジション": rec["pos"], "選手": rec["name"],
            "クラブ": club["name"], "リーグ名": U21,
            "最新試合関与試合": rec["date"], "経過": format_elapsed(rec["date"]),
            "ステータス": rec["status"], "出場時間": rec["minutes_text"] or "",
            "対戦相手": rec["opponent"], "試合結果": rec["result"],
            "ゴール数": str(rec["goals"]) if played else "",
            "警告/退場": f"{rec['yellow']}/{rec['red']}" if played else "",
            "メンバー入り": acc["メンバー入り"], "先発": acc["先発"],
            "途中出場": acc["途中出場"], "ベンチ入りのみ": acc["ベンチ入りのみ"],
            "通算出場時間(分)": acc["通算出場時間(分)"],
            "備考": f"Jリーグ公式({U21})", "URL": rec["url"],
        })
    rows.sort(key=lambda r: (U21_POS_ORDER.get(r["ポジション"], 9),
                             r["背番号"] if isinstance(r["背番号"], int) else 999,
                             r["選手"]))
    return rows


def u21_team_rows(club):
    """指定したU-21クラブの台帳を最新にして、シート用の行を作る"""
    season = CFG["year"]
    games = jl_u21_games(club["id"])
    con = u21_db_connect()
    try:
        if games:
            u21_collect_games(con, season, club, games)
        else:
            log(f"{club['name']}: {season}シーズンの{U21}に消化済みの試合がありません",
                "WARN")
        rows = u21_ledger_rows(con, season, club)
    finally:
        con.close()
    if not rows:
        debug(f"{club['name']}: 台帳に{season}シーズンの記録がありません")
    else:
        debug(f"{club['name']}: 台帳から{len(rows)}人ぶんの行を作成")
    return rows


def u21_team_sheet_rows(clubs):
    """対象クラブそれぞれの台帳を更新して、U-21シート用の行にする

    1クラブ失敗しても残りは出す。レンタル選手側の出力は止めない。
    """
    rows, errors = [], []
    for club in clubs:
        print(f"■ {club['name']} [{U21}]")
        try:
            rows.extend(u21_team_rows(club))
        except ScrapeError as e:
            log(e.summary(), "ERROR")
            errors.append((club["name"], e.summary()))
        except Exception as e:
            log(f"想定外のエラー: {type(e).__name__}: {e}", "ERROR")
            if CFG.get("debug"):
                traceback.print_exc()
            errors.append((club["name"], f"{type(e).__name__}: {e}"))
    return rows, errors


def minutes_to_entry(minutes):
    """出場時間から出場時刻を逆算し、"後半15分" のような表記にする

    メンバー表を読めない大会(出場履歴しか無い場合)は分数しか分からないため、
    そこから入った時刻を逆算する。メンバー表がある大会は jl_entry_label を
    イベントの時刻に直接当てるので、こちらは使わない。
    """
    if not minutes or minutes >= FULL_MATCH_MINUTES:
        return None, None
    entry = FULL_MATCH_MINUTES - minutes
    return entry, jl_entry_label(entry)


def scrape_jleague(player, note_prefix=""):
    df, url = jl_history_df(player)
    if df.empty:
        return blank(player, NO_MATCH, f"{note_prefix}出場履歴に試合行がない")

    latest = df.iloc[-1]
    debug(f"出場履歴 {len(df)}行、最新={latest['_date']}")

    status, minutes_text = jl_status_from_row(latest, df.columns)
    row = result(player, status, match=jl_value(latest, df.columns, JL_COL_OPP),
                 date=latest["_date"].strftime("%Y-%m-%d"), url=url,
                 note=(note_prefix + "Jリーグ公式").strip())
    row.update(jl_extra_fields(latest, df.columns))
    row["出場時間"] = minutes_text
    return row


# =========================================================
# BeSoccer (リザーブチームなど、ESPN・Football-LAB に載らない大会)
#
# ESPN も Football-LAB もトップチームの記録しか持たないため、リザーブチームの
# 出場(例: アルヘンティノス・ジュニアーズII のトルネオ・プロジェクシオン)は
# BeSoccer の選手ページ /player/matches/<名前>-<選手ID> から拾う。
# 1回の取得で直近100試合(全大会・全シーズン)が入っているので、大会が増えても
# 追加のリクエストは要らない。
#
# 試合一覧の1行の構造(2026年8月に確認):
#   <tr class="row-body">
#     <td class="tiny"><img alt="Reserve League"
#          src=".../league_logos/torneo-proyeccion.png?size=60x&lossy=1"></td>
#     <td class="color-cell bg-match-res win|draw|lose">19 Aug</td>
#     <td class="match-cell"><a class="match"
#          href=".../match/ca-aldosivi-reservas/argentinos-juniors-reservas/2026347773">
#          ALD <img alt="Aldosivi II"> 1-0 <img alt="Argentinos Juniors II"> ARG</a></td>
#     <td class="tiny">46'</td>
#     <td class="events-cell"><div class="img-ico event-18"></div></td>
#   </tr>
#
# * 大会は行頭のリーグロゴのファイル名で見分ける。
#   torneo-proyeccion = リザーブリーグ / apertura-argentina = トップチーム。
# * 勝敗は日付セルの win/draw/lose クラスに対象選手のチーム視点で入っている。
# * 日付セルに年が無いので、試合URL末尾のID先頭4桁(=開催年)から補う。
# * 出場時間セルは 空欄=メンバー表に無い / 0'=ベンチ入りのみ / それ以外=出場分数。
# * イベントアイコンの番号(event-<n>)は CSS スプライト sprite_events_new.png の
#   表示位置から種別を確認した。1=ゴール 3=退場 4=警告2回で退場 5=警告
#   18=交代OUT 19=交代IN 20=負傷。
#
# 同じ試合一覧が「イベントと得点」(#event)と「ポジションとELO」(#rating)の
# 2タブ分あり、出場時間とイベントは #event 側にしか無い。
# =========================================================
BESOCCER = "BeSoccer"
BS_GOAL, BS_RED, BS_SECOND_YELLOW, BS_YELLOW = "1", "3", "4", "5"
BS_SUB_OUT, BS_SUB_IN, BS_INJURY = "18", "19", "20"

RE_BS_LEAGUE_LOGO = re.compile(r"/league_logos/([^/.?]+)")
RE_BS_MATCH_ID = re.compile(r"/(\d{4})\d+/?$")
RE_BS_EVENT = re.compile(r"\bevent-(\d+)\b")
RE_BS_DAY_MONTH = re.compile(r"(\d{1,2})\s*([A-Za-z]{3})")
RE_BS_SCORE = re.compile(r"^(\d+)\s*-\s*(\d+)$")
RE_BS_SCORE_PK = re.compile(r"^(\d+)\s*\((\d+)\s*-\s*(\d+)\)\s*(\d+)$")
BS_MONTHS = {m: i for i, m in enumerate(
    ("jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"), start=1)}

# BeSoccer はボット対策が厳しいため、requestsを先に投げて失敗後に
# Playwrightへ切り替える方式は使わない。最初から通常のGoogle Chromeを
# Persistent Contextで起動し、同じ専用プロファイルを継続利用する。
# User-Agent / Client Hintsは固定値で上書きせず、実Chrome自身の値を使用する。
# 短時間に何度も叩くとCAPTCHA(Client Challenge)に飛ばされることがあるため、
# 1回の実行で選手ページは1度しか取りに行かない。
RE_BS_CHALLENGE = re.compile(r"Client Challenge|_Incapsula_Resource|captcha|incapsula|imperva", re.I)

# 取得したHTMLは cache/ に残す。
# * HTML … ブロックされた回に前回分で行を出せるようにする(備考へ取得日時を入れる)。
#           取り立てなら取りに行かないので、続けて実行してもアクセスは増えない。
# * diagnostics … --debug時に403/406/CAPTCHA等の原因調査材料を保存する。
#   Cookie/Authorization/Storageの「値」は保存せず、Cookie名・Storageキー名だけ残す。
BS_CACHE_DIR = app_dir() / "cache"
BS_CACHE_FRESH_SEC = 30 * 60
# BeSoccer専用のPersistent Profile。Cookie / Local Storage等を実行をまたいで再利用する。
# Playwright ChromiumはChromeと同じプロファイルを共有できないため、フォールバック時だけ子ディレクトリを使う。
BS_PROFILE_DIR = Path(os.environ.get("LOCALAPPDATA", app_dir())) / "LoanWatch"
BS_CHROMIUM_PROFILE_DIR = BS_PROFILE_DIR / "chromium"
BS_DIAG_DIR = BS_CACHE_DIR / "besoccer_diagnostics"
_BS_HTML_CACHE = {}

# ログへ値を残さないヘッダー。Set-Cookie/Cookieは名前だけを抽出する。
BS_SECRET_HEADER_NAMES = {
    "authorization", "proxy-authorization", "x-api-key", "x-auth-token",
    "x-csrf-token", "x-xsrf-token",
}
BS_SECRET_QUERY_RE = re.compile(
    r"(?:token|auth|key|secret|session|sid|captcha|challenge|signature|sig|hash|nonce)", re.I
)


def besoccer_cache_path(url):
    digest = hashlib.sha1(url.encode("utf-8")).hexdigest()[:12]
    return BS_CACHE_DIR / f"besoccer_{digest}.html"


def besoccer_cache_read(url):
    """(HTML, 取得日時) を返す。まだ無ければ (None, None)"""
    path = besoccer_cache_path(url)
    try:
        if not path.is_file():
            return None, None
        return (path.read_text(encoding="utf-8"),
                datetime.fromtimestamp(path.stat().st_mtime))
    except OSError as e:
        debug(f"BeSoccerのキャッシュを読めません: {type(e).__name__}: {e}")
        return None, None


def besoccer_cache_write(url, html):
    try:
        BS_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        besoccer_cache_path(url).write_text(html, encoding="utf-8")
    except OSError as e:
        debug(f"BeSoccerのキャッシュを書けません: {type(e).__name__}: {e}")


def besoccer_fetch(url):
    """BeSoccerの選手ページを (HTML, キャッシュの取得日時) で返す

    同じURLは1回の実行で1度しか取りに行かない。取り立てのキャッシュがあれば
    それを使い、取りに行って弾かれた場合も前回分があればそちらで行を出す。
    """
    if url in _BS_HTML_CACHE:
        debug(f"BeSoccerのHTMLは取得済み: {url}")
        return _BS_HTML_CACHE[url]

    cached, fetched_at = besoccer_cache_read(url)
    if cached and (datetime.now() - fetched_at).total_seconds() < BS_CACHE_FRESH_SEC:
        log(f"BeSoccerは{fetched_at:%m/%d %H:%M}に取得済みのため、そのまま使います")
        return _BS_HTML_CACHE.setdefault(url, (cached, fetched_at))

    try:
        html = besoccer_download(url)
    except ScrapeError as e:
        if not cached:
            raise
        log(f"BeSoccerを取得できません({e.message}) → "
            f"{fetched_at:%m/%d %H:%M}に取得した内容で行を出します", "WARN")
        return _BS_HTML_CACHE.setdefault(url, (cached, fetched_at))

    besoccer_cache_write(url, html)
    return _BS_HTML_CACHE.setdefault(url, (html, None))


def _bs_json_default(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    return str(value)


def _bs_write_json(path, value):
    try:
        path.write_text(json.dumps(value, ensure_ascii=False, indent=2,
                                   default=_bs_json_default), encoding="utf-8")
    except Exception as e:
        debug(f"BeSoccer診断JSONを書けません({path.name}): {type(e).__name__}: {e}")


def _bs_write_text(path, value):
    try:
        path.write_text(str(value), encoding="utf-8", errors="replace")
    except Exception as e:
        debug(f"BeSoccer診断テキストを書けません({path.name}): {type(e).__name__}: {e}")


def _bs_cookie_names_from_cookie_header(value):
    """Cookie: a=1; b=2 から値を捨ててCookie名だけ返す。"""
    names = []
    for part in str(value or "").split(";"):
        name = part.split("=", 1)[0].strip()
        if name and name.lower() not in {"path", "domain", "expires", "max-age",
                                        "secure", "httponly", "samesite"}:
            names.append(name)
    return sorted(set(names))


def _bs_set_cookie_name(value):
    """Set-Cookie 1行からCookie名だけ返す。"""
    first = str(value or "").split(";", 1)[0]
    return first.split("=", 1)[0].strip() if "=" in first else ""


def _bs_sanitize_header_array(headers):
    """Playwright headers_array() の値から認証情報を除去する。"""
    out = []
    for h in headers or []:
        name = str(h.get("name", ""))
        value = str(h.get("value", ""))
        lower = name.lower()
        if lower == "cookie":
            value = f"<redacted; cookie_names={_bs_cookie_names_from_cookie_header(value)}>"
        elif lower == "set-cookie":
            cookie_name = _bs_set_cookie_name(value)
            value = f"<redacted; cookie_name={cookie_name or 'unknown'}>"
        elif lower in BS_SECRET_HEADER_NAMES or "password" in lower or "secret" in lower:
            value = "<redacted>"
        out.append({"name": name, "value": value})
    return out


def _bs_safe_url(url):
    """診断ログ用URL。疑わしいquery値だけ伏せ、パラメータ名は残す。"""
    try:
        from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit
        parts = urlsplit(str(url))
        query = []
        for key, value in parse_qsl(parts.query, keep_blank_values=True):
            query.append((key, "<redacted>" if BS_SECRET_QUERY_RE.search(key) else value))
        return urlunsplit((parts.scheme, parts.netloc, parts.path,
                           urlencode(query, doseq=True), parts.fragment))
    except Exception:
        return str(url)


def _bs_profile_summary(profile_dir=None):
    """プロファイルの中身は読まず、存在・ファイル数・更新日時だけ診断材料にする。"""
    profile_dir = Path(profile_dir or BS_PROFILE_DIR)
    info = {"path": str(profile_dir), "exists": profile_dir.exists()}
    if not profile_dir.exists():
        return info
    try:
        files = [x for x in profile_dir.rglob("*") if x.is_file()]
        info["file_count"] = len(files)
        info["total_bytes"] = sum(x.stat().st_size for x in files)
        mtimes = [x.stat().st_mtime for x in files]
        if mtimes:
            info["newest_file_mtime"] = datetime.fromtimestamp(max(mtimes)).isoformat()
        # ファイル名だけ。Cookie DB等の内容は読み出さない。
        info["top_level_entries"] = sorted(x.name for x in profile_dir.iterdir())[:100]
    except Exception as e:
        info["error"] = f"{type(e).__name__}: {e}"
    return info


def _bs_chrome_unavailable(exc):
    """Google Chromeが未インストール/実行ファイル不在と判断できるPlaywright例外か。

    プロファイルロックや起動直後クラッシュなど、Chromeが「存在するのに起動できない」
    障害までChromiumへ隠れてフォールバックしないよう、明確な不在メッセージだけを対象にする。
    """
    text = str(exc).lower()
    needles = (
        "distribution 'chrome' is not found",
        'distribution "chrome" is not found',
        "chrome distribution is not found",
        "google chrome is not installed",
        "chrome is not installed",
        "could not find chrome",
        "chrome executable not found",
        "chrome not found",
        "executable doesn't exist",
        "executable does not exist",
        "enoent",
    )
    return any(n in text for n in needles)


async def _bs_browser_environment(page):
    """ブラウザ自身がJSへ公開している環境情報を取得する。偽装は行わない。"""
    script = r"""
    async () => {
      const result = {
        userAgent: navigator.userAgent,
        appVersion: navigator.appVersion,
        platform: navigator.platform,
        vendor: navigator.vendor,
        webdriver: navigator.webdriver,
        language: navigator.language,
        languages: Array.from(navigator.languages || []),
        cookieEnabled: navigator.cookieEnabled,
        doNotTrack: navigator.doNotTrack,
        hardwareConcurrency: navigator.hardwareConcurrency,
        deviceMemory: navigator.deviceMemory,
        maxTouchPoints: navigator.maxTouchPoints,
        pdfViewerEnabled: navigator.pdfViewerEnabled,
        onLine: navigator.onLine,
        windowChromePresent: !!window.chrome,
        screen: {
          width: screen.width, height: screen.height,
          availWidth: screen.availWidth, availHeight: screen.availHeight,
          colorDepth: screen.colorDepth, pixelDepth: screen.pixelDepth,
          devicePixelRatio: window.devicePixelRatio,
          innerWidth: window.innerWidth, innerHeight: window.innerHeight,
          outerWidth: window.outerWidth, outerHeight: window.outerHeight
        },
        timezone: Intl.DateTimeFormat().resolvedOptions(),
        plugins: Array.from(navigator.plugins || []).map(p => ({name:p.name, filename:p.filename})),
        mimeTypes: Array.from(navigator.mimeTypes || []).map(m => m.type),
      };
      if (navigator.userAgentData) {
        result.userAgentData = {
          brands: navigator.userAgentData.brands,
          mobile: navigator.userAgentData.mobile,
          platform: navigator.userAgentData.platform,
        };
        if (navigator.userAgentData.getHighEntropyValues) {
          try {
            result.userAgentData.highEntropy = await navigator.userAgentData.getHighEntropyValues([
              'architecture','bitness','formFactors','fullVersionList','model',
              'platformVersion','uaFullVersion','wow64'
            ]);
          } catch (e) {
            result.userAgentData.highEntropyError = String(e);
          }
        }
      }
      try {
        result.storage = {
          localStorageKeys: Object.keys(localStorage),
          sessionStorageKeys: Object.keys(sessionStorage)
        };
      } catch (e) {
        result.storage = {error:String(e)};
      }
      try {
        result.indexedDB = indexedDB.databases
          ? (await indexedDB.databases()).map(db => ({name:db.name, version:db.version}))
          : {supported:false};
      } catch (e) {
        result.indexedDB = {error:String(e)};
      }
      try {
        result.serviceWorkers = navigator.serviceWorker
          ? (await navigator.serviceWorker.getRegistrations()).map(r => ({
              scope:r.scope,
              active:r.active ? r.active.scriptURL : null,
              waiting:r.waiting ? r.waiting.scriptURL : null,
              installing:r.installing ? r.installing.scriptURL : null
            }))
          : [];
      } catch (e) {
        result.serviceWorkers = {error:String(e)};
      }
      return result;
    }
    """
    try:
        return await page.evaluate(script)
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"}


async def _bs_cookie_metadata(ctx, url):
    """Cookie値は絶対に保存せず、WAF系Cookieの存在を確認できるメタ情報だけ返す。"""
    try:
        cookies = await ctx.cookies([url])
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"}
    out = []
    for c in cookies:
        out.append({
            "name": c.get("name"),
            "domain": c.get("domain"),
            "path": c.get("path"),
            "expires": c.get("expires"),
            "httpOnly": c.get("httpOnly"),
            "secure": c.get("secure"),
            "sameSite": c.get("sameSite"),
            "value": "<redacted>",
        })
    return out


async def _bs_save_diagnostics(page, ctx, requested_url, response, reason,
                               network_events, failed_requests, console_events,
                               page_errors, profile_before, browser_channel="chrome",
                               browser_source="installed_google_chrome", profile_dir=None,
                               fallback_reason=""):
    """403/406/CAPTCHA等の診断材料をcache/besoccer_diagnosticsへ保存する。

    CAPTCHA突破のための値は収集せず、Cookie/Authorization/Storage値は伏せる。
    戻り値は作成した診断ディレクトリ(Path)。
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
    status = response.status if response else 0
    digest = hashlib.sha1(requested_url.encode("utf-8")).hexdigest()[:8]
    safe_reason = re.sub(r"[^A-Za-z0-9_-]+", "_", str(reason))[:40].strip("_") or "diagnostic"
    diag_dir = BS_DIAG_DIR / f"{timestamp}_{status}_{safe_reason}_{digest}"
    try:
        diag_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        debug(f"BeSoccer診断フォルダを作成できません: {type(e).__name__}: {e}")
        return None

    # JSチャレンジによるDOM更新や遷移直後の状態も拾うため短く待つ。
    try:
        await page.wait_for_timeout(800)
    except Exception:
        pass

    try:
        final_url = page.url
    except Exception:
        final_url = ""
    try:
        title = await page.title()
    except Exception as e:
        title = f"<title取得失敗: {type(e).__name__}: {e}>"

    request_headers = []
    response_headers = []
    security_details = None
    server_addr = None
    status_text = ""
    if response:
        try:
            request_headers = _bs_sanitize_header_array(await response.request.headers_array())
        except Exception as e:
            request_headers = [{"error": f"{type(e).__name__}: {e}"}]
        try:
            response_headers = _bs_sanitize_header_array(await response.headers_array())
        except Exception as e:
            response_headers = [{"error": f"{type(e).__name__}: {e}"}]
        try:
            security_details = await response.security_details()
        except Exception as e:
            security_details = {"error": f"{type(e).__name__}: {e}"}
        try:
            server_addr = await response.server_addr()
        except Exception as e:
            server_addr = {"error": f"{type(e).__name__}: {e}"}
        try:
            status_text = response.status_text
        except Exception:
            pass

    # レスポンスのraw body。challenge token等が含まれる可能性はあるため、
    # 外部送信はせずローカル診断フォルダだけに保存する。
    body_info = {"saved": False}
    if response:
        try:
            body = await response.body()
            (diag_dir / "response_body.bin").write_bytes(body)
            body_info = {"saved": True, "bytes": len(body), "sha256": hashlib.sha256(body).hexdigest()}
            # HTML/textとして読める場合の確認用コピー。
            try:
                decoded = body.decode("utf-8", errors="replace")
                _bs_write_text(diag_dir / "response_body.html", decoded)
            except Exception:
                pass
        except Exception as e:
            body_info = {"saved": False, "error": f"{type(e).__name__}: {e}"}

    try:
        dom = await page.content()
        _bs_write_text(diag_dir / "dom.html", dom)
        dom_info = {"chars": len(dom), "sha256": hashlib.sha256(dom.encode("utf-8", errors="replace")).hexdigest()}
    except Exception as e:
        dom_info = {"error": f"{type(e).__name__}: {e}"}

    screenshot_info = {"saved": False}
    try:
        screenshot_path = diag_dir / "screenshot.png"
        await page.screenshot(path=str(screenshot_path), full_page=True)
        screenshot_info = {"saved": True, "path": screenshot_path.name}
    except Exception as e:
        screenshot_info = {"saved": False, "error": f"{type(e).__name__}: {e}"}

    browser_environment = await _bs_browser_environment(page)
    cookie_metadata = await _bs_cookie_metadata(ctx, final_url or requested_url)

    try:
        playwright_version = importlib_metadata.version("playwright")
    except Exception:
        playwright_version = "unknown"

    summary = {
        "captured_at_local": datetime.now().astimezone().isoformat(),
        "reason": reason,
        "navigation_status": status,
        "navigation_status_text": status_text,
        "requested_url": _bs_safe_url(requested_url),
        "final_url": _bs_safe_url(final_url),
        "title": title,
        "headless": False,
        "channel": browser_channel,
        "browser_source": browser_source,
        "profile_dir": str(profile_dir or BS_PROFILE_DIR),
        "fallback_reason": fallback_reason,
        "locale": "en-US",
        "timezone_id": "Asia/Tokyo",
        "viewport": {"width": 1440, "height": 1000},
        "python": {
            "version": sys.version,
            "executable": sys.executable,
            "implementation": platform.python_implementation(),
        },
        "os": {
            "platform": platform.platform(),
            "system": platform.system(),
            "release": platform.release(),
            "version": platform.version(),
            "machine": platform.machine(),
        },
        "app_version": globals().get("APP_VERSION", "unknown"),
        "playwright_version": playwright_version,
        "profile_before_launch": profile_before,
        "profile_after_capture": _bs_profile_summary(profile_dir),
        "response_body": body_info,
        "dom": dom_info,
        "screenshot": screenshot_info,
        "security_details": security_details,
        "server_addr": server_addr,
        "network_event_count": len(network_events),
        "failed_request_count": len(failed_requests),
        "console_event_count": len(console_events),
        "page_error_count": len(page_errors),
        "privacy_note": (
            "Cookie/Authorization/Storageの値は保存していません。"
            "response_bodyにはサイトが返したchallenge用の一時値が含まれる可能性があります。"
        ),
    }

    _bs_write_json(diag_dir / "summary.json", summary)
    _bs_write_json(diag_dir / "request_headers.json", request_headers)
    _bs_write_json(diag_dir / "response_headers.json", response_headers)
    _bs_write_json(diag_dir / "browser_environment.json", browser_environment)
    _bs_write_json(diag_dir / "cookies_metadata.json", cookie_metadata)
    _bs_write_json(diag_dir / "network.json", network_events)
    _bs_write_json(diag_dir / "request_failures.json", failed_requests)
    _bs_write_json(diag_dir / "console.json", console_events)
    _bs_write_json(diag_dir / "page_errors.json", page_errors)

    # 人が最初に見るための簡易テキスト。
    ua = browser_environment.get("userAgent", "") if isinstance(browser_environment, dict) else ""
    webdriver = browser_environment.get("webdriver", "") if isinstance(browser_environment, dict) else ""
    ua_data = browser_environment.get("userAgentData", {}) if isinstance(browser_environment, dict) else {}
    hint_lines = [
        f"captured_at: {summary['captured_at_local']}",
        f"reason: {reason}",
        f"status: {status} {status_text}",
        f"requested_url: {_bs_safe_url(requested_url)}",
        f"final_url: {_bs_safe_url(final_url)}",
        f"title: {title}",
        f"headless: {summary['headless']}",
        f"browser_channel: {browser_channel}",
        f"browser_source: {browser_source}",
        f"profile_dir: {profile_dir or BS_PROFILE_DIR}",
        f"fallback_reason: {fallback_reason or '(none)'}",
        f"Playwright: {playwright_version}",
        f"navigator.userAgent: {ua}",
        f"navigator.webdriver: {webdriver}",
        f"navigator.userAgentData: {json.dumps(ua_data, ensure_ascii=False, default=_bs_json_default)}",
        f"server_addr: {json.dumps(server_addr, ensure_ascii=False, default=_bs_json_default)}",
        f"security_details: {json.dumps(security_details, ensure_ascii=False, default=_bs_json_default)}",
        f"network_events: {len(network_events)}",
        f"request_failures: {len(failed_requests)}",
        "",
        "詳細は summary.json / request_headers.json / response_headers.json /",
        "browser_environment.json / cookies_metadata.json / network.json / screenshot.png /",
        "response_body.html / dom.html を確認してください。",
        "Cookie/Authorization/Storageの値は意図的に保存していません。",
    ]
    _bs_write_text(diag_dir / "README.txt", "\n".join(hint_lines))
    log(f"BeSoccer診断情報を保存しました: {diag_dir}", "WARN")
    return diag_dir


async def _render_besoccer(url):
    """BeSoccer専用ブラウザ取得。

    1. まずPCにインストール済みのGoogle Chrome(channel="chrome")を使う。
    2. Chromeが明確に未インストール/実行ファイル不在の場合だけ、Playwright同梱の
       Chromium(channel="chromium")へフォールバックする。
    3. UA / Client Hintsはどちらも固定値で上書きせず、実ブラウザ自身の値を使う。
    4. Chrome/ChromiumでUser Data Directoryを分離し、プロファイル互換性問題を避ける。
    5. --debug時は成功/失敗とも実際に使用したchannel/sourceを診断情報へ保存する。
    """
    chrome_profile = BS_PROFILE_DIR
    chromium_profile = BS_CHROMIUM_PROFILE_DIR

    try:
        chrome_profile.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        raise ScrapeError(
            "BeSoccer用Chromeプロファイルを作成できません",
            "cacheフォルダへの書き込み権限を確認",
            str(chrome_profile),
        ) from e

    async with async_playwright() as pw:
        browser_channel = "chrome"
        browser_source = "installed_google_chrome"
        profile_dir = chrome_profile
        fallback_reason = ""
        profile_before = _bs_profile_summary(profile_dir)

        try:
            ctx = await pw.chromium.launch_persistent_context(
                user_data_dir=str(profile_dir),
                channel="chrome",
                headless=False,
                locale="en-US",
                timezone_id="Asia/Tokyo",
                viewport={"width": 1440, "height": 1000},
            )
            debug("BeSoccerブラウザ: インストール済みGoogle Chrome (channel=chrome)")
        except Exception as chrome_error:
            if not _bs_chrome_unavailable(chrome_error):
                # Chromeは存在している可能性が高い。プロファイルロックやポリシー等を
                # Chromiumフォールバックで隠さず、そのまま原因を通知する。
                raise ScrapeError(
                    "Google Chromeを起動できません",
                    "Chromeのプロファイルロック、企業ポリシー、実行権限を確認",
                    str(chrome_error),
                ) from chrome_error

            fallback_reason = f"{type(chrome_error).__name__}: {chrome_error}"
            log("Google Chromeが見つからないため、Playwright Chromiumへ切り替えます", "WARN")
            debug(f"Chrome起動失敗詳細: {fallback_reason}")

            browser_channel = "chromium"
            browser_source = "playwright_bundled_chromium"
            profile_dir = chromium_profile
            try:
                profile_dir.mkdir(parents=True, exist_ok=True)
            except OSError as e:
                raise ScrapeError(
                    "BeSoccer用Chromiumプロファイルを作成できません",
                    "%LOCALAPPDATA%\\LoanWatch への書き込み権限を確認",
                    str(profile_dir),
                ) from e
            profile_before = _bs_profile_summary(profile_dir)

            try:
                ctx = await pw.chromium.launch_persistent_context(
                    user_data_dir=str(profile_dir),
                    channel="chromium",
                    headless=False,
                    locale="en-US",
                    timezone_id="Asia/Tokyo",
                    viewport={"width": 1440, "height": 1000},
                )
                log("BeSoccerブラウザ: Playwright Chromium (channel=chromium)", "WARN")
            except Exception as chromium_error:
                command = f'"{sys.executable}" -m playwright install chromium'
                raise ScrapeError(
                    "Google ChromeもPlaywright Chromiumも起動できません",
                    f"Chromeをインストールするか、同梱/インストール済みChromiumを確認。開発環境では {command}",
                    f"Chrome: {chrome_error}\nChromium: {chromium_error}",
                ) from chromium_error

        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        # ナビゲーションの実HTTPステータスを控えておく。本文が空のエラー応答
        # (BeSoccerの406など)は、Chromeがレスポンスではなくナビゲーション失敗
        # として扱うため page.goto から resp を受け取れない。そのときに
        # 403/406 の正規経路へ戻すための材料になる。--debug の有無に関わらず必要。
        nav_status = {"code": None}

        def on_navigation_response(res):
            try:
                if res.request.is_navigation_request():
                    nav_status["code"] = res.status
            except Exception:
                pass

        page.on("response", on_navigation_response)

        # --debug時だけネットワーク/Consoleを収集する。Cookie値等はイベントには保存しない。
        network_events = []
        failed_requests = []
        console_events = []
        page_errors = []
        if CFG.get("debug"):
            def on_request(req):
                try:
                    network_events.append({
                        "event": "request",
                        "time": datetime.now().astimezone().isoformat(),
                        "method": req.method,
                        "resource_type": req.resource_type,
                        "navigation": req.is_navigation_request(),
                        "url": _bs_safe_url(req.url),
                    })
                except Exception:
                    pass

            def on_response(res):
                try:
                    network_events.append({
                        "event": "response",
                        "time": datetime.now().astimezone().isoformat(),
                        "status": res.status,
                        "status_text": res.status_text,
                        "resource_type": res.request.resource_type,
                        "url": _bs_safe_url(res.url),
                    })
                except Exception:
                    pass

            def on_request_failed(req):
                try:
                    failed_requests.append({
                        "time": datetime.now().astimezone().isoformat(),
                        "method": req.method,
                        "resource_type": req.resource_type,
                        "url": _bs_safe_url(req.url),
                        "failure": req.failure,
                    })
                except Exception:
                    pass

            def on_console(msg):
                try:
                    console_events.append({
                        "time": datetime.now().astimezone().isoformat(),
                        "type": msg.type,
                        "text": msg.text,
                    })
                except Exception:
                    pass

            def on_page_error(exc):
                try:
                    page_errors.append({
                        "time": datetime.now().astimezone().isoformat(),
                        "error": str(exc),
                    })
                except Exception:
                    pass

            page.on("request", on_request)
            page.on("response", on_response)
            page.on("requestfailed", on_request_failed)
            page.on("console", on_console)
            page.on("pageerror", on_page_error)

        async def save_diag(response, reason):
            if not CFG.get("debug"):
                return None
            return await _bs_save_diagnostics(
                page, ctx, url, response, reason,
                network_events, failed_requests, console_events,
                page_errors, profile_before,
                browser_channel=browser_channel,
                browser_source=browser_source,
                profile_dir=profile_dir,
                fallback_reason=fallback_reason,
            )

        try:
            try:
                resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
                status = resp.status if resp else 0
            except Exception as goto_error:
                # ERR_HTTP_RESPONSE_CODE_FAILURE は406専用ではなく404などでも起きる。
                # 決め打ちせず、responseイベントで控えた実ステータスで判定する。
                if "ERR_HTTP_RESPONSE_CODE_FAILURE" not in str(goto_error):
                    raise
                resp, status = None, nav_status["code"]
                debug(f"page.goto がHTTPエラー応答で失敗。"
                      f"responseイベントのステータス={status}")
                if status is None:
                    # ステータスを取れなかった。原因は確定できないが、HTTPエラー応答が
                    # 返っている以上すぐ叩き直しても通らない。リトライせず終える。
                    diag_dir = await save_diag(None, "http_response_code_failure")
                    hint = "BeSoccerがHTTPエラー応答を返した可能性。時間をおいて再実行する"
                    if diag_dir:
                        hint += f"。診断情報: {diag_dir}"
                    raise ScrapeError(
                        "BeSoccerからHTTPエラー応答(ステータス不明)", hint, url,
                    ) from goto_error

            if status == 404:
                await save_diag(resp, "http_404")
                raise ScrapeError("HTTP 404 ページなし", "BeSoccerの選手URLを確認", url)
            if status in (403, 406):
                diag_dir = await save_diag(resp, f"http_{status}")
                hint = "時間をおいて再実行する。必要なら --headed で画面を確認"
                if diag_dir:
                    hint += f"。診断情報: {diag_dir}"
                raise ScrapeError(
                    f"HTTP {status} BeSoccerのボット確認でブロック",
                    hint,
                    url,
                )
            if status >= 400:
                diag_dir = await save_diag(resp, f"http_{status}")
                hint = "BeSoccerの選手URLを確認"
                if diag_dir:
                    hint += f"。診断情報: {diag_dir}"
                raise ScrapeError(f"HTTP {status}", hint, url)

            # Cookie同意バナーがあれば閉じる。
            for sel in ("#onetrust-accept-btn-handler",
                        "button:has-text('I Accept')", "button:has-text('Accept')"):
                try:
                    await page.click(sel, timeout=1500)
                    debug(f"BeSoccer Cookie同意をクリック: {sel}")
                    break
                except Exception:
                    pass

            try:
                await page.wait_for_selector("tr.row-body", timeout=10000)
            except Exception:
                log("BeSoccerの試合行(tr.row-body)が現れませんでした", "WARN")

            html = await page.content()
            if RE_BS_CHALLENGE.search(html[:20000]):
                diag_dir = await save_diag(resp, "challenge_in_dom")
                hint = "時間をおいて再実行する。必要なら --headed で画面を確認"
                if diag_dir:
                    hint += f"。診断情報: {diag_dir}"
                raise ScrapeError(
                    "BeSoccerのボット確認ページ(CAPTCHA)が返ってきた",
                    hint,
                    url,
                )

            await save_diag(resp, "success")
            return html
        except ScrapeError:
            raise
        except Exception as e:
            diag_dir = None
            try:
                diag_dir = await save_diag(None, f"exception_{type(e).__name__}")
            except Exception:
                pass
            if diag_dir:
                debug(f"BeSoccer例外時の診断情報: {diag_dir}")
            raise
        finally:
            # Persistent Contextの状態は選択したuser_data_dirへ自動的に保存される。
            await ctx.close()

def render_besoccer(url):
    """_render_besoccer を同期的に呼ぶ。"""
    last = None
    for attempt in range(1, CFG["retry"] + 1):
        try:
            return asyncio.run(_render_besoccer(url))
        except ScrapeError:
            # HTTP 403/406やCAPTCHAは連続リトライしない。
            raise
        except Exception as e:
            last = e
            if attempt < CFG["retry"]:
                log(f"BeSoccerブラウザ取得に失敗({type(e).__name__}) → "
                    f"{attempt}回目リトライ", "WARN")
                time.sleep(CFG["sleep"] * 2)
    # Playwrightの例外は Call log が続く複数行になる。Excelの備考へそのまま
    # 入ると列が読めなくなるので、1行目だけを出し、全文は detail(画面表示)へ回す。
    head = str(last).splitlines()[0].strip() if last else ""
    raise ScrapeError(
        f"BeSoccerブラウザ取得に失敗: {type(last).__name__}: {head}",
        "Google ChromeとPlaywrightの実行環境を確認",
        f"{url}\n{last}",
    )


def besoccer_download(url):
    """BeSoccerはrequestsを経由せず、最初から実Chromeで取得する。"""
    return render_besoccer(url)


def besoccer_match_table(soup):
    """出場時間とイベントアイコンが入っているタブ(#event)を選ぶ"""
    panel = soup.select_one("div.pl-matches") or soup
    for tab in panel.select("div.tab-content"):
        if tab.select_one("td.events-cell"):
            return tab
    return panel.select_one("#event") or panel


def besoccer_day_month(date_cell):
    """'19 Aug' の日付セルを (日, 月) にする"""
    if date_cell is None:
        return None, None
    m = RE_BS_DAY_MONTH.search(date_cell.get_text(" ", strip=True))
    if not m:
        return None, None
    return int(m.group(1)), BS_MONTHS.get(m.group(2).lower())


def besoccer_outcome(date_cell):
    """日付セルの色クラスから、対象選手のチーム視点の勝敗を取り出す"""
    classes = (date_cell.get("class") or []) if date_cell else []
    return next((name for name in ("win", "draw", "lose") if name in classes), "")


def besoccer_score(text):
    """'1-0' / '1(5-4)1' を (ホーム得点, アウェー得点, PK表記) にする"""
    m = RE_BS_SCORE.match(text)
    if m:
        return int(m.group(1)), int(m.group(2)), ""
    m = RE_BS_SCORE_PK.match(text)
    if m:
        home, pk_home, pk_away, away = (int(g) for g in m.groups())
        return home, away, f"(PK{pk_home}-{pk_away})"
    return None, None, ""


def besoccer_minutes(match_cell):
    """出場時間セルを分数にする。ベンチ入りのみは 0、数字が無ければ None"""
    cell = match_cell.find_next_sibling("td") if match_cell else None
    return _to_minutes(cell.get_text(" ", strip=True)) if cell else None


def besoccer_id_year(url):
    """試合URL末尾のID先頭4桁(=開催年)を返す。読めなければ None"""
    m = RE_BS_MATCH_ID.search(urlparse(url).path)
    if not m:
        return None
    year = int(m.group(1))
    return year if 1990 <= year <= date.today().year + 1 else None


def besoccer_parse_rows(html, base_url):
    """試合一覧を新しい順の辞書リストにする(年はまだ入っていない)"""
    soup = BeautifulSoup(html, "lxml")
    matches = []
    for tr in besoccer_match_table(soup).select("tr.row-body"):
        link = tr.select_one("a.match[href]")
        match_cell = tr.select_one("td.match-cell")
        # .image は行頭のリーグロゴと両チームのエンブレム(.shield)の両方に付く
        logo = tr.select_one("td .image:not(.shield) img[src]")
        date_cell = tr.select_one("td.bg-match-res")
        if link is None or match_cell is None or logo is None:
            continue
        day, month = besoccer_day_month(date_cell)
        if day is None or month is None:
            debug(f"日付を読めない行を飛ばします: {tr.get_text(' ', strip=True)[:60]}")
            continue
        shields = [img.get("alt", "") for img in link.select("div.shield img")]
        marker = link.select_one("div.marker-wrapper")
        home_goals, away_goals, pk = besoccer_score(
            marker.get_text(" ", strip=True) if marker else "")
        events_cell = tr.select_one("td.events-cell")
        logo_m = RE_BS_LEAGUE_LOGO.search(logo["src"])
        url = urljoin(base_url, link["href"])
        matches.append({
            "logo": logo_m.group(1) if logo_m else "",
            "league": logo.get("alt", "").strip(),
            "day": day, "month": month, "date": None,
            "url": url, "id_year": besoccer_id_year(url),
            "home": shields[0] if len(shields) > 0 else "",
            "away": shields[1] if len(shields) > 1 else "",
            "home_goals": home_goals, "away_goals": away_goals, "pk": pk,
            "outcome": besoccer_outcome(date_cell),
            "minutes": besoccer_minutes(match_cell),
            "events": RE_BS_EVENT.findall(str(events_cell)) if events_cell else [],
        })
    return matches


def besoccer_fill_years(matches):
    """各行に年を入れ、日付を決められた行だけ返す

    日付セルに年が無いため、まず試合IDの先頭4桁を使う。読めない行は、
    一覧が新しい順に並んでいることを利用して、月が増えたら年を1つ戻す。
    """
    year = next((m["id_year"] for m in matches if m["id_year"]), date.today().year)
    prev_month = None
    for m in matches:
        if m["id_year"]:
            year = m["id_year"]
        elif prev_month is not None and m["month"] > prev_month:
            year -= 1
        try:
            m["date"] = date(year, m["month"], m["day"])
        except ValueError:
            debug(f"日付にできない組み合わせ: {year}-{m['month']}-{m['day']}")
        prev_month = m["month"]
    return [m for m in matches if m["date"]]


def besoccer_side(match, team):
    """対象クラブがホームかアウェーかを返す。判別できなければ None"""
    want = norm(team).casefold()
    names = (("home", norm(match["home"]).casefold()),
             ("away", norm(match["away"]).casefold()))
    if not want:
        return None
    for side, name in names:
        if name == want:
            return side
    for side, name in names:
        if name and (want in name or name in want):
            return side
    return None


def besoccer_result_text(match, side):
    """対象クラブ視点の '勝 3-0' / '分 1-1(PK5-4)' を作る"""
    if match["home_goals"] is None or side is None:
        return ""
    scored, conceded = ((match["home_goals"], match["away_goals"]) if side == "home"
                        else (match["away_goals"], match["home_goals"]))
    outcome = {"win": "勝", "draw": "分", "lose": "敗"}.get(match["outcome"])
    if outcome is None:
        outcome = "勝" if scored > conceded else "敗" if scored < conceded else "分"
    return f"{outcome} {scored}-{conceded}{match['pk']}"


def besoccer_status(match):
    """出場時間セルとイベントアイコンから出場区分と出場時間表示を決める"""
    minutes, events = match["minutes"], match["events"]
    if minutes is None:
        # メンバー表に載っていない。BeSoccerは招集外と未掲載を区別しない
        return OUT_SQUAD, ""
    if minutes <= 0:
        return BENCH, ""
    status = SUBBED_IN if BS_SUB_IN in events else STARTER
    return status, describe_minutes(status, minutes, None)


def besoccer_row(player, team_conf, match, fetched_at=None):
    """BeSoccerの1試合をExcelの1行にする

    fetched_at が入っているときは、この実行で取り直せずキャッシュを使ったという
    ことなので、いつ時点の内容かが分かるように備考へ出す。
    """
    team = str(team_conf.get("team", "")).strip()
    side = besoccer_side(match, team)
    opponent = match["away"] if side == "home" else match["home"] if side == "away" else ""
    status, minutes_text = besoccer_status(match)

    notes = [BESOCCER if fetched_at is None
             else f"{BESOCCER}({fetched_at:%m/%d %H:%M}取得)"]
    if side is None:
        notes.append(f"要確認 チーム名が一致しない({match['home']} / {match['away']})")
        log(f"BeSoccerの対戦カードに '{team}' が見つかりません", "WARN")
    if status == OUT_SQUAD:
        notes.append("出場データなし(招集外とデータ未掲載は区別できない)")
    if BS_INJURY in match["events"]:
        notes.append("負傷")

    row = result(player, status, match=opponent,
                 date=match["date"].strftime("%Y-%m-%d"),
                 url=match["url"], note=" / ".join(notes),
                 match_result=besoccer_result_text(match, side),
                 goals=str(match["events"].count(BS_GOAL)),
                 cards=(f"{match['events'].count(BS_YELLOW)}/"
                        f"{match['events'].count(BS_RED) + match['events'].count(BS_SECOND_YELLOW)}"),
                 league=str(team_conf.get("league", "")).strip() or match["league"])
    row["クラブ"] = str(team_conf.get("club", "")).strip() or player["club"]
    row["出場時間"] = minutes_text
    return row


def bs_short_note(note, limit=120):
    """Excelの備考へ入れる文言を1行に丸める

    Playwrightの例外などは複数行になり、そのまま入れると列が読めなくなる。
    全文は画面のログと cache/besoccer_diagnostics/ に残るので、ここは短くてよい。
    """
    one_line = " ".join(str(note).split())
    return one_line if len(one_line) <= limit else one_line[:limit - 1] + "…"


def besoccer_blank_row(player, team_conf, note):
    """試合を取れなかったチームの行を 試合なし で作る

    黙って行を消すと、設定漏れなのか取得失敗なのか画面から分からなくなる。
    """
    row = blank(player, NO_MATCH, f"{BESOCCER} / {bs_short_note(note)}")
    row["クラブ"] = str(team_conf.get("club", "")).strip() or player["club"]
    row["リーグ名"] = str(team_conf.get("league", "")).strip()
    return row


def besoccer_blank_rows(player, note):
    """設定してある全チーム分の 試合なし 行を作る(取得そのものが失敗したとき)"""
    return [besoccer_blank_row(player, team_conf, note)
            for team_conf in (player.get("besoccer") or {}).get("teams") or []]


def besoccer_rows(player):
    """players.json の besoccer 設定にあるチームごとに、今季の最新試合を1行ずつ作る"""
    conf = player.get("besoccer") or {}
    url = str(conf.get("url", "")).strip()
    teams = list(conf.get("teams") or [])
    if not url or not teams:
        return []

    html, fetched_at = besoccer_fetch(url)
    matches = besoccer_fill_years(besoccer_parse_rows(html, url))
    if not matches:
        raise ScrapeError("BeSoccerの試合一覧を読めない",
                          "--save-html でHTMLを保存し表構造を確認する", url)
    debug(f"BeSoccerの試合一覧 {len(matches)}件 "
          f"({matches[-1]['date']} 〜 {matches[0]['date']})")

    season_start = date(CFG["year"], 1, 1)
    rows = []
    for team_conf in teams:
        logo = str(team_conf.get("logo", "")).strip()
        target = [m for m in matches if m["logo"] == logo and m["date"] >= season_start]
        if not target:
            log(f"BeSoccer: ロゴ {logo} の{CFG['year']}年の試合が0件", "WARN")
            rows.append(besoccer_blank_row(
                player, team_conf,
                f"{CFG['year']}年の試合が0件(リーグロゴ名 {logo} を確認)"))
            continue
        rows.append(besoccer_row(player, team_conf,
                                 max(target, key=lambda m: m["date"]), fetched_at))
    return rows


def with_besoccer_rows(rows, player):
    """BeSoccer側の大会の行を足して返す

    トップチーム側の行は必ず残す。BeSoccerが取れなかった場合も、設定して
    あるチームの行は 試合なし で出し、理由を備考に書く。行が黙って消えると
    設定漏れなのか取得失敗なのか画面から判断できないため。
    """
    if CFG.get("no_besoccer") or player.get("source") == "besoccer" \
            or not player.get("besoccer"):
        return rows
    try:
        extra = besoccer_rows(player)
    except ScrapeError as e:
        log(f"BeSoccerの確認に失敗: {e.summary()}", "WARN")
        extra = besoccer_blank_rows(player, e.summary())
    except Exception as e:
        log(f"BeSoccerの確認に失敗: {type(e).__name__}: {e}", "WARN")
        extra = besoccer_blank_rows(player, f"{type(e).__name__}: {e}")
    for row in extra:
        # 日付が入っている行だけ重複判定する(試合なしの行は日付も試合も空)
        if row["日付"] and any(r["日付"] == row["日付"] and r["対戦相手"] == row["対戦相手"]
                              for r in rows):
            debug(f"BeSoccerの{row['リーグ名']}は既存の行と同じなので追加しない")
            continue
        log(f"BeSoccerの出場記録を追加: {row['リーグ名'] or '大会不明'} "
            f"{row['日付'] or '日付なし'} {row['ステータス']}")
        rows.append(row)
    return rows


def scrape_besoccer(player):
    """source が besoccer の選手。BeSoccerの設定だけで全行を作る"""
    if not player.get("besoccer"):
        return blank(player, NO_MATCH, "besoccer だが besoccer 設定が未記入")
    return besoccer_rows(player)


# =========================================================
# manual (手入力)
# =========================================================
def scrape_manual(player):
    latest = player.get("latest")
    if not latest:
        return blank(player, NO_MATCH, "manual だが latest が未設定")
    row = result(player, latest.get("ステータス", NO_MATCH),
                 match=latest.get("対戦相手", latest.get("試合", "")),
                 date=latest.get("日付", ""),
                 url=player.get("url", ""), note="手入力")
    row["出場時間"] = latest.get("出場時間", "")
    return row


SCRAPERS = {
    "espn": scrape_espn,
    "jleague": scrape_jleague,
    "transfermarkt": scrape_transfermarkt,
    "football-lab": scrape_footballlab,
    "besoccer": scrape_besoccer,
    "manual": scrape_manual,
}


def excel_display_width(value):
    """Excel列幅の概算用。全角文字は2文字分として数える。"""
    width = 0
    for ch in str(value):
        width += 2 if unicodedata.east_asian_width(ch) in ("F", "W", "A") else 1
    return width


def style_sheet(ws):
    """見出し・縞模様・列幅・固定行・フィルタを整える(シート共通)"""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="DDEBF7")
    even_row_fill = PatternFill("solid", fgColor="F7F7F7")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")

    for row in ws.iter_rows(min_row=2):
        if row[0].row % 2 == 0:
            for cell in row:
                cell.fill = even_row_fill
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_width = 0
        for cell in column_cells:
            if cell.value is None:
                continue
            max_width = max(max_width, excel_display_width(cell.value))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_width + 2, 10), 80)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def format_excel(path, team_latest_urls=None):
    """team_latest_urls: {0始まりのデータ行番号: URL} — 「チーム最新試合日付」列の
    日付セルにハイパーリンクを張るための対応表。

    U-21シートを足したブックにも対応するため、全シートに書式を当てる。
    ハイパーリンクはレンタル選手シートにしか無いので、そこだけに張る。
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Font

    wb = load_workbook(path)
    for ws in wb.worksheets:
        style_sheet(ws)

    if team_latest_urls:
        ws = wb[MAIN_SHEET] if MAIN_SHEET in wb.sheetnames else wb.worksheets[0]
        headers = [c.value for c in ws[1]]
        if "チーム最新試合日付" in headers:
            col_idx = headers.index("チーム最新試合日付") + 1
            link_font = Font(color="0563C1", underline="single")
            for row_i, url in team_latest_urls.items():
                cell = ws.cell(row=row_i + 2, column=col_idx)
                if cell.value and url:
                    cell.hyperlink = url
                    cell.font = link_font

    wb.save(path)


# =========================================================
# メイン
# =========================================================
def build_parser():
    ap = argparse.ArgumentParser(
        description=f"レンタル移籍選手の最新試合 出場ステータス一覧 (LoanWatch {APP_VERSION})")
    ap.add_argument("--version", action="version", version=f"LoanWatch {APP_VERSION}")
    ap.add_argument("--year", type=int, default=2026, help="対象シーズン(既定: 2026)")
    ap.add_argument("--sleep", type=float, default=2.0, help="アクセス間隔(秒)")
    ap.add_argument("--timeout", type=int, default=30, help="HTTPタイムアウト(秒)")
    ap.add_argument("--retry", type=int, default=2, help="リトライ回数")
    ap.add_argument("--only", nargs="*", default=None, help="指定した選手名だけ実行")
    ap.add_argument("--source", choices=list(SCRAPERS), default=None,
                    help="指定したsourceだけ実行")
    ap.add_argument("--debug", action="store_true", help="テーブル構造をダンプ")
    ap.add_argument("--headed", action="store_true", help="Playwrightの画面を表示")
    ap.add_argument("--save-html", action="store_true", dest="save_html",
                    help="取得したHTMLをファイル保存")
    ap.add_argument("--no-fallback", action="store_true", dest="no_fallback",
                    help="Football-LAB未更新時にJリーグ公式へ切り替えない")
    ap.add_argument("--no-besoccer", action="store_true", dest="no_besoccer",
                    help="players.json の besoccer 設定を無視し、追加の行を出さない")
    ap.add_argument("--stale-days", type=int, default=14, dest="stale_days",
                    help="この日数以上前の試合は注意表示する(既定: 14)")
    ap.add_argument("--players-config", default=str(DEFAULT_PLAYERS_CONFIG),
                    help="選手設定JSONファイル(既定: players.json)")
    ap.add_argument("--u21-club", nargs="+", default=None, dest="u21_club",
                    metavar="CLUB",
                    help="U-21シートの対象クラブ(IDまたはクラブ名。複数可。"
                         "既定: u21_clubs.json の default)")
    ap.add_argument("--no-u21-team", action="store_true", dest="no_u21_team",
                    help="U-21 Ｊリーグのチーム単位シートを作らない")
    ap.add_argument("--u21-clubs-config", default=str(u21_clubs.default_config_path()),
                    dest="u21_clubs_config",
                    help="U-21クラブ一覧ファイル(既定: u21_clubs.json)")
    ap.add_argument("--out", default="loan_players_latest", help="出力ファイル名(拡張子なし)")
    ap.add_argument("--no-excel", action="store_true", dest="no_excel",
                    help="Excelファイルを出力しない")
    ap.add_argument("--no-csv", action="store_true", dest="no_csv",
                    help="CSVファイルを出力しない")
    ap.add_argument("--playwright-smoke-test", action="store_true", help=argparse.SUPPRESS)
    return ap


def resolve_u21_clubs(args):
    """--u21-club の指定を u21_clubs.json に照らして解決する

    戻り値は (対象クラブのリスト, 解決できなかった指定のリスト)。
    指定が無ければ設定ファイルの default を1件だけ対象にする。
    """
    if args.no_u21_team:
        return [], []
    clubs, default_id = u21_clubs.load_clubs(args.u21_clubs_config)
    resolved, unknown = [], []
    for token in (args.u21_club or [default_id]):
        club = u21_clubs.find_club(clubs, token)
        if club is None:
            unknown.append(str(token))
        elif club not in resolved:
            resolved.append(club)
    return resolved, unknown


def main():
    args = build_parser().parse_args()
    CFG.update(vars(args))

    if args.playwright_smoke_test:
        try:
            ok = playwright_smoke_test()
        except Exception as e:
            print(f"PLAYWRIGHT_SMOKE_TEST_FAILED: {type(e).__name__}: {e}", file=sys.stderr)
            return 1
        print("PLAYWRIGHT_SMOKE_TEST_OK" if ok else "PLAYWRIGHT_SMOKE_TEST_FAILED")
        return 0 if ok else 1

    if args.no_excel and args.no_csv:
        print("出力形式が選択されていません。Excel または CSV を有効にしてください", file=sys.stderr)
        return 1

    try:
        targets = load_players(args.players_config)
    except (OSError, json.JSONDecodeError, ValueError) as e:
        print(f"選手設定ファイルを読み込めません: {e}", file=sys.stderr)
        return 1

    if args.only:
        wanted = {norm(x) for x in args.only}
        targets = [p for p in targets if norm(p["name_ja"]) in wanted]
    if args.source:
        targets = [p for p in targets if p["source"] == args.source]
    if not targets:
        print("対象の選手がいません。--only / --source の指定を確認してください", file=sys.stderr)
        return 1

    _JL_HISTORY_CACHE.clear()
    _JL_CLUB_CACHE.clear()
    _BS_HTML_CACHE.clear()
    _JL_U21_CACHE.clear()
    rows, errors, success = [], [], 0
    for p in targets:
        print(f"■ {p['name_ja']} ({p['club']}) [{p['source']}]")
        scraper = SCRAPERS.get(p["source"])
        if scraper is None:
            log(f"未対応のsource: {p['source']}", "ERROR")
            rows.append(blank(p, NO_MATCH, f"未対応source: {p['source']}"))
            errors.append((p["name_ja"], f"未対応source: {p['source']}"))
            continue
        failed = None
        try:
            # スクレイパは1行(dict)、または大会ごとの複数行(list)を返す
            res = scraper(p)
            player_rows = list(res) if isinstance(res, list) else [res]
            success += 1
        except ScrapeError as e:
            log(e.summary(), "ERROR")
            if e.detail:
                print(f"     {e.detail}")
            failed = blank(p, NO_MATCH, e.summary())
            errors.append((p["name_ja"], e.summary()))
        except Exception as e:
            log(f"想定外のエラー: {type(e).__name__}: {e}", "ERROR")
            if args.debug:
                traceback.print_exc()
            failed = blank(p, NO_MATCH, f"{type(e).__name__}: {e}")
            errors.append((p["name_ja"], f"{type(e).__name__}: {e}"))

        # BeSoccer は別サイトなので、トップチーム側が失敗しても独立して取りに行く
        player_rows = with_besoccer_rows([failed] if failed else player_rows, p)
        for r in player_rows:
            if r is not failed:
                log(f"{r['リーグ名'] or '大会不明'} / {r['ステータス']} "
                    f"{r['出場時間']} ({r['日付']})")
            rows.append(r)

    u21_clubs_wanted, u21_unknown = resolve_u21_clubs(args)
    for token in u21_unknown:
        log(f"U-21の対象クラブが見つかりません: {token} "
            f"(u21_clubs.json の id かクラブ名で指定してください)", "WARN")
    u21_rows, u21_errors = (u21_team_sheet_rows(u21_clubs_wanted)
                            if u21_clubs_wanted else ([], []))

    # どのスクレイパ経由でも 経過 列が欠けないよう、日付から再計算する
    for r in rows:
        r["経過"] = format_elapsed(r.get("日付", ""))
        for c in COLUMNS:
            r.setdefault(c, "")

    # 「チーム最新試合日付」列にハイパーリンクを張るための対応表
    # (行番号は rows と df/Excel のデータ行が同じ並び順である前提)
    team_latest_urls = {i: r["_team_latest_url"] for i, r in enumerate(rows)
                        if r.get("_team_latest_url")}

    df = pd.DataFrame(rows, columns=COLUMNS)
    print("\n" + "=" * 78)
    print(f" 最新試合の出場状況  (集計日: {date.today():%Y-%m-%d} / LoanWatch {APP_VERSION})")
    print("=" * 78)
    print(df.drop(columns=["URL"]).to_string(index=False))

    u21_df = pd.DataFrame(u21_rows, columns=U21_COLUMNS) if u21_rows else None
    if u21_df is not None:
        print("\n" + "=" * 78)
        print(f" {U21} 出場台帳  (シート: {U21_SHEET})")
        print("=" * 78)
        print(u21_df.drop(columns=["URL", "備考"]).to_string(index=False))

    # 日付が古い、あるいは日付が取れなかった選手を明示する
    stale, undated = [], []
    for r in rows:
        n = days_ago(r["日付"])
        if n is None:
            if r["ステータス"] != NO_MATCH:
                undated.append(r["選手"])
        elif n >= args.stale_days:
            stale.append(f"{r['選手']}({r['日付']} / {n}日前)")
    if stale:
        print(f"\n[注意] {args.stale_days}日以上前の試合しか取得できていない選手:")
        for s in stale:
            print(f"  - {s}")
    if undated:
        print(f"\n[注意] 日付を特定できなかった選手: {', '.join(undated)}")

    print(f"\n成功 {success}件 / 失敗 {len(errors)}件")
    for name, msg in errors:
        print(f"  - {name}: {msg}")
    if u21_errors:
        print(f"\n[注意] {U21}のチーム取得に失敗したクラブ:")
        for name, msg in u21_errors:
            print(f"  - {name}: {msg}")

    try:
        saved = []
        if not args.no_excel:
            # U-21は列構成が違うので、同じブックの別シートに分ける
            xlsx_path = f"{args.out}.xlsx"
            with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name=MAIN_SHEET)
                if u21_df is not None:
                    u21_df.to_excel(writer, index=False, sheet_name=U21_SHEET)
            format_excel(xlsx_path, team_latest_urls)
            saved.append(xlsx_path)
        if not args.no_csv:
            # CSVはシートを持てないため、U-21は別ファイルにする
            csv_path = f"{args.out}.csv"
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            saved.append(csv_path)
            if u21_df is not None:
                u21_csv_path = f"{args.out}_u21.csv"
                u21_df.to_csv(u21_csv_path, index=False, encoding="utf-8-sig")
                saved.append(u21_csv_path)
        print(" / ".join(saved) + " に保存しました")
    except Exception as e:
        print(f"ファイル保存に失敗: {type(e).__name__}: {e}", file=sys.stderr)
        return 1

    return 1 if errors or u21_errors else 0


if __name__ == "__main__":
    sys.exit(main())
