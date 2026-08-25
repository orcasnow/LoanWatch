"""LoanWatch - 名古屋グランパス レンタル移籍選手の最新試合 出場ステータス一覧

出力: 選手ごと1行 (先発 / 途中出場 / ベンチ入りのみ / ベンチ外 / 試合なし)
      ただしトップチームのほかに U-21 Ｊリーグ・天皇杯にも今季の記録がある
      選手は、大会ごとに1行ずつ出力する。

準備:
    pip install playwright requests beautifulsoup4 lxml html5lib pandas openpyxl
    playwright install chromium

実行:
    python LoanWatchMain.py
    python LoanWatchMain.py --year 2026 --debug
    python LoanWatchMain.py --headed --save-html          # 画面表示＋HTML保存
    python LoanWatchMain.py --only 貴田遼河 --debug        # 1人だけ試す
    python LoanWatchMain.py --source espn --debug         # ESPN対象の2人だけ

GUI版:
    python LoanWatchGui.py            # 実行設定・ログ画面＋「選手設定を編集」ボタン
    python PlayersEditor.py           # players.json 編集画面を単体起動

出力列:
    日付 / 経過 を先頭寄りに置き、いつの試合かがひと目で分かるようにしている。
    リーグ名 はその試合の大会名(明治安田Ｊ１ / U-21 Ｊリーグ / 天皇杯 / ESPNは英語表記)。
    ベンチ外・ベンチ入りのみは大会名を特定できないため、Jリーグ公式の
    所属リーグで代用し、備考に「リーグ名は所属リーグから推定」と入れる。
    試合 は対戦相手のみ、試合結果 は対象クラブ視点の勝敗・スコア。
    警告/退場 は「警告数/退場数」で表示する。
    --stale-days で指定した日数より古い試合しか取れなかった選手は末尾に列挙する。
"""

import argparse
import asyncio
import os
import re
import sys
import time
import traceback
import json
import unicodedata
from datetime import date, datetime
from io import StringIO
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs


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
           "出場時間", "試合", "試合結果", "ゴール数", "警告/退場", "備考", "URL"]
RE_GAME_URL = re.compile(r"/soccer/match/_/gameId/(\d+)")

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Referer": "https://www.google.com/",
}

# 選手マスタは players.json で管理する。
# source: "espn" | "transfermarkt" | "football-lab" | "manual"
# espn_id は ESPN 選手URLの /id/<数字>/、jleague_url は出場時間補完用。
# espn_team_id を書いておくと、チームの結果ページ
#   (https://www.espn.com/soccer/team/results/_/id/<espn_team_id>)
# と突き合わせ、選手ページの Last 5 Matches がチームの最新試合より
# 古い場合に備考へ「要確認」を出す(選手ページの反映遅れ検知用)。
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
    return players


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
            "試合": match,
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
async def _render(url, headless=True, wait_ms=6000,
                  wait_selector="table.items", locale="ja-JP", scroll=False):
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=headless)
        ctx = await browser.new_context(
            locale=locale, timezone_id="Asia/Tokyo",
            viewport={"width": 1440, "height": 1000},
            extra_http_headers={"Accept-Language": f"{locale},en;q=0.9"})
        page = await ctx.new_page()
        try:
            resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            status = resp.status if resp else 0
            if status == 403:
                raise ScrapeError("HTTP 403 アクセス拒否",
                                  "IPがボット判定された可能性。回線・実行環境を変える", url)
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
    """_render を同期的に呼ぶラッパ"""
    try:
        return asyncio.run(_render(url, headless=not CFG["headed"],
                                   wait_selector=wait_selector,
                                   locale=locale, scroll=scroll))
    except ScrapeError:
        raise
    except Exception as e:
        raise ScrapeError(f"ブラウザ取得に失敗: {type(e).__name__}: {e}",
                          "playwright install chromium を実行済みか確認", url)


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
                r.encoding = r.apparent_encoding
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
    """トップチームのリーグ戦以外(U-21 Ｊリーグ・天皇杯)の行を足す"""
    if CFG.get("no_fallback") or not player.get("jleague_url"):
        return rows
    return with_emperor_row(with_u21_row(rows, player, season_start),
                            player, season_start)


def with_emperor_row(rows, player, season_start):
    """天皇杯の最新試合があれば行を足して返す

    天皇杯はJFA主催のため、Jリーグ公式の出場履歴にも Football-LAB にも載らない。
    クラブページの試合ログから今季の天皇杯を探し、試合ページのメンバー表で
    出場区分を決める。メンバー表を見るのでベンチ外も判別できる。
    """
    try:
        cup = jl_emperor_row(player, season_start)
    except ScrapeError as e:
        log(f"天皇杯の確認に失敗: {e.summary()}", "WARN")
        return rows
    if cup is None:
        return rows
    if any(r["日付"] == cup["日付"] and r["試合"] == cup["試合"] for r in rows):
        debug("天皇杯の試合は1行目と同じなので追加しない")
        return rows
    log(f"天皇杯の出場記録を追加: {cup['日付']} {cup['ステータス']}")
    rows.append(cup)
    return rows


def with_u21_row(rows, player, season_start):
    """U-21 Ｊリーグの最新試合があれば行を足して返す

    トップチームの行は Football-LAB のレポート(=リーグ戦・カップ戦)から作るため、
    U-21 Ｊリーグの出場はJリーグ公式の出場履歴からしか拾えない。
    """
    try:
        df, _ = jl_history_df(player)
    except ScrapeError as e:
        log(f"U-21 Ｊリーグの確認に失敗: {e.summary()}", "WARN")
        return rows
    u21 = jl_u21_row(player, df, season_start)
    if u21 is None:
        return rows
    # Football-LAB が取れずJリーグ公式に落ちた場合、1行目自体が
    # U-21 の試合ということがある。その時は二重に出さない。
    if any(r["日付"] == u21["日付"] and r["試合"] == u21["試合"] for r in rows):
        debug("U-21 Ｊリーグの試合は1行目と同じなので追加しない")
        return rows
    log(f"U-21 Ｊリーグの出場履歴を追加: {u21['日付']} {u21['ステータス']}")
    rows.append(u21)
    return rows


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


def jl_u21_row(player, df, season_start):
    """出場履歴から U-21 Ｊリーグの今季最新試合を1行にする。無ければ None

    U-21 Ｊリーグはトップチームとは別日程で開催されるため、Football-LAB の
    マッチレポートには載らない。両方に出場している選手を1行にまとめると
    どちらかが消えてしまうので、大会ごとに別の行として出す。
    """
    if df.empty:
        return None
    comp_col = find_col(df.columns, [JL_COL_COMP])
    if comp_col is None:
        debug("出場履歴に大会列が無いためU-21の判定をスキップ")
        return None
    u21 = df[df[comp_col].map(lambda v: bool(RE_U21.search(norm(v))))]
    # 出場履歴は過去シーズン分も含むため、今季の記録だけに絞る
    u21 = u21[u21["_date"] >= (season_start or date(CFG["year"], 1, 1))]
    if u21.empty:
        return None

    latest = u21.iloc[-1]
    status, minutes_text = jl_status_from_row(latest, df.columns)
    row = result(player, status, match=jl_value(latest, df.columns, JL_COL_OPP),
                 date=latest["_date"].strftime("%Y-%m-%d"),
                 url=player.get("jleague_url") or player.get("url"),
                 note="Jリーグ公式(U-21 Ｊリーグ) / ベンチ外は判別不可")
    row.update(jl_extra_fields(latest, df.columns))
    row["出場時間"] = minutes_text
    return row


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
RE_JL_LINEUP_SUBS = re.compile(
    r'"m-lineup-list-item","ref":"[^"]*","subs":(\[[^\]]*\]).{0,400}?"memberId":"(\d+)"',
    re.S)
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


def jl_emperor_url(game_id):
    """gameId 2026081917 → https://www.jleague.jp/match/emperor/2026/081917/"""
    return f"https://www.jleague.jp/match/emperor/{game_id[:4]}/{game_id[4:]}/"


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


def jl_emperor_status(player_id, starters, members, subs):
    """メンバー表から、その選手の出場区分と出場時間の表示を決める"""
    events = subs.get(player_id) or []

    def minute(kind):
        for e in events:
            if e.get("type") == kind:
                m = re.search(r"(\d+)", str(e.get("time", "")))
                return int(m.group(1)) if m else None
        return None

    if player_id in starters:
        out = minute("out")
        return STARTER, (f"先発〜{out}分で交代({out}分)" if out else "フル出場")
    if player_id in members:
        entry = minute("in")
        if entry is None:
            return BENCH, ""
        minutes = max(90 - entry, 0)
        _, label = minutes_to_entry(minutes)
        return SUBBED_IN, (f"{label}から出場(約{minutes}分)" if label else f"約{minutes}分")
    return OUT_SQUAD, ""


def jl_emperor_row(player, season_start):
    """今季の天皇杯の最新試合から、その選手の出場状況を1行にする。無ければ None"""
    m = RE_JL_PLAYER_ID.search(player.get("jleague_url") or "")
    slug = jl_page_value(player, "slug")
    if not m or not slug:
        debug("選手IDまたはクラブslugが取れないため天皇杯をスキップ")
        return None
    player_id = m.group(1)

    # クラブページの試合ログは過去シーズン分も含むので、今季開幕以降に絞る
    floor = season_start or date(CFG["year"], 1, 1)
    cup = [g for g in jl_club_games(slug, player["name_ja"])
           if EMPEROR in g["league"] and g["date"] and g["date"] >= floor and g["game_id"]]
    if not cup:
        debug(f"{CFG['year']}シーズンの天皇杯の試合は無し")
        return None

    latest = max(cup, key=lambda g: g["date"])
    match_url = jl_emperor_url(latest["game_id"])
    html = fetch(match_url)
    save_html(html, f"{player['name_ja']}_emperor")
    starters, members, subs, events = jl_match_members(html)
    if not members:
        raise ScrapeError("天皇杯の試合ページからメンバー表を取得できない",
                          "--save-html でHTMLを保存し構造を確認する", match_url)

    status, minutes_text = jl_emperor_status(player_id, starters, members, subs)
    played = status in (STARTER, SUBBED_IN)
    ev = events.get(player_id) or {}
    row = result(player, status, match=latest["opponent"],
                 date=latest["date"].strftime("%Y-%m-%d"),
                 url=match_url, league=EMPEROR, match_result=latest["result"],
                 goals=str(ev.get("goal", 0)) if played else "",
                 cards=f"{ev.get('yellow-card', 0)}/{ev.get('red-card', 0)}" if played else "",
                 note="Jリーグ公式(天皇杯)")
    row["出場時間"] = minutes_text
    return row


def minutes_to_entry(minutes):
    """出場時間から出場時刻を逆算し、'後半15分' のような表記にする"""
    if not minutes or minutes >= 90:
        return None, None
    entry = 90 - minutes
    if entry >= 45:
        return entry, f"後半{entry - 45}分"
    return entry, f"前半{entry}分"


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
# manual (手入力)
# =========================================================
def scrape_manual(player):
    latest = player.get("latest")
    if not latest:
        return blank(player, NO_MATCH, "manual だが latest が未設定")
    row = result(player, latest.get("ステータス", NO_MATCH),
                 match=latest.get("試合", ""), date=latest.get("日付", ""),
                 url=player.get("url", ""), note="手入力")
    row["出場時間"] = latest.get("出場時間", "")
    return row


SCRAPERS = {
    "espn": scrape_espn,
    "jleague": scrape_jleague,
    "transfermarkt": scrape_transfermarkt,
    "football-lab": scrape_footballlab,
    "manual": scrape_manual,
}


def excel_display_width(value):
    """Excel列幅の概算用。全角文字は2文字分として数える。"""
    width = 0
    for ch in str(value):
        width += 2 if unicodedata.east_asian_width(ch) in ("F", "W", "A") else 1
    return width


def format_excel(path, team_latest_urls=None):
    """team_latest_urls: {0始まりのデータ行番号: URL} — 「チーム最新試合日付」列の
    日付セルにハイパーリンクを張るための対応表。"""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = load_workbook(path)
    ws = wb.active

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

    if team_latest_urls:
        headers = [c.value for c in ws[1]]
        if "チーム最新試合日付" in headers:
            col_idx = headers.index("チーム最新試合日付") + 1
            link_font = Font(color="0563C1", underline="single")
            for row_i, url in team_latest_urls.items():
                cell = ws.cell(row=row_i + 2, column=col_idx)
                if cell.value and url:
                    cell.hyperlink = url
                    cell.font = link_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


# =========================================================
# メイン
# =========================================================
def build_parser():
    ap = argparse.ArgumentParser(description="レンタル移籍選手の最新試合 出場ステータス一覧")
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
    ap.add_argument("--stale-days", type=int, default=14, dest="stale_days",
                    help="この日数以上前の試合は注意表示する(既定: 14)")
    ap.add_argument("--players-config", default=str(DEFAULT_PLAYERS_CONFIG),
                    help="選手設定JSONファイル(既定: players.json)")
    ap.add_argument("--out", default="loan_players_latest", help="出力ファイル名(拡張子なし)")
    ap.add_argument("--no-excel", action="store_true", dest="no_excel",
                    help="Excelファイルを出力しない")
    ap.add_argument("--no-csv", action="store_true", dest="no_csv",
                    help="CSVファイルを出力しない")
    ap.add_argument("--playwright-smoke-test", action="store_true", help=argparse.SUPPRESS)
    return ap


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
    rows, errors, success = [], [], 0
    for p in targets:
        print(f"■ {p['name_ja']} ({p['club']}) [{p['source']}]")
        scraper = SCRAPERS.get(p["source"])
        if scraper is None:
            log(f"未対応のsource: {p['source']}", "ERROR")
            rows.append(blank(p, NO_MATCH, f"未対応source: {p['source']}"))
            errors.append((p["name_ja"], f"未対応source: {p['source']}"))
            continue
        try:
            # スクレイパは1行(dict)、または大会ごとの複数行(list)を返す
            res = scraper(p)
            for r in (res if isinstance(res, list) else [res]):
                log(f"{r['リーグ名'] or '大会不明'} / {r['ステータス']} "
                    f"{r['出場時間']} ({r['日付']})")
                rows.append(r)
            success += 1
        except ScrapeError as e:
            log(e.summary(), "ERROR")
            if e.detail:
                print(f"     {e.detail}")
            rows.append(blank(p, NO_MATCH, e.summary()))
            errors.append((p["name_ja"], e.summary()))
        except Exception as e:
            log(f"想定外のエラー: {type(e).__name__}: {e}", "ERROR")
            if args.debug:
                traceback.print_exc()
            rows.append(blank(p, NO_MATCH, f"{type(e).__name__}: {e}"))
            errors.append((p["name_ja"], f"{type(e).__name__}: {e}"))

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
    print(f" 最新試合の出場状況  (集計日: {date.today():%Y-%m-%d})")
    print("=" * 78)
    print(df.drop(columns=["URL"]).to_string(index=False))

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

    try:
        saved = []
        if not args.no_excel:
            xlsx_path = f"{args.out}.xlsx"
            df.to_excel(xlsx_path, index=False)
            format_excel(xlsx_path, team_latest_urls)
            saved.append(xlsx_path)
        if not args.no_csv:
            csv_path = f"{args.out}.csv"
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            saved.append(csv_path)
        print(" / ".join(saved) + " に保存しました")
    except Exception as e:
        print(f"ファイル保存に失敗: {type(e).__name__}: {e}", file=sys.stderr)
        return 1

    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
