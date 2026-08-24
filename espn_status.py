"""ESPN の選手ページから最新試合の出場状況を取得する

処理の流れ:
    1. 選手ページの Last 5 Matches の最上位行から Result 列のリンク(試合ページ)を取る
    2. 試合ページの Formations & Lineups から対象選手の項目を探す
    3. 先発 / 途中出場 / ベンチ入りのみ / ベンチ外 と出場時間を判定する

ESPN は React で描画されるため Playwright を使う。
ページ内の window['__espnfitt__'] に元データのJSONが埋まっているので、
DOM解析とJSON解析の二本立てで、どちらかが通れば結果を出す。

準備:
    pip install playwright beautifulsoup4 lxml
    playwright install chromium

実行:
    python espn_status.py                     # 貴田遼河の最新試合
    python espn_status.py --dump              # 構造調査用。JSONとHTMLを保存
    python espn_status.py --debug --headed
"""

import argparse
import asyncio
import json
import re
import sys
from datetime import date
from urllib.parse import urljoin

from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# =========================================================
# 対象選手
# =========================================================
TARGET = {
    "name_ja": "貴田遼河",
    "espn_id": "345269",
    "jersey": "34",
    "name_frags": ["kida", "貴田"],   # 小文字で照合する
    "url": "https://www.espn.com/soccer/player/_/id/345269/ryoga-kida",
}

STARTER, SUBBED_IN, BENCH, OUT_SQUAD, NO_MATCH = \
    "先発", "途中出場", "ベンチ入りのみ", "ベンチ外", "試合なし"

RE_GAME_URL = re.compile(r"/soccer/match/_/gameId/(\d+)")
RE_MINUTE = re.compile(r"(\d{1,3})\s*'")

CFG = {}


class ScrapeError(Exception):
    def __init__(self, message, hint="", detail=""):
        super().__init__(message)
        self.message, self.hint, self.detail = message, hint, detail

    def summary(self):
        return self.message + (f" / 対処: {self.hint}" if self.hint else "")


def log(msg, level="INFO"):
    print(f"  [{level}] {msg}")


def debug(msg):
    if CFG.get("debug"):
        log(msg, "DEBUG")


# =========================================================
# 取得
# =========================================================
async def _render(url, wait_selector=None, headless=True, wait_ms=8000):
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=headless)
        ctx = await browser.new_context(
            locale="en-US", timezone_id="Asia/Tokyo",
            viewport={"width": 1440, "height": 1000})
        page = await ctx.new_page()
        try:
            resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            status = resp.status if resp else 0
            if status >= 400:
                raise ScrapeError(f"HTTP {status}", "URLと選手IDを確認", url)

            for sel in ("#onetrust-accept-btn-handler",
                        "button:has-text('I Accept')",
                        "button:has-text('Accept')"):
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
            # 遅延読み込み対策で一度スクロールしてから待つ
            await page.mouse.wheel(0, 3000)
            await page.wait_for_timeout(2000)
            return await page.content()
        finally:
            await browser.close()


def render(url, wait_selector=None):
    try:
        return asyncio.run(_render(url, wait_selector, headless=not CFG["headed"]))
    except ScrapeError:
        raise
    except Exception as e:
        raise ScrapeError(f"ブラウザ取得に失敗: {type(e).__name__}: {e}",
                          "playwright install chromium を実行済みか確認", url)


def save(text, name, ext="html"):
    if not CFG.get("dump"):
        return
    fn = f"{re.sub(r'\\W+', '_', name)[:60]}.{ext}"
    with open(fn, "w", encoding="utf-8") as f:
        f.write(text)
    log(f"保存: {fn}")


# =========================================================
# 埋め込みJSON (window['__espnfitt__'])
# =========================================================
def extract_fitt(html):
    """ページに埋まっている ESPN の状態JSONを取り出す。無ければ None"""
    m = re.search(r"window\[.__espnfitt__.\]\s*=\s*(\{.*?\});?\s*</script>",
                  html, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError as e:
        debug(f"__espnfitt__ のJSON解析に失敗: {e}")
        return None


def walk(node, path=""):
    """JSONを再帰的に走査して (パス, 値) を全て返す"""
    if isinstance(node, dict):
        for k, v in node.items():
            yield from walk(v, f"{path}.{k}")
    elif isinstance(node, list):
        for i, v in enumerate(node):
            yield from walk(v, f"{path}[{i}]")
    else:
        yield path, node


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


# =========================================================
# 1. Last 5 Matches → 最上位行の試合URL
# =========================================================
def parse_row_date(text):
    """'Thu 3/12' のような表記を date にする。年が無いので当年基準で推定する"""
    m = re.search(r"\b(\d{1,2})/(\d{1,2})\b", text)
    if not m:
        return None
    today = date.today()
    mo, d = int(m.group(1)), int(m.group(2))
    try:
        guess = date(today.year, mo, d)
    except ValueError:
        return None
    # 未来の日付になったら前年の試合とみなす
    return date(today.year - 1, mo, d) if guess > today else guess


def latest_match_from_dom(html, base_url):
    """Last 5 Matches の全行を集め、日付が最も新しい行の試合リンクを返す

    ESPN の表は昇順のことがあるため、最上位行を無条件に採用しない。
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
            continue   # ヘッダ行
        texts = [c.get_text(" ", strip=True) for c in cells]
        row_text = " | ".join(texts)
        url = urljoin(base_url, link["href"])
        # APP列(Started / Sub)があれば拾っておく。試合ページの判定と突き合わせる
        app = next((t for t in texts
                    if re.fullmatch(r"(Started|Sub|Subbed On|Bench|-)", t, re.I)), "")
        rows.append({"row": row_text, "url": url,
                     "game_id": RE_GAME_URL.search(url).group(1),
                     "result": link.get_text(" ", strip=True),
                     "app": app, "date": parse_row_date(row_text)})

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


def latest_match_from_json(html, base_url):
    """DOMで取れない場合に、埋め込みJSONから試合リンクを拾う"""
    fitt = extract_fitt(html)
    if fitt is None:
        return None
    ids = []
    for path, value in walk(fitt):
        if isinstance(value, str):
            m = RE_GAME_URL.search(value)
            if m and m.group(1) not in [i[0] for i in ids]:
                ids.append((m.group(1), path, value))
    if not ids:
        return None
    debug(f"JSONから見つかった試合ID: {[i[0] for i in ids][:8]}")
    game_id, path, value = ids[0]     # 出現順の先頭を最新とみなす
    return {"row": f"(JSON {path})", "game_id": game_id, "result": "",
            "url": urljoin(base_url, value if value.startswith("/") else
                           f"/soccer/match/_/gameId/{game_id}")}


def get_latest_match(player_url):
    html = render(player_url, wait_selector="table")
    save(html, "espn_player")
    info = latest_match_from_dom(html, player_url) \
        or latest_match_from_json(html, player_url)
    if info is None:
        raise ScrapeError("Last 5 Matches から試合リンクを取得できない",
                          "--dump でHTMLを保存し、表の構造を確認する", player_url)
    return info


# =========================================================
# 2. Formations & Lineups → 選手の出場状況
# =========================================================
def matches_target(text):
    t = str(text).lower()
    return any(f in t for f in TARGET["name_frags"])


def deep_find_minute(node, depth=0):
    """選手ノードの中から '85' や "85'" のような分表記を探す"""
    MIN_KEYS = ("clock", "clck", "time", "min", "dsplyvl", "displayvalue", "value", "txt")
    if depth > 6:
        return None
    if isinstance(node, dict):
        for k, v in node.items():
            if any(mk in k.lower() for mk in MIN_KEYS) and isinstance(v, (str, int)):
                m = re.match(r"^(\d{1,3})\s*'?$", str(v).strip())
                if m and 0 < int(m.group(1)) <= 130:
                    return int(m.group(1))
        for v in node.values():
            got = deep_find_minute(v, depth + 1)
            if got is not None:
                return got
    elif isinstance(node, list):
        for v in node:
            got = deep_find_minute(v, depth + 1)
            if got is not None:
                return got
    return None


def player_display_name(obj):
    if isinstance(obj, dict):
        for k in ("shrtNm", "name", "displayName", "lstName"):
            if obj.get(k):
                return str(obj[k])
    return str(obj) if obj else ""


def find_player_record(fitt):
    """ESPN実構造 page.content.gamepackage.lineUps[*].playersMap[<id>] を探す"""
    pid = TARGET["espn_id"]
    try:
        lineups = fitt["page"]["content"]["gamepackage"]["lineUps"]
        for i, lu in enumerate(lineups):
            pmap = lu.get("playersMap") or {}
            if pid in pmap:
                return f".page.content.gamepackage.lineUps[{i}].playersMap.{pid}", pmap[pid], lu
    except (KeyError, TypeError, IndexError):
        debug("既知のパスに playersMap が無い。再帰探索に切り替える")

    # 構造が変わった場合の保険
    hits = find_dicts(fitt, lambda d: str(d.get("id", "")) == pid or
                      (isinstance(d.get("playersMap"), dict) and pid in d["playersMap"]))
    for path, d in hits:
        if isinstance(d.get("playersMap"), dict):
            return f"{path}.playersMap.{pid}", d["playersMap"][pid], d
        return path, d, None
    return None, None, None


def status_from_json(html):
    """試合ページのJSONから出場状況を判定する

    ESPNの選手レコードには starter / subbedIn が無い。実際にあるのは:
      nmbr            背番号
      frmtnPlc        フォーメーション上の位置(先発は1〜11)
      subbedOutPlayer 自分と入れ替わってピッチを出た選手(=自分は途中出場)
      plyrPlys        出来事の一覧(交代の分数がここに入る)
    """
    fitt = extract_fitt(html)
    if fitt is None:
        return None
    save(json.dumps(fitt, ensure_ascii=False, indent=1), "espn_match_fitt", "json")

    path, rec, lineup = find_player_record(fitt)
    if rec is None:
        debug("JSONに選手レコードが無い(メンバー入りしていない可能性)")
        return {"status": OUT_SQUAD, "minutes": None, "entry": None,
                "subbed_in": None, "subbed_out": None, "via": "JSON(レコード無し)"}

    debug(f"採用: {path} / keys={list(rec)}")
    if CFG.get("debug"):
        for k in ("nmbr", "frmtnPlc", "subbedOutPlayer", "plyrPlys", "stats"):
            if k in rec:
                v = json.dumps(rec[k], ensure_ascii=False)
                print(f"     {k}: {v[:300]}")

    swapped = rec.get("subbedOutPlayer")
    frmtn = rec.get("frmtnPlc")
    try:
        frmtn_num = int(str(frmtn))
    except (TypeError, ValueError):
        frmtn_num = 0

    entry = deep_find_minute(rec.get("plyrPlys")) or deep_find_minute(rec)

    if swapped:
        # 誰かと入れ替わって入った = 途中出場
        status = SUBBED_IN
        minutes = (90 - entry) if entry else None
    elif 1 <= frmtn_num <= 11:
        status = STARTER
        minutes = entry if entry else 90
    else:
        status = BENCH
        minutes = None

    return {"status": status, "minutes": minutes, "entry": entry,
            "subbed_in": player_display_name(swapped) if swapped else None,
            "subbed_out": None,
            "via": f"JSON {path} (frmtnPlc={frmtn}, subbedOutPlayer="
                   f"{'有' if swapped else '無'})"}


def status_from_dom(html):
    """Formations & Lineups セクションのDOMから出場状況を判定する"""
    soup = BeautifulSoup(html, "lxml")

    section = None
    for tag in soup.find_all(string=re.compile(r"Formations?\s*&\s*Lineups?|Lineups?", re.I)):
        cand = tag.find_parent(lambda t: t.name in ("section", "div"))
        if cand and cand.find(string=re.compile("|".join(TARGET["name_frags"]), re.I)):
            section = cand
            break
    if section is None:
        section = soup
        debug("Formations & Lineups セクションを特定できずページ全体から探す")

    node = section.find(string=re.compile("|".join(TARGET["name_frags"]), re.I))
    if node is None:
        return {"status": OUT_SQUAD, "minutes": None,
                "subbed_in": None, "subbed_out": None,
                "via": "DOM(名前が見つからない=ベンチ外)"}

    # 選手名を含む最小のブロックだけを見る。
    # 親を辿りすぎると Starters と Substitutes の両方を飲み込んで判定を誤る。
    block = node.find_parent(["li", "tr"]) or node.parent
    text = block.get_text(" ", strip=True)
    debug(f"DOM該当ブロック: {text[:120]}")

    # 選手名の直前にある見出しで、先発か控えかを判定する
    heading = ""
    for prev in node.find_all_previous(string=True, limit=300):
        s = str(prev).strip()
        if re.fullmatch(r"(Starters?|Starting\s*XI|Substitutes?|Subs?|Bench|Reserves)",
                        s, re.I):
            heading = s
            break
    debug(f"直前の見出し: {heading or '(なし)'}")

    minutes = None
    m = RE_MINUTE.search(text)
    if m:
        minutes = m.group(1)

    if re.search(r"Substitut|Bench|Reserve", heading, re.I):
        status = SUBBED_IN if minutes else BENCH
    elif heading:
        status = STARTER
    else:
        # 見出しが取れない場合は交代表示の有無で判断する
        status = SUBBED_IN if re.search(r"(sub|in\b|▲|↑)", text, re.I) and minutes \
            else (STARTER if minutes else BENCH)

    return {"status": status, "minutes": minutes,
            "subbed_in": None, "subbed_out": None,
            "via": f"DOM(見出し: {heading or '不明'})"}


def get_status(match_url):
    html = render(match_url, wait_selector="body")
    save(html, "espn_match")
    if CFG.get("dump"):
        soup = BeautifulSoup(html, "lxml")
        hits = soup.find_all(string=re.compile("|".join(TARGET["name_frags"]), re.I))
        log(f"ページ内の選手名ヒット {len(hits)}件")
        for h in hits[:5]:
            parent = h.find_parent(["li", "tr", "div"]) or h.parent
            print(f"     {parent.get_text(' ', strip=True)[:120]}")

    res = status_from_json(html)
    if res is None:
        debug("JSONで判定できなかったのでDOMに切り替える")
        res = status_from_dom(html)
    return res


# =========================================================
# メイン
# =========================================================
def inspect_json(filename):
    """保存済みの espn_match_fitt.json から対象選手のノードを表示する"""
    with open(filename, encoding="utf-8") as f:
        fitt = json.load(f)
    path, rec, lineup = find_player_record(fitt)
    if rec is None:
        print(f"選手ID {TARGET['espn_id']} のレコードが見つかりません")
        return 1
    print(f"パス: {path}\n")
    print(json.dumps(rec, ensure_ascii=False, indent=2))
    if lineup:
        print("\n--- 同じlineUps要素のキー ---")
        print(list(lineup))
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", default=TARGET["url"], help="ESPNの選手ページURL")
    ap.add_argument("--inspect", metavar="JSON",
                    help="保存済みJSONから選手ノードを表示して終了")
    ap.add_argument("--debug", action="store_true")
    ap.add_argument("--headed", action="store_true", help="ブラウザ画面を表示")
    ap.add_argument("--dump", action="store_true", help="HTML/JSONを保存し診断情報を出す")
    CFG.update(vars(ap.parse_args()))

    if CFG.get("inspect"):
        return inspect_json(CFG["inspect"])

    print(f"■ {TARGET['name_ja']} (ESPN id={TARGET['espn_id']})")
    try:
        match = get_latest_match(CFG["url"])
    except ScrapeError as e:
        log(e.summary(), "ERROR")
        if e.detail:
            print(f"     {e.detail}")
        print(f"\n結果: {NO_MATCH}")
        return 1

    log(f"最新試合: {match['result'] or '(結果不明)'} / gameId={match['game_id']}")
    log(f"試合ページ: {match['url']}")

    try:
        res = get_status(match["url"])
    except ScrapeError as e:
        log(e.summary(), "ERROR")
        print(f"\n結果: {NO_MATCH}")
        return 1

    print()
    print(f"選手      : {TARGET['name_ja']} (#{TARGET['jersey']})")
    print(f"試合      : {match['row']}")
    print(f"試合ページ: {match['url']}")
    print(f"ステータス: {res['status']}")
    if res.get("entry"):
        print(f"出場開始  : {res['entry']}分")
    print(f"出場時間  : {f'約{res["minutes"]}分' if res.get('minutes') else '-'}")
    if res.get("subbed_in"):
        print(f"交代相手  : {res['subbed_in']} と交代して出場")
    print(f"判定経路  : {res['via']}")

    # Last 5 Matches の APP列と食い違っていたら知らせる
    app = (match.get("app") or "").lower()
    expected = {"sub": SUBBED_IN, "started": STARTER}.get(app)
    if expected and expected != res["status"]:
        log(f"一覧のAPP列は '{match['app']}' ({expected}) で、判定 "
            f"'{res['status']}' と食い違っています", "WARN")
    return 0


if __name__ == "__main__":
    sys.exit(main())