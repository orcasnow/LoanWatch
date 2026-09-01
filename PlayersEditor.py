"""players.json editor for LoanWatch.

起動:
    python PlayersEditor.py
    python PlayersEditor.py path/to/players.json
"""

import argparse
import json
import os
import re
import shutil
import sys
import threading
import unicodedata
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk
from urllib.parse import quote, urljoin

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

import requests
import urllib3
from bs4 import BeautifulSoup


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def configure_playwright_browsers():
    if getattr(sys, "frozen", False):
        bundle_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(bundle_dir / "ms-playwright"))


configure_playwright_browsers()


DEFAULT_PLAYERS_CONFIG = app_dir() / "players.json"

SOURCE_DISPLAY_TO_VALUE = {
    "Jリーグ": "football-lab",
    "海外リーグ": "espn",
}
SOURCE_VALUE_TO_DISPLAY = {value: label for label, value in SOURCE_DISPLAY_TO_VALUE.items()}
SOURCE_DISPLAY_CHOICES = tuple(SOURCE_DISPLAY_TO_VALUE)

TEAM_URLS = {
    "鹿島アントラーズ": "https://www.football-lab.jp/kasm/report",
    "水戸ホーリーホック": "https://www.football-lab.jp/mito/report",
    "浦和レッズ": "https://www.football-lab.jp/uraw/report",
    "ジェフユナイテッド千葉": "https://www.football-lab.jp/chib/report",
    "柏レイソル": "https://www.football-lab.jp/kasw/report",
    "FC東京": "https://www.football-lab.jp/fctk/report",
    "東京ヴェルディ": "https://www.football-lab.jp/tk-v/report",
    "FC町田ゼルビア": "https://www.football-lab.jp/mcd/report",
    "川崎フロンターレ": "https://www.football-lab.jp/ka-f/report",
    "横浜F・マリノス": "https://www.football-lab.jp/y-fm/report",
    "清水エスパルス": "https://www.football-lab.jp/shim/report",
    "名古屋グランパス": "https://www.football-lab.jp/nago/report",
    "京都サンガF.C.": "https://www.football-lab.jp/kyot/report",
    "ガンバ大阪": "https://www.football-lab.jp/g-os/report",
    "セレッソ大阪": "https://www.football-lab.jp/c-os/report",
    "ヴィッセル神戸": "https://www.football-lab.jp/kobe/report",
    "ファジアーノ岡山": "https://www.football-lab.jp/okay/report",
    "サンフレッチェ広島": "https://www.football-lab.jp/hiro/report",
    "アビスパ福岡": "https://www.football-lab.jp/fuku/report",
    "V・ファーレン長崎": "https://www.football-lab.jp/ngsk/report",
    "北海道コンサドーレ札幌": "https://www.football-lab.jp/sapp/report",
    "ヴァンラーレ八戸": "https://www.football-lab.jp/hach/report",
    "ベガルタ仙台": "https://www.football-lab.jp/send/report",
    "ブラウブリッツ秋田": "https://www.football-lab.jp/aki/report",
    "モンテディオ山形": "https://www.football-lab.jp/yama/report",
    "いわきFC": "https://www.football-lab.jp/ifc/report",
    "栃木シティ": "https://www.football-lab.jp/to-c/report",
    "RB大宮アルディージャ": "https://www.football-lab.jp/omiy/report",
    "横浜FC": "https://www.football-lab.jp/y-fc/report",
    "湘南ベルマーレ": "https://www.football-lab.jp/shon/report",
    "ヴァンフォーレ甲府": "https://www.football-lab.jp/kofu/report",
    "アルビレックス新潟": "https://www.football-lab.jp/niig/report",
    "カターレ富山": "https://www.football-lab.jp/toya/report",
    "ジュビロ磐田": "https://www.football-lab.jp/iwat/report",
    "藤枝MYFC": "https://www.football-lab.jp/fuji/report",
    "徳島ヴォルティス": "https://www.football-lab.jp/toku/report",
    "FC今治": "https://www.football-lab.jp/imab/report",
    "サガン鳥栖": "https://www.football-lab.jp/tosu/report",
    "大分トリニータ": "https://www.football-lab.jp/oita/report",
    "テゲバジャーロ宮崎": "https://www.football-lab.jp/myzk/report",
    "福島ユナイテッドFC": "https://www.football-lab.jp/fksm/report",
    "栃木SC": "https://www.football-lab.jp/to-s/report",
    "ザスパ群馬": "https://www.football-lab.jp/gnm/report",
    "SC相模原": "https://www.football-lab.jp/sagm/report",
    "松本山雅FC": "https://www.football-lab.jp/mats/report",
    "AC長野パルセイロ": "https://www.football-lab.jp/naga/report",
    "ツエーゲン金沢": "https://www.football-lab.jp/kana/report",
    "FC岐阜": "https://www.football-lab.jp/gifu/report",
    "レイラック滋賀FC": "https://www.football-lab.jp/rsfc/report",
    "FC大阪": "https://www.football-lab.jp/f-os/report",
    "奈良クラブ": "https://www.football-lab.jp/nara/report",
    "ガイナーレ鳥取": "https://www.football-lab.jp/totr/report",
    "レノファ山口FC": "https://www.football-lab.jp/r-ya/report",
    "カマタマーレ讃岐": "https://www.football-lab.jp/sanu/report",
    "愛媛FC": "https://www.football-lab.jp/ehim/report",
    "高知ユナイテッドSC": "https://www.football-lab.jp/kusc/report",
    "ギラヴァンツ北九州": "https://www.football-lab.jp/kiky/report",
    "ロアッソ熊本": "https://www.football-lab.jp/kuma/report",
    "鹿児島ユナイテッドFC": "https://www.football-lab.jp/kufc/report",
    "FC琉球": "https://www.football-lab.jp/ryuk/report",
}

TEAM_CHOICES = tuple(TEAM_URLS)

TABLE_COLUMNS = (
    "name_ja",
    "name_en",
    "club",
    "source",
    "espn_id",
    "espn_team",
    "espn_team_id",
    "jleague_url",
    "url",
    "besoccer",
)

HEADINGS = {
    "name_ja": "選手名",
    "name_en": "英語名",
    "club": "クラブ",
    "source": "source",
    "espn_id": "ESPN ID",
    "espn_team": "ESPNチーム",
    "espn_team_id": "ESPNチームID",
    "jleague_url": "JリーグURL",
    "url": "URL",
    "besoccer": "BeSoccer追加チーム",
}

# BeSoccer(リザーブチームなどの追加行)の入力欄。players.json では
#   "besoccer": {"url": ..., "teams": [{"logo","team","club","league"}, ...]}
# という入れ子になるので、画面用に平らな bs_* として持ち、保存時に組み立てる。
# 画面で編集できるのは teams の1件目だけで、2件目以降は元のJSONのまま残す。
BESOCCER_FORM_FIELDS = ("bs_url", "bs_logo", "bs_team", "bs_club", "bs_league")
BESOCCER_REQUIRED = (("bs_url", "BeSoccer選手URL"),
                     ("bs_logo", "リーグロゴ名"),
                     ("bs_team", "BeSoccerチーム名"))

KNOWN_FORM_FIELDS = {
    "name_ja",
    "name_en",
    "club",
    "source",
    "url",
    "espn_id",
    "espn_team",
    "espn_team_id",
    "jleague_url",
    *BESOCCER_FORM_FIELDS,
}

OUTPUT_FIELDS_BY_SOURCE = {
    "espn": (
        "name_ja",
        "name_en",
        "club",
        "source",
        "espn_id",
        "espn_team",
        "espn_team_id",
        "url",
    ),
    "football-lab": (
        "name_ja",
        "name_en",
        "club",
        "source",
        "url",
        "jleague_url",
    ),
}

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    )
}

JLEAGUE_BASE = "https://www.jleague.jp/"
JLEAGUE_CLUB_LIST_URLS = (
    "https://www.jleague.jp/j1/club/",
    "https://www.jleague.jp/j2/club/",
    "https://www.jleague.jp/j3/club/",
)
_JLEAGUE_CLUB_PLAYER_URLS = None


def load_players_config(path):
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, dict):
        players = data.get("players")
        if players is None:
            raise ValueError("JSONに players 配列がありません")
        if not isinstance(players, list):
            raise ValueError("players は配列である必要があります")
        if not all(isinstance(player, dict) for player in players):
            raise ValueError("players の各要素はオブジェクトである必要があります")
        return data, [dict(player) for player in players]

    if isinstance(data, list):
        if not all(isinstance(player, dict) for player in data):
            raise ValueError("配列の各要素はオブジェクトである必要があります")
        return data, [dict(player) for player in data]

    raise ValueError("JSONのトップレベルはオブジェクトまたは配列である必要があります")


def besoccer_summary(player):
    """一覧表示用に besoccer 設定を 'ロゴ名: チーム名' の並びにする"""
    teams = (player.get("besoccer") or {}).get("teams") or []
    return " / ".join(f"{team.get('logo', '')}: {team.get('team', '')}"
                      for team in teams if isinstance(team, dict))


def build_output_data(original_data, players):
    if isinstance(original_data, dict):
        data = dict(original_data)
        data["players"] = players
        return data
    return players


def backup_path_for(path):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_{timestamp}.backup{path.suffix}")


def norm_text(value):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value)))


def request_text(url):
    try:
        response = requests.get(url, timeout=15, headers=REQUEST_HEADERS)
    except requests.exceptions.SSLError:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response = requests.get(url, timeout=15, headers=REQUEST_HEADERS, verify=False)
    response.raise_for_status()
    return response.text


def request_json(url):
    try:
        response = requests.get(url, timeout=15, headers=REQUEST_HEADERS)
    except requests.exceptions.SSLError:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response = requests.get(url, timeout=15, headers=REQUEST_HEADERS, verify=False)
    response.raise_for_status()
    return response.json()


def espn_team_from_player_detail(espn_id, league_slug):
    if not espn_id or not league_slug:
        return ""

    url = (
        "https://site.web.api.espn.com/apis/common/v3/sports/soccer/"
        f"{quote(league_slug)}/athletes/{espn_id}"
    )
    try:
        data = request_json(url)
    except Exception:
        return ""

    team = ((data.get("athlete") or {}).get("team") or {})
    for key in ("displayName", "shortDisplayName", "name", "location"):
        value = str(team.get(key, "")).strip()
        if value:
            return value
    return ""


def repair_jleague_text(value):
    text = str(value)
    try:
        text = text.encode("utf-8").decode("unicode_escape")
    except UnicodeDecodeError:
        pass
    try:
        text = text.encode("latin1").decode("utf-8")
    except UnicodeError:
        pass
    return text


def load_jleague_club_player_urls():
    global _JLEAGUE_CLUB_PLAYER_URLS
    if _JLEAGUE_CLUB_PLAYER_URLS is not None:
        return _JLEAGUE_CLUB_PLAYER_URLS

    urls = {}
    for list_url in JLEAGUE_CLUB_LIST_URLS:
        soup = BeautifulSoup(request_text(list_url), "lxml")
        for link in soup.select("a[href*='/club/']"):
            name = link.get_text(" ", strip=True)
            href = link.get("href", "")
            match = re.search(r"/club/([^/]+)/?", href)
            if not name or not match or name == "クラブ":
                continue
            urls[norm_text(name)] = urljoin(JLEAGUE_BASE, f"/club/{match.group(1)}/player/")

    _JLEAGUE_CLUB_PLAYER_URLS = urls
    return urls


def jleague_player_url_from_club_page(name_ja, club):
    if not name_ja.strip() or not club.strip():
        return ""

    club_urls = load_jleague_club_player_urls()
    roster_url = club_urls.get(norm_text(club))
    if not roster_url:
        return ""

    html = request_text(roster_url)
    target = norm_text(name_ja)
    pattern = (
        r'playerId\\":\\"(\d+)\\".*?'
        r'playerName\\":\\"(.*?)\\".*?'
        r'playerNameEn\\":\\"(.*?)\\"'
    )
    for player_id, player_name, _player_name_en in re.findall(pattern, html):
        if norm_text(repair_jleague_text(player_name)) == target:
            return urljoin(JLEAGUE_BASE, f"/player/{player_id}/")
    return ""


def jleague_player_url_from_search_page(name_ja):
    if not name_ja.strip():
        return ""

    query = quote(name_ja.strip())
    search_urls = [
        f"https://www.jleague.jp/search/?q={query}",
        f"https://www.jleague.jp/search/?keyword={query}",
    ]
    patterns = [
        r"https?://www\.jleague\.jp/player/\d+/?",
        r"/player/\d+/?",
        r"\\/player\\/\d+\\/?",
    ]
    for search_url in search_urls:
        html = request_text(search_url)
        for pattern in patterns:
            match = re.search(pattern, html)
            if not match:
                continue
            path = match.group(0).replace("\\/", "/")
            if path.startswith("http"):
                return path
            return urljoin("https://www.jleague.jp", path)
    return ""


def _open_jleague_search_box(page, timeout):
    for selector in (
        "button.o-navigation-bar__search-trigger",
        "button.o-navigation-bar__search-button",
        "button[aria-label='検索']",
    ):
        locator = page.locator(selector)
        if locator.count() == 0:
            continue
        try:
            locator.first.click(force=True, timeout=timeout)
            page.wait_for_timeout(1200)
            break
        except Exception:
            continue

    for locator in (
        page.get_by_role("searchbox"),
        page.get_by_role("textbox"),
        page.locator("input[type='search'], input[name*='keyword'], input[name*='q']"),
    ):
        if locator.count() > 0:
            box = locator.first
            box.wait_for(state="visible", timeout=timeout)
            return box
    raise RuntimeError("検索入力欄が見つかりませんでした")


def jleague_player_url_from_browser_search(name_ja):
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except ImportError:
        return ""

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(
            locale="ja-JP",
            user_agent=REQUEST_HEADERS["User-Agent"],
            viewport={"width": 1440, "height": 900},
        )
        page = ctx.new_page()
        try:
            page.goto(JLEAGUE_BASE, wait_until="domcontentloaded", timeout=15000)
            page.wait_for_timeout(3000)
            box = _open_jleague_search_box(page, 8000)
            box.fill(name_ja)
            box.press("Enter")
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except PlaywrightTimeoutError:
                pass
            links = page.eval_on_selector_all(
                "main a[href], [role='dialog'] a[href], .google-search-results a[href]",
                """els => els.map(e => ({
                    text: (e.innerText || '').trim(),
                    href: e.getAttribute('href')
                })).filter(x => x.href)""",
            )
            target = norm_text(name_ja)
            for item in links:
                href = urljoin(JLEAGUE_BASE, item["href"])
                if re.search(r"/player/\d+/?", href) and (
                    not item["text"] or target in norm_text(item["text"])
                ):
                    return href
            html = page.content()
            match = re.search(r"https?://www\.jleague\.jp/player/\d+/?|/player/\d+/?", html)
            if match:
                return urljoin(JLEAGUE_BASE, match.group(0))
            return ""
        finally:
            ctx.close()
            browser.close()


def find_jleague_player_url(name_ja, club="", football_lab_url=""):
    if not name_ja.strip():
        return ""

    last_error = None
    for finder in (
        lambda: jleague_player_url_from_club_page(name_ja, club),
        lambda: jleague_player_url_from_search_page(name_ja),
        lambda: jleague_player_url_from_browser_search(name_ja),
    ):
        try:
            url = finder()
        except Exception as e:
            last_error = e
            continue
        if url:
            return url
    if last_error:
        raise last_error
    return ""


def find_espn_player(name_en):
    if not name_en.strip():
        return {}

    query = quote(name_en.strip())
    url = (
        "https://site.web.api.espn.com/apis/search/v2"
        f"?query={query}&limit=5&region=us&lang=en&site=espn"
    )
    data = request_json(url)
    for result in data.get("results", []):
        if result.get("type") != "player":
            continue
        for item in result.get("contents", []):
            web_url = ((item.get("link") or {}).get("web") or "").strip()
            if "/soccer/player/" not in web_url:
                continue
            match = re.search(r"/id/(\d+)", web_url)
            espn_id = match.group(1) if match else ""
            league_slug = item.get("defaultLeagueSlug", "")
            team_name = espn_team_from_player_detail(espn_id, league_slug)
            return {
                "url": web_url,
                "espn_id": espn_id,
                "espn_team": team_name or item.get("subtitle", ""),
            }
    return {}


class PlayersEditor(tk.Tk):
    def __init__(self, config_path):
        super().__init__()
        self.config_path = Path(config_path)
        self.original_data, self.players = load_players_config(self.config_path)
        self.selected_index = None
        self.dirty = False
        self.vars = {field: tk.StringVar(self) for field in KNOWN_FORM_FIELDS}
        self.source_var = self.vars["source"]
        self.source_display_var = tk.StringVar(self)
        self.search_after_id = None
        self.search_generation = 0
        self.loading_form = False

        self.title(f"選手設定の編集 - {self.config_path.name} (LoanWatch {APP_VERSION})")
        self.geometry("1180x760")
        self.minsize(960, 640)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.create_widgets()
        self.vars["name_ja"].trace_add("write", self.on_name_change)
        self.vars["name_en"].trace_add("write", self.on_name_change)
        self.refresh_table()
        self.clear_form()

    def create_widgets(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main = ttk.Frame(self, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)

        table_frame = ttk.LabelFrame(main, text="登録選手")
        table_frame.grid(row=0, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            table_frame,
            columns=TABLE_COLUMNS,
            show="headings",
            selectmode="browse",
        )
        for column in TABLE_COLUMNS:
            self.tree.heading(column, text=HEADINGS[column])
            width = 120
            if column in ("url", "jleague_url"):
                width = 260
            elif column in ("club", "besoccer"):
                width = 180
            self.tree.column(column, width=width, minwidth=80, stretch=True)

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.on_table_select)

        form_frame = ttk.LabelFrame(main, text="設定フォーム")
        form_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        form_frame.columnconfigure(1, weight=1)
        form_frame.columnconfigure(3, weight=1)

        ttk.Label(form_frame, text="source *").grid(
            row=0, column=0, sticky="w", padx=(8, 4), pady=4
        )
        self.source_combo = ttk.Combobox(
            form_frame,
            textvariable=self.source_display_var,
            values=SOURCE_DISPLAY_CHOICES,
            state="readonly",
        )
        self.source_combo.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=4)
        self.source_combo.bind("<<ComboboxSelected>>", self.on_source_change)

        ttk.Label(form_frame, text="選手名(日本語) *").grid(
            row=1, column=0, sticky="w", padx=(8, 4), pady=4
        )
        ttk.Entry(form_frame, textvariable=self.vars["name_ja"]).grid(
            row=1, column=1, sticky="ew", padx=(0, 8), pady=4
        )
        ttk.Label(form_frame, text="選手名(英語)").grid(
            row=1, column=2, sticky="w", padx=(8, 4), pady=4
        )
        ttk.Entry(form_frame, textvariable=self.vars["name_en"]).grid(
            row=1, column=3, sticky="ew", padx=(0, 8), pady=4
        )

        self.source_frame = ttk.Frame(form_frame)
        self.source_frame.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(4, 0))
        self.source_frame.columnconfigure(1, weight=1)
        self.source_frame.columnconfigure(3, weight=1)

        button_frame = ttk.Frame(main)
        button_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        button_frame.columnconfigure(5, weight=1)

        ttk.Button(button_frame, text="選択解除", command=self.clear_selection).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(button_frame, text="追加", command=self.add_player).grid(
            row=0, column=1, padx=(0, 6)
        )
        ttk.Button(button_frame, text="行を上書き", command=self.update_player).grid(
            row=0, column=2, padx=(0, 6)
        )
        ttk.Button(button_frame, text="削除", command=self.delete_player).grid(
            row=0, column=3, padx=(0, 18)
        )
        ttk.Button(button_frame, text="JSON保存", command=self.save_json).grid(
            row=0, column=4, padx=(0, 6)
        )

        self.status_var = tk.StringVar(self)
        ttk.Label(button_frame, textvariable=self.status_var).grid(
            row=0, column=5, sticky="e"
        )

        self.rebuild_source_fields()

    def rebuild_source_fields(self):
        for child in self.source_frame.winfo_children():
            child.destroy()

        source = self.source_var.get().strip()
        if not source:
            ttk.Label(self.source_frame, text="最初に source を選択してください").grid(
                row=0, column=0, sticky="w", padx=(8, 4), pady=4
            )
            return

        if source == "football-lab":
            ttk.Label(self.source_frame, text="クラブ *").grid(
                row=0, column=0, sticky="w", padx=(8, 4), pady=4
            )
            self.club_combo = ttk.Combobox(
                self.source_frame,
                textvariable=self.vars["club"],
                values=TEAM_CHOICES,
                state="readonly",
            )
            self.club_combo.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=4)
            self.club_combo.bind("<<ComboboxSelected>>", self.on_club_change)

            ttk.Label(self.source_frame, text="Football-LAB URL *").grid(
                row=0, column=2, sticky="w", padx=(8, 4), pady=4
            )
            ttk.Entry(self.source_frame, textvariable=self.vars["url"]).grid(
                row=0, column=3, sticky="ew", padx=(0, 8), pady=4
            )

            ttk.Label(self.source_frame, text="Jリーグ公式URL").grid(
                row=1, column=0, sticky="w", padx=(8, 4), pady=4
            )
            ttk.Entry(self.source_frame, textvariable=self.vars["jleague_url"]).grid(
                row=1, column=1, columnspan=3, sticky="ew", padx=(0, 8), pady=4
            )
            self.build_besoccer_fields(2)
            return

        if source == "espn":
            ttk.Label(self.source_frame, text="クラブ *").grid(
                row=0, column=0, sticky="w", padx=(8, 4), pady=4
            )
            ttk.Entry(self.source_frame, textvariable=self.vars["club"]).grid(
                row=0, column=1, sticky="ew", padx=(0, 8), pady=4
            )
            ttk.Label(self.source_frame, text="ESPN URL *").grid(
                row=0, column=2, sticky="w", padx=(8, 4), pady=4
            )
            ttk.Entry(self.source_frame, textvariable=self.vars["url"]).grid(
                row=0, column=3, sticky="ew", padx=(0, 8), pady=4
            )

            ttk.Label(self.source_frame, text="ESPNチームID").grid(
                row=1, column=0, sticky="w", padx=(8, 4), pady=4
            )
            ttk.Entry(self.source_frame, textvariable=self.vars["espn_team_id"]).grid(
                row=1, column=1, sticky="ew", padx=(0, 8), pady=4
            )
            ttk.Label(
                self.source_frame,
                text="(espn.com/soccer/team/results/_/id/<ここ> 。空欄可、鮮度チェック用)",
            ).grid(row=1, column=2, columnspan=2, sticky="w", padx=(8, 4), pady=4)
            self.build_besoccer_fields(2)

    def build_besoccer_fields(self, row):
        """BeSoccerから追加の1行を出すための任意項目を並べる

        ESPN・Football-LAB はトップチームの記録しか持たないため、リザーブチーム
        (例: アルヘンティノス・ジュニアーズII)の出場はここに書いて補う。
        """
        frame = ttk.LabelFrame(
            self.source_frame,
            text="BeSoccer(リザーブチームなどの行を追加する。すべて空欄なら使わない)",
        )
        frame.grid(row=row, column=0, columnspan=4, sticky="ew",
                   padx=(8, 8), pady=(8, 4))
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(3, weight=1)

        ttk.Label(frame, text="BeSoccer選手URL").grid(
            row=0, column=0, sticky="w", padx=(8, 4), pady=4
        )
        ttk.Entry(frame, textvariable=self.vars["bs_url"]).grid(
            row=0, column=1, columnspan=3, sticky="ew", padx=(0, 8), pady=4
        )
        ttk.Label(
            frame,
            text="(besoccer.com/player/matches/<選手> のURL)",
        ).grid(row=1, column=0, columnspan=4, sticky="w", padx=(8, 4))

        ttk.Label(frame, text="リーグロゴ名").grid(
            row=2, column=0, sticky="w", padx=(8, 4), pady=4
        )
        ttk.Entry(frame, textvariable=self.vars["bs_logo"]).grid(
            row=2, column=1, sticky="ew", padx=(0, 8), pady=4
        )
        ttk.Label(frame, text="BeSoccerチーム名").grid(
            row=2, column=2, sticky="w", padx=(8, 4), pady=4
        )
        ttk.Entry(frame, textvariable=self.vars["bs_team"]).grid(
            row=2, column=3, sticky="ew", padx=(0, 8), pady=4
        )
        ttk.Label(
            frame,
            text="(ロゴ名は試合一覧の行頭画像 league_logos/<ここ>.png "
                 "例: torneo-proyeccion / チーム名は例: Argentinos Juniors II)",
        ).grid(row=3, column=0, columnspan=4, sticky="w", padx=(8, 4))

        ttk.Label(frame, text="クラブ表示名").grid(
            row=4, column=0, sticky="w", padx=(8, 4), pady=4
        )
        ttk.Entry(frame, textvariable=self.vars["bs_club"]).grid(
            row=4, column=1, sticky="ew", padx=(0, 8), pady=4
        )
        ttk.Label(frame, text="リーグ表示名").grid(
            row=4, column=2, sticky="w", padx=(8, 4), pady=4
        )
        ttk.Entry(frame, textvariable=self.vars["bs_league"]).grid(
            row=4, column=3, sticky="ew", padx=(0, 8), pady=4
        )
        ttk.Label(
            frame,
            text="(表示名は空欄可。出力の クラブ / リーグ名 列に出す日本語表記)",
        ).grid(row=5, column=0, columnspan=4, sticky="w", padx=(8, 4), pady=(0, 4))

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        for i, player in enumerate(self.players):
            values = []
            for column in TABLE_COLUMNS:
                value = player.get(column, "")
                if column == "source":
                    value = SOURCE_VALUE_TO_DISPLAY.get(value, value)
                elif column == "besoccer":
                    value = besoccer_summary(player)
                values.append(value)
            self.tree.insert("", "end", iid=str(i), values=values)
        self.status_var.set(f"{len(self.players)} 件 / {self.config_path}")

    def on_table_select(self, _event=None):
        selected = self.tree.selection()
        if not selected:
            return
        self.selected_index = int(selected[0])
        self.populate_form(self.players[self.selected_index])

    def populate_form(self, player):
        self.cancel_pending_search()
        self.search_generation += 1
        self.loading_form = True
        for field in KNOWN_FORM_FIELDS:
            self.vars[field].set(str(player.get(field, "")))
        self.populate_besoccer(player)
        self.source_display_var.set(
            SOURCE_VALUE_TO_DISPLAY.get(self.source_var.get(), "")
        )
        self.rebuild_source_fields()
        self.loading_form = False

    def populate_besoccer(self, player):
        """besoccer ブロックの1件目を bs_* の入力欄へ展開する"""
        conf = player.get("besoccer") or {}
        teams = [t for t in (conf.get("teams") or []) if isinstance(t, dict)]
        first = teams[0] if teams else {}
        self.vars["bs_url"].set(str(conf.get("url", "")))
        self.vars["bs_logo"].set(str(first.get("logo", "")))
        self.vars["bs_team"].set(str(first.get("team", "")))
        self.vars["bs_club"].set(str(first.get("club", "")))
        self.vars["bs_league"].set(str(first.get("league", "")))

    def collect_besoccer(self, base):
        """bs_* の入力欄を players.json の besoccer ブロックにする

        画面で編集できるのは teams の1件目だけなので、2件目以降は元のまま残す。
        """
        values = {field: self.vars[field].get().strip()
                  for field in BESOCCER_FORM_FIELDS}
        if not any(values.values()):
            return None
        team = {"logo": values["bs_logo"], "team": values["bs_team"]}
        if values["bs_club"]:
            team["club"] = values["bs_club"]
        if values["bs_league"]:
            team["league"] = values["bs_league"]
        rest = [t for t in ((base or {}).get("besoccer") or {}).get("teams", [])[1:]
                if isinstance(t, dict)]
        return {"url": values["bs_url"], "teams": [team] + rest}

    def clear_form(self):
        self.cancel_pending_search()
        self.search_generation += 1
        self.loading_form = True
        for var in self.vars.values():
            var.set("")
        self.source_display_var.set("")
        self.rebuild_source_fields()
        self.loading_form = False

    def clear_selection(self):
        self.tree.selection_remove(self.tree.selection())
        self.selected_index = None
        self.clear_form()

    def on_source_change(self, _event=None):
        self.cancel_pending_search()
        self.search_generation += 1
        source = SOURCE_DISPLAY_TO_VALUE.get(self.source_display_var.get(), "")
        self.source_var.set(source)
        for field in ("club", "url", "jleague_url", "espn_id", "espn_team"):
            self.vars[field].set("")
        self.rebuild_source_fields()
        self.schedule_player_search()

    def cancel_pending_search(self):
        if self.search_after_id is not None:
            self.after_cancel(self.search_after_id)
            self.search_after_id = None

    def on_club_change(self, _event=None):
        if self.source_var.get() != "football-lab":
            return
        url = TEAM_URLS.get(self.vars["club"].get().strip())
        if url:
            self.vars["url"].set(url)
        self.schedule_player_search()

    def on_name_change(self, *_args):
        if self.loading_form:
            return
        self.schedule_player_search()

    def schedule_player_search(self):
        if self.search_after_id is not None:
            self.after_cancel(self.search_after_id)
            self.search_after_id = None
        source = self.source_var.get().strip()
        if source == "football-lab" and self.vars["name_ja"].get().strip():
            self.search_after_id = self.after(3000, self.start_player_search)
        elif source == "espn" and self.vars["name_en"].get().strip():
            self.search_after_id = self.after(3000, self.start_player_search)

    def start_player_search(self):
        self.search_after_id = None
        source = self.source_var.get().strip()
        name_ja = self.vars["name_ja"].get().strip()
        name_en = self.vars["name_en"].get().strip()
        self.search_generation += 1
        generation = self.search_generation

        if source == "football-lab":
            self.status_var.set(f"Jリーグ公式URLを検索中: {name_ja}")
            worker = threading.Thread(
                target=self.search_jleague_worker,
                args=(
                    generation,
                    name_ja,
                    self.vars["club"].get().strip(),
                    self.vars["url"].get().strip(),
                ),
                daemon=True,
            )
        elif source == "espn":
            self.status_var.set(f"ESPN URLを検索中: {name_en}")
            worker = threading.Thread(
                target=self.search_espn_worker,
                args=(generation, name_en),
                daemon=True,
            )
        else:
            return
        worker.start()

    def search_jleague_worker(self, generation, name_ja, club, football_lab_url):
        try:
            url = find_jleague_player_url(name_ja, club=club, football_lab_url=football_lab_url)
            error = ""
        except Exception as e:
            url = ""
            error = f"{type(e).__name__}: {e}"
        self.after(0, self.apply_jleague_search_result, generation, name_ja, url, error)

    def apply_jleague_search_result(self, generation, name_ja, url, error):
        if generation != self.search_generation or self.source_var.get() != "football-lab":
            return
        if self.vars["name_ja"].get().strip() != name_ja:
            return
        if url:
            self.vars["jleague_url"].set(url)
            self.status_var.set(f"Jリーグ公式URLを設定しました: {url}")
        elif error:
            self.status_var.set(f"Jリーグ公式URLを取得できませんでした: {error}")
        else:
            self.status_var.set("Jリーグ公式URLは見つかりませんでした")

    def search_espn_worker(self, generation, name_en):
        try:
            player = find_espn_player(name_en)
            error = ""
        except Exception as e:
            player = {}
            error = f"{type(e).__name__}: {e}"
        self.after(0, self.apply_espn_search_result, generation, name_en, player, error)

    def apply_espn_search_result(self, generation, name_en, player, error):
        if generation != self.search_generation or self.source_var.get() != "espn":
            return
        if self.vars["name_en"].get().strip() != name_en:
            return
        if player:
            self.vars["url"].set(player.get("url", ""))
            self.vars["espn_id"].set(player.get("espn_id", ""))
            self.vars["espn_team"].set(player.get("espn_team", ""))
            if player.get("espn_team") and not self.vars["club"].get().strip():
                self.vars["club"].set(player["espn_team"])
            self.status_var.set(f"ESPN URLを設定しました: {player.get('url', '')}")
        elif error:
            self.status_var.set(f"ESPN URLを取得できませんでした: {error}")
        else:
            self.status_var.set("ESPN URLは見つかりませんでした")

    def validate_form(self):
        source = self.source_var.get().strip()
        if source not in SOURCE_VALUE_TO_DISPLAY:
            messagebox.showerror("入力エラー", "source を選択してください")
            return False

        if source == "espn" and not self.vars["espn_id"].get().strip():
            match = re.search(r"/id/(\d+)", self.vars["url"].get().strip())
            if match:
                self.vars["espn_id"].set(match.group(1))

        required_fields = [
            ("name_ja", "選手名(日本語)"),
            ("club", "クラブ"),
            ("url", "URL"),
        ]
        if source == "espn":
            required_fields.append(("name_en", "選手名(英語)"))
            required_fields.append(("espn_id", "ESPN ID"))
        missing = [label for field, label in required_fields if not self.vars[field].get().strip()]
        # BeSoccer欄は丸ごと空欄なら使わない。1つでも書いたら必須3項目をそろえる
        if any(self.vars[field].get().strip() for field in BESOCCER_FORM_FIELDS):
            missing += [label for field, label in BESOCCER_REQUIRED
                        if not self.vars[field].get().strip()]
        if missing:
            messagebox.showerror(
                "入力エラー",
                "必須項目を入力してください:\n" + "\n".join(f"- {label}" for label in missing),
            )
            return False
        return True

    def collect_form(self, base=None):
        if not self.validate_form():
            return None

        source = self.source_var.get().strip()
        # besoccer は最後に組み立て直すので、ここではいったん外しておく
        player = {
            key: value
            for key, value in (base or {}).items()
            if key not in KNOWN_FORM_FIELDS and key != "besoccer"
        }

        for field in OUTPUT_FIELDS_BY_SOURCE.get(source, ()):
            value = source if field == "source" else self.vars[field].get().strip()
            if value:
                player[field] = value

        # BeSoccer欄をすべて空にして上書きすると、設定ごと消える
        besoccer = self.collect_besoccer(base)
        if besoccer:
            player["besoccer"] = besoccer
        return player

    def add_player(self):
        if self.selected_index is not None or self.tree.selection():
            messagebox.showinfo("追加", "追加する場合は、先に「選択解除」をクリックしてください")
            return
        player = self.collect_form()
        if player is None:
            return
        self.players.append(player)
        self.dirty = True
        self.refresh_table()
        self.tree.selection_set(str(len(self.players) - 1))
        self.tree.see(str(len(self.players) - 1))

    def update_player(self):
        if self.selected_index is None:
            messagebox.showinfo("行を上書き", "上書きする行を表から選択してください")
            return
        player = self.collect_form(base=self.players[self.selected_index])
        if player is None:
            return
        self.players[self.selected_index] = player
        self.dirty = True
        selected = self.selected_index
        self.refresh_table()
        self.tree.selection_set(str(selected))
        self.tree.see(str(selected))

    def delete_player(self):
        if self.selected_index is None:
            messagebox.showinfo("削除", "削除する行を表から選択してください")
            return
        player = self.players[self.selected_index]
        name = player.get("name_ja") or player.get("name_en") or "(名前なし)"
        if not messagebox.askyesno("削除確認", f"{name} を削除しますか？"):
            return
        del self.players[self.selected_index]
        self.selected_index = None
        self.dirty = True
        self.refresh_table()
        self.clear_form()

    def save_json(self):
        output_data = build_output_data(self.original_data, self.players)
        backup_path = backup_path_for(self.config_path)
        temp_path = self.config_path.with_name(f".{self.config_path.name}.tmp")
        try:
            shutil.copy2(self.config_path, backup_path)
            with temp_path.open("w", encoding="utf-8", newline="\n") as f:
                json.dump(output_data, f, ensure_ascii=False, indent=2)
                f.write("\n")
            temp_path.replace(self.config_path)
        except OSError as e:
            if temp_path.exists():
                temp_path.unlink()
            messagebox.showerror("保存エラー", f"保存に失敗しました:\n{e}")
            return

        self.original_data = output_data
        self.dirty = False
        self.status_var.set(f"保存しました。バックアップ: {backup_path.name}")
        messagebox.showinfo(
            "保存完了",
            f"{self.config_path.name} を保存しました。\n"
            f"バックアップ: {backup_path.name}",
        )

    def on_close(self):
        if self.dirty:
            ok = messagebox.askyesno(
                "未保存の変更",
                "JSON保存していない変更があります。終了しますか？",
            )
            if not ok:
                return
        self.destroy()


def parse_args():
    parser = argparse.ArgumentParser(description="players.json GUI editor")
    parser.add_argument(
        "players_config",
        nargs="?",
        default=str(DEFAULT_PLAYERS_CONFIG),
        help="編集する players.json のパス",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    app = PlayersEditor(args.players_config)
    app.mainloop()


if __name__ == "__main__":
    main()
