"""U-21 Ｊリーグの取得設定 (u21_clubs.json) の読み込み

LoanWatchMain.py(CUI)と LoanWatchGui.py(GUI)の両方から使うため、
追加ライブラリに依存しない標準ライブラリだけで書いている。

設定できるもの:
  default            … GUI・CLIの既定の対象クラブID
  season_start_month … 日程検索の開始月(この月の1日から当日までを検索する)
  schedules          … 試合日程の取得元。順位決定戦などが別大会・別URLに
                       なった場合は2件目を足せば、同じ台帳に取り込まれる
  clubs              … クラブID一覧

id は日程検索ページ
    https://www.jleague.jp/u-21/match/search-list/?...&club=<id>
の &club= の値。players.json の u21_club_id と同じ体系で、
U-21 ファジアーノ岡山 = 33189 のように対応する。

設定ファイルが無い・壊れている場合は FALLBACK_* で動く。exeを単体で配っても
「クラブ一覧が空で何も選べない」状態にならないようにするため。
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

CONFIG_NAME = "u21_clubs.json"

# u21_clubs.json が無いときに使う既定値(2026/27シーズンの11クラブ)
FALLBACK_CLUBS = [
    {"id": "33180", "name": "U-21 浦和レッズ", "group": "EAST"},
    {"id": "33181", "name": "U-21 ＦＣ東京", "group": "EAST"},
    {"id": "33182", "name": "U-21 東京ヴェルディ", "group": "EAST"},
    {"id": "33183", "name": "U-21 川崎フロンターレ", "group": "EAST"},
    {"id": "33184", "name": "U-21 清水エスパルス", "group": "EAST"},
    {"id": "33185", "name": "U-21 ジュビロ磐田", "group": "EAST"},
    {"id": "33186", "name": "U-21 名古屋グランパス", "group": "WEST"},
    {"id": "33187", "name": "U-21 ガンバ大阪", "group": "WEST"},
    {"id": "33188", "name": "U-21 ヴィッセル神戸", "group": "WEST"},
    {"id": "33189", "name": "U-21 ファジアーノ岡山", "group": "WEST"},
    {"id": "33190", "name": "U-21 Ｖ・ファーレン長崎", "group": "WEST"},
]
FALLBACK_DEFAULT_ID = "33186"  # U-21 名古屋グランパス
FALLBACK_SEASON_START_MONTH = 8
FALLBACK_SCHEDULES = [{
    "name": "U-21 Ｊリーグ",
    "path": "u-21",
    "url": ("https://www.jleague.jp/u-21/match/search-list/"
            "?startdate={start}&enddate={end}&period=custom&club={club}"),
}]


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_config_path():
    return app_dir() / CONFIG_NAME


def _clean_clubs(raw):
    clubs = []
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict):
            continue
        club_id = str(item.get("id", "")).strip()
        name = str(item.get("name", "")).strip()
        if club_id.isdigit() and name:
            clubs.append({"id": club_id, "name": name,
                          "group": str(item.get("group", "")).strip()})
    return clubs


def _clean_schedules(raw):
    """日程取得元を整える。url に {start}{end}{club} が揃っていないものは捨てる"""
    schedules = []
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict):
            continue
        url = str(item.get("url", "")).strip()
        path = str(item.get("path", "")).strip().strip("/")
        if not url or not path:
            continue
        if not all(f"{{{k}}}" in url for k in ("start", "end", "club")):
            continue
        schedules.append({"name": str(item.get("name", "")).strip() or path,
                          "path": path, "url": url})
    return schedules


def load_config(config_path=None):
    """取得設定をまとめて返す

    ファイルが無い・壊れている場合も例外にせず既定値を返す。U-21は付随機能で
    あり、ここで実行全体を止めたくないため。個々の項目も、書かれていなければ
    既定値で補う(古い u21_clubs.json をそのまま使えるように)。
    """
    path = Path(config_path or default_config_path())
    try:
        with path.open("r", encoding="utf-8") as f:
            config = json.load(f)
    except (OSError, ValueError):
        config = {}
    if not isinstance(config, dict):
        config = {"clubs": config}

    clubs = _clean_clubs(config.get("clubs")) or list(FALLBACK_CLUBS)
    ids = {c["id"] for c in clubs}
    wanted = str(config.get("default", "")).strip()
    default_id = wanted if wanted in ids else (
        FALLBACK_DEFAULT_ID if FALLBACK_DEFAULT_ID in ids else clubs[0]["id"])

    month = config.get("season_start_month", FALLBACK_SEASON_START_MONTH)
    try:
        month = int(month)
    except (TypeError, ValueError):
        month = FALLBACK_SEASON_START_MONTH
    if not 1 <= month <= 12:
        month = FALLBACK_SEASON_START_MONTH

    return {
        "clubs": clubs,
        "default_id": default_id,
        "season_start_month": month,
        "schedules": _clean_schedules(config.get("schedules"))
                     or [dict(s) for s in FALLBACK_SCHEDULES],
    }


def load_clubs(config_path=None):
    """(クラブ一覧, 既定クラブID) を返す。GUIのプルダウン用"""
    config = load_config(config_path)
    return config["clubs"], config["default_id"]


def find_club(clubs, token):
    """ID・クラブ名・その一部からクラブを1件探す。見つからなければ None

    GUIのプルダウンはクラブ名、CLIはIDでの指定が自然なので、どちらも通す。
    """
    key = str(token).strip()
    if not key:
        return None
    for club in clubs:
        if club["id"] == key:
            return club
    for club in clubs:
        if club["name"] == key:
            return club
    lowered = key.casefold()
    hits = [c for c in clubs if lowered in c["name"].casefold()]
    return hits[0] if len(hits) == 1 else None
