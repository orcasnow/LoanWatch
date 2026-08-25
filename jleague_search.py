"""
jleague.jp のヘッダー検索ボタンを外部から叩いて検索を実行するスクリプト。

前提:
    pip install playwright
    playwright install chromium

使い方:
    python jleague_search.py "名古屋グランパス"
    python jleague_search.py "河面旺成" --capture   # 裏で飛ぶAPIを覗く
"""

from __future__ import annotations

import argparse
import json
import sys
from urllib.parse import urljoin

from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

BASE = "https://www.jleague.jp/"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)


def _open_search_box(page, timeout: int):
    """検索ボタンをクリックして、入力欄のロケータを返す。"""
    # 第一候補: アクセシビリティ名（aria-label="検索"）で取る。最も壊れにくい。
    btn = page.get_by_role("button", name="検索")
    if btn.count() == 0:
        # 第二候補: サイト独自のBEMクラス。rt-* はRadix生成なので使わない。
        btn = page.locator("button.o-navigation-bar__search-button")
    btn.first.wait_for(state="visible", timeout=timeout)
    btn.first.click()

    # 入力欄。role=searchbox が無ければ textbox、それも無ければ input[type=search]
    for locator in (
        page.get_by_role("searchbox"),
        page.get_by_role("textbox"),
        page.locator("input[type='search'], input[name*='keyword'], input[name*='q']"),
    ):
        if locator.count() > 0:
            box = locator.first
            box.wait_for(state="visible", timeout=timeout)
            return box
    raise RuntimeError("検索入力欄が見つかりませんでした（DOM構造が変わった可能性）")


def search(keyword: str, headless: bool = True, timeout: int = 15000, dump: bool = False):
    """keyword で検索を実行し、結果ページのURLとリンク一覧を返す。"""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        ctx = browser.new_context(
            locale="ja-JP",
            user_agent=UA,
            viewport={"width": 1440, "height": 900},  # 狭いとヘッダーが畳まれる
        )
        page = ctx.new_page()
        page.goto(BASE, wait_until="domcontentloaded", timeout=timeout)

        box = _open_search_box(page, timeout)
        box.fill(keyword)
        box.press("Enter")

        # Next.jsのクライアント遷移なのでURL変化 or ネットワーク静止で待つ
        try:
            page.wait_for_url(lambda url: url.rstrip("/") != BASE.rstrip("/"), timeout=timeout)
        except PWTimeout:
            pass  # 同一URL上でモーダル内に結果を出すタイプの可能性
        try:
            page.wait_for_load_state("networkidle", timeout=timeout)
        except PWTimeout:
            pass

        result_url = page.url
        links = page.eval_on_selector_all(
            "main a[href], [role='dialog'] a[href]",
            """els => els.map(e => ({
                text: (e.innerText || '').trim().slice(0, 120),
                href: e.getAttribute('href')
            })).filter(x => x.text && x.href)""",
        )
        # 内部リンクだけ、重複除去
        seen, results = set(), []
        for item in links:
            href = urljoin(BASE, item["href"])
            if "jleague.jp" not in href or href in seen:
                continue
            seen.add(href)
            results.append({"text": item["text"], "url": href})

        if dump:
            page.screenshot(path="jleague_search.png", full_page=True)
            with open("jleague_search.html", "w", encoding="utf-8") as f:
                f.write(page.content())

        ctx.close()
        browser.close()
        return result_url, results


def capture_requests(keyword: str, timeout: int = 15000):
    """検索実行時に飛ぶXHR/fetchを記録する。API直叩きに切り替えるための下調べ用。"""
    captured = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # 初回は目視推奨
        page = browser.new_context(locale="ja-JP", user_agent=UA,
                                   viewport={"width": 1440, "height": 900}).new_page()

        page.on("request", lambda req: captured.append(
            {"method": req.method, "url": req.url, "post_data": req.post_data}
        ) if req.resource_type in ("xhr", "fetch") else None)

        page.goto(BASE, wait_until="domcontentloaded", timeout=timeout)
        before = len(captured)
        box = _open_search_box(page, timeout)
        box.fill(keyword)
        box.press("Enter")
        page.wait_for_timeout(5000)

        print("遷移後URL:", page.url)
        for req in captured[before:]:
            print(json.dumps(req, ensure_ascii=False))
        browser.close()
    return captured


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("keyword")
    parser.add_argument("--capture", action="store_true", help="裏で飛ぶAPIを調べる")
    parser.add_argument("--show", action="store_true", help="ブラウザを表示する")
    parser.add_argument("--dump", action="store_true", help="スクショとHTMLを保存")
    args = parser.parse_args()

    if args.capture:
        capture_requests(args.keyword)
        return

    url, results = search(args.keyword, headless=not args.show, dump=args.dump)
    print("検索結果ページ:", url)
    for r in results:
        print(f"- {r['text']}\n  {r['url']}")
    if not results:
        print("結果が取れませんでした。--show --dump で実際の画面を確認してください。",
              file=sys.stderr)


if __name__ == "__main__":
    main()
