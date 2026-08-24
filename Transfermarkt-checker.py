"""Transfermarktの選手ページをPlaywrightで取得できるか確認するスクリプト

事前準備:
    pip install playwright
    playwright install chromium
実行:
    python tm_check.py
    python tm_check.py --url https://... --headed
"""

import argparse
import asyncio
import re
import sys

from playwright.async_api import async_playwright

DEFAULT_URL = "https://www.transfermarkt.jp/ryoga-kida/leistungsdatendetails/spieler/916601"


async def fetch_rendered(url: str, wait_ms: int = 4000, headless: bool = True) -> str:
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=headless)
        ctx = await browser.new_context(
            locale="ja-JP",
            timezone_id="Asia/Tokyo",
            viewport={"width": 1366, "height": 900},
        )
        page = await ctx.new_page()
        try:
            resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            print(f"HTTP {resp.status if resp else '?'} / title={await page.title()!r}")

            for sel in ["button:has-text('同意')", "button:has-text('Accept')",
                        "#onetrust-accept-btn-handler"]:
                try:
                    await page.click(sel, timeout=2000)
                    print(f"Cookie同意をクリック: {sel}")
                    break
                except Exception:
                    pass

            try:
                await page.wait_for_selector("table.items", timeout=wait_ms)
                print("table.items を検出")
            except Exception:
                print("table.items が現れませんでした")

            html = await page.content()
            print(f"レンダリング後: {len(html)}文字 / table数: {html.count('<table')}")
            return html
        finally:
            await browser.close()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", default=DEFAULT_URL)
    ap.add_argument("--wait", type=int, default=4000, help="table待機のミリ秒")
    ap.add_argument("--headed", action="store_true", help="ブラウザ画面を表示する")
    ap.add_argument("--out", default="page.html", help="保存先HTMLファイル")
    args = ap.parse_args()

    try:
        html = asyncio.run(
            fetch_rendered(args.url, wait_ms=args.wait, headless=not args.headed))
    except Exception as e:
        print(f"取得に失敗しました: {type(e).__name__}: {e}", file=sys.stderr)
        return 1

    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTMLを保存しました: {args.out}")

    if re.search(r"(captcha|Just a moment|Access denied|Attention Required|cf-)", html, re.I):
        print("→ ブロックページの可能性があります。--headed で画面を確認してください")
    return 0


if __name__ == "__main__":
    sys.exit(main())