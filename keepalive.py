"""완생 앱이 잠들지 않게 주기적으로 실제 브라우저로 방문한다.

단순 GET 요청은 페이지만 받아오고 앱(WebSocket)을 깨우지 못해서, 헤드리스 크롬(Playwright)으로 연다.
앱이 잠들어 있으면 'Yes, get this app back up!' 버튼을 눌러 깨운다.
"""
import asyncio
import sys

from playwright.async_api import async_playwright

APP_URL = "https://wan-seng-d8suknxdjamwjnavvuqdg5.streamlit.app/"
WAKE_TEXT = "get this app back up"


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(APP_URL, wait_until="domcontentloaded", timeout=90_000)
        await page.wait_for_timeout(8_000)

        button = page.get_by_role("button", name=WAKE_TEXT, exact=False)
        if await button.count() > 0:
            print("앱이 잠들어 있음 → 깨우기 버튼 클릭")
            await button.first.click()
            await page.wait_for_timeout(90_000)  # 앱이 다시 켜질 때까지 기다림
            still = await page.get_by_role("button", name=WAKE_TEXT, exact=False).count()
            if still:
                print("깨우기 실패: 버튼이 그대로 남아 있음")
                await browser.close()
                sys.exit(1)
            print("깨우기 완료")
        else:
            # 깨어 있어도 잠시 머물러 실제 방문(WebSocket 연결)으로 기록되게 한다
            await page.wait_for_timeout(30_000)
            print("앱이 깨어 있음 (방문 기록 완료)")
        await browser.close()


asyncio.run(main())
