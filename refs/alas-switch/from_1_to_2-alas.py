import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("http://127.0.0.1:22267/")
    page.wait_for_timeout(2000)
    page.get_by_role("button", name="日常").click()
    page.wait_for_timeout(2000)
    page.get_by_role("button", name="停止").click()
    page.wait_for_timeout(2000)
    page.get_by_role("button", name="主线").click()
    page.wait_for_timeout(2000)
    page.get_by_role("button", name="启动").click()
    page.wait_for_timeout(2000)
    #page.get_by_role("button", name="启动").click()
    #page.get_by_role("button", name="停止").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
