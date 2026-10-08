import asyncio
from playwright.async_api import async_playwright
import os

async def generate_pdf():
    html_path = 'file:///' + os.path.abspath('ebook.html').replace('\\', '/')
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(html_path, wait_until='networkidle')
        await page.pdf(path='Innovation_Lab_Manual.pdf', format='A4', print_background=True)
        await browser.close()

if __name__ == '__main__':
    asyncio.run(generate_pdf())
