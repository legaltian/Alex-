---
name: alex-pkulaw-downloader
description: >
  Use when the user provides a 北大法宝 case URL, a case number (案号),
  or a saved PKULaw HTML file. Triggers on: 北大法宝, pkulaw.com, 案号,
  下载判决书, 搜这个案子, 下载案例, 查一下这个案号, .html 文件.
---

# /alex-pkulaw-downloader

## 概述

从北大法宝下载判例为 Markdown，提取元数据 + 裁判规则 + 判决书全文。

输出到 `~/Desktop/北大法宝判例/`。

## 前置条件

Chrome 调试模式（一次性配置）：

```bash
pkill -f "Google Chrome" 2>/dev/null; sleep 2
mkdir -p /tmp/cdp-chrome
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \
  --remote-debugging-port=9222 \
  --user-data-dir=/tmp/cdp-chrome > /dev/null 2>&1 &
```

## 工作流程

### 输入：HTML 文件

直接本地提取：

```bash
python3 ~/.claude/skills/alex-pkulaw-downloader/extract.py "文件路径.html" "原始URL"
```

### 输入：判例 URL

CDP 抓取 HTML → 本地提取：

```bash
python3 << 'PYEOF'
import asyncio
from playwright.async_api import async_playwright

url = "用户提供的链接"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        page = await browser.new_page()
        await page.goto(url, wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(3000)
        html = await page.content()
        await page.close()
        with open('/tmp/pkulaw_case.html', 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"✅ HTML: {len(html)} 字")

asyncio.run(main())
PYEOF

python3 ~/.claude/skills/alex-pkulaw-downloader/extract.py /tmp/pkulaw_case.html "URL"
```

### 输入：案号

先搜索 → 再 CDP 抓取 → 本地提取：

```bash
python3 << 'PYEOF'
import asyncio
from playwright.async_api import async_playwright

case_num = "用户提供的案号"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        page = await browser.new_page()
        
        print(f"🔍 搜索: {case_num}")
        await page.goto("https://www.pkulaw.com/case/", wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(3000)
        
        # 切换搜索字段为「案号」
        dropdown = page.locator('.areajSelect')
        await dropdown.hover()
        await page.wait_for_timeout(300)
        await dropdown.click()
        await page.wait_for_timeout(800)
        try:
            await page.locator('a[areacode="Code"]').click(timeout=5000)
        except:
            await page.evaluate("document.querySelector('a[areacode=\"Code\"]')?.click()")
        await page.wait_for_timeout(500)
        
        await page.fill('#txtSearch', case_num)
        await page.press('#txtSearch', 'Enter')
        await page.wait_for_timeout(5000)
        await page.wait_for_load_state("networkidle")
        
        # 取第一个结果（优先 /gac/）
        links = await page.query_selector_all('a[href*="/gac/"], a[href*="/pfnl/"]')
        if not links:
            print("❌ 未找到结果")
            return
        
        gac_link = pfnl_link = None
        for link in links:
            href = await link.get_attribute('href') or ''
            if '/gac/' in href and not gac_link: gac_link = link
            if '/pfnl/' in href and not pfnl_link: pfnl_link = link
        
        result_link = gac_link or pfnl_link
        href = await result_link.get_attribute('href')
        txt = await result_link.inner_text()
        case_url = f"https://www.pkulaw.com{href}" if href.startswith('/') else href
        print(f"📄 {txt[:60]}")
        
        await page.goto(case_url, wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(3000)
        html = await page.content()
        await page.close()
        
        with open('/tmp/pkulaw_case.html', 'w', encoding='utf-8') as f:
            f.write(html)
        with open('/tmp/pkulaw_case_url.txt', 'w') as f:
            f.write(case_url)
        print(f"✅ HTML: {len(html)} 字")

asyncio.run(main())
PYEOF

python3 ~/.claude/skills/alex-pkulaw-downloader/extract.py /tmp/pkulaw_case.html "$(cat /tmp/pkulaw_case_url.txt)"
```

## 输出格式

```
# 案件名称

**案号：** ...
**审理法院：** ...
**审结日期：** ...
**审理程序：** ...
**文书类型：** ...
**案由：** ...
**来源：** 北大法宝
**原始链接：** ...

---

## 裁判规则
**关键词：** ...
**核心问题：** ...
**裁判要点：** ...

---

## 判决书全文
### 当事人
...
### 审理经过
...
### 本院认为
...
### 裁判结果
...
### 落款
...
```

## 依赖

- `playwright`：`pip3 install --break-system-packages playwright`
- Chrome 浏览器
- Python 3 标准库（extract.py 零额外依赖）
