---
name: alex-wechat-saver
description: >
  Use when the user provides a WeChat public account article URL (mp.weixin.qq.com)
  and wants to save it as a clean Markdown file. Triggers on: 微信文章链接, 公众号文章,
  mp.weixin.qq.com, 保存这篇文章, 抓取这篇文章.
---

# /wechat-article-saver

## 概述

用户丢一个微信公众号文章链接 → curl 抓取 HTML → Python 提取标题/作者/公众号/发布时间/正文 →
保存为干净 Markdown 到 `~/Desktop/公众号文件暂存/`。

文件命名：`YYYYMMDD-作者-公众号-文章标题.md`

## 工作流程

### Step 1 — 确认链接

从用户消息中提取微信公众号 URL（匹配 `mp.weixin.qq.com`）。如果用户没给链接，提示用户提供。

### Step 2 — curl 抓取 HTML

```bash
curl -sL -A "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36" \
  "<URL>" -o /tmp/wechat_article.html
```

**关键：** 必须从用户本机发出请求（微信域名被海外 IP 屏蔽，不能走 WebFetch）。

检查文件大小：`wc -c /tmp/wechat_article.html`。如果小于 10KB，可能被拦截。

### Step 3 — Python 提取并保存

```bash
python3 << 'PYEOF'
import re, html as html_module, os, datetime

url = "用户提供的链接"

with open('/tmp/wechat_article.html', 'r', encoding='utf-8') as f:
    content = f.read()

# ── 标题 ──
title_m = re.search(r'<meta\s+property="og:title"\s+content="([^"]+)"', content)
title = html_module.unescape(title_m.group(1)) if title_m else "未找到标题"

# ── 作者（作者栏显示的创作者姓名） ──
author = ""
am = re.search(r'<meta\s+property="og:article:author"\s+content="([^"]+)"', content)
if am:
    author = html_module.unescape(am.group(1)).strip()
if not author:
    am = re.search(r'var\s+author\s*=\s*["\']([^"\']+)["\']', content)
    if am:
        author = html_module.unescape(am.group(1)).strip()
if not author:
    author = "未知作者"

# ── 公众号名称 ──
account = ""
nm = re.search(r'var\s+nickname\s*=\s*htmlDecode\("([^"]+)"\)', content)
if nm:
    account = nm.group(1).strip()
if not account:
    nm = re.search(r"nick_name\s*[:=]\s*['\"]([^'\"]+)['\"]", content)
    if nm:
        account = nm.group(1).strip()
if not account:
    account = "未知公众号"

# ── 发布时间（完整到分钟） ──
pub_date = ""   # YYYYMMDD
pub_time = ""   # YYYYMMDDHHMM

def ts_to_str(ts, fmt):
    try:
        return datetime.datetime.fromtimestamp(int(ts)).strftime(fmt)
    except:
        return ""

# 优先级1: og:article:publish_time meta 标签
tm = re.search(r'<meta\s+property="og:article:publish_time"\s+content="([^"]+)"', content)
if tm:
    pub_date = ts_to_str(tm.group(1), '%Y%m%d')
    pub_time = ts_to_str(tm.group(1), '%Y%m%d%H%M')

# 优先级2: create_time JS 变量（Unix 秒级时间戳）
if not pub_date:
    cm = re.search(r"create_time\s*[:=]\s*['\"]?(\d{10})", content)
    if cm:
        pub_date = ts_to_str(cm.group(1), '%Y%m%d')
        pub_time = ts_to_str(cm.group(1), '%Y%m%d%H%M')

# 优先级3: 全文匹配 "20XX年X月X日 HH:MM"
if not pub_date:
    dm = re.search(r'(20\d{2})年(\d{1,2})月(\d{1,2})日\s*(\d{1,2}):(\d{2})', content)
    if dm:
        pub_date = f"{dm.group(1)}{int(dm.group(2)):02d}{int(dm.group(3)):02d}"
        pub_time = f"{dm.group(1)}{int(dm.group(2)):02d}{int(dm.group(3)):02d}{int(dm.group(4)):02d}{dm.group(5)}"
    else:
        dm = re.search(r'(20\d{2})年(\d{1,2})月(\d{1,2})日', content)
        if dm:
            pub_date = f"{dm.group(1)}{int(dm.group(2)):02d}{int(dm.group(3)):02d}"
            pub_time = pub_date + "0000"

# 兜底
if not pub_date:
    pub_date = datetime.date.today().strftime('%Y%m%d')
    pub_time = pub_date + "0000"

# ── 提取正文 ──
body_m = re.search(
    r'<div class="rich_media_content[^"]*"[^>]*>(.*?)</div>\s*</div>\s*</div>\s*<div[^>]*id="js_pc_qr_code"',
    content, re.DOTALL
)
if not body_m:
    body_m = re.search(r'<div id="js_content"[^>]*>(.*?)</div>\s*<div[^>]*id="js_pc_qr_code"', content, re.DOTALL)
if not body_m:
    body_m = re.search(r'<div class="rich_media_content[^"]*"[^>]*>(.*?)</div>', content, re.DOTALL)

if body_m:
    body = body_m.group(1)
    body = re.sub(r'<br\s*/?>', '\n', body)
    body = re.sub(r'</?p[^>]*>', '\n\n', body)
    body = re.sub(r'</?section[^>]*>', '\n', body)
    body = re.sub(r'</?span[^>]*>', '', body)
    body = re.sub(r'</?div[^>]*>', '\n', body)
    body = re.sub(r'<img[^>]*>', '', body)
    body = re.sub(r'<[^>]+>', '', body)
    body = html_module.unescape(body)
    body = re.sub(r'\n{3,}', '\n\n', body)
    body = re.sub(r' {2,}', ' ', body)

    # ── 清理微信页面残留的 JS / UI 碎片 ──
    body = re.sub(r"var\s+first_sceen__time\s*=\s*\(\+new\s+Date\(\)\).*?(?:\n|$)", '', body)
    body = re.sub(r"if\s*\(\s*[\"'][\s\"]*==\s*1[^)]*getElementById[^}]*\}", '', body)
    body = re.sub(r'document\.getElementById[^;]*;', '', body)
    body = re.sub(r'[^\n]*?addEventListener\("selectstart"[^;]*;[^\n]*', '', body)
    body = re.sub(r'[^\n]*?e\.preventDefault\(\)[^\n]*', '', body)
    body = re.sub(r'预览时标签不可点', '', body)
    body = re.sub(r'阅读原文', '', body)
    body = re.sub(r'<script[^>]*>.*?</script>', '', body, flags=re.DOTALL)
    # 清理独立的 JS 符号残留行（如 ");" "}" ）
    body = re.sub(r'^\s*[\);\}]+\s*$', '', body, flags=re.MULTILINE)
    # 清理空行
    body = re.sub(r'\n{3,}', '\n\n', body)
    # 去除每行首尾空白，删除纯空白行
    body = '\n'.join(line.strip() for line in body.split('\n') if line.strip())
    body = body.strip()
else:
    body = "⚠️ 未提取到正文，请手动查看原始 HTML。"

# ── 生成文件名: YYYYMMDDHHMM-作者-公众号-标题.md ──
def safe_name(s):
    s = re.sub(r'[\\/*?:"<>|]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

sa = safe_name(author)
sg = safe_name(account)
st = safe_name(title)
if len(st) > 50:
    st = st[:50]

filename = f"{pub_date}-{sa}-{sg}-{st}.md"

output_dir = os.path.expanduser("~/Desktop/公众号文件暂存")
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, filename)

# ── 构造 Markdown ──
md = f"# {title}\n"
md += f"**作者：** {author}\n"
md += f"**公众号：** {account}\n"
md += f"**发布时间：** {pub_date}\n"
md += f"**原始链接：** {url}\n\n---\n\n{body}\n"

with open(output_path, 'w', encoding='utf-8') as f:
    f.write(md)

print(f"文件: {os.path.basename(output_path)}")
print(f"作者: {author}")
print(f"公众号: {account}")
print(f"日期: {pub_date}")
print(f"字数: {len(body)}")
PYEOF
```

### Step 4 — 反馈结果

```
✅ 已保存
────────────────────
文件：202307270730-闵佳凤-勤思知律-字体侵权应如何处理？.md
位置：~/Desktop/公众号文件暂存/
作者：闵佳凤 | 公众号：勤思知律
字数：2,667 | 日期：2023年7月27日 07:30
────────────────────
```

用 `open` 打开文件供查看。如果有旧文件（同文章但命名不同），自动删除。

## 错误处理

| 异常 | 处理 |
|------|------|
| curl 返回空 | 提示用户确认链接是否有效、是否被删除 |
| 文件 < 10KB | 可能被微信拦截，提示用户在浏览器中打开确认 |
| 提取不到标题 | "未命名文章" |
| 提取不到作者 | "未知作者" |
| 提取不到公众号 | "未知公众号" |
| 提取不到日期 | 当天日期，时间填 0000 |
| 提取不到正文 | 保存含错误提示的 md |

## 依赖

- `curl`（macOS 自带）
- Python 3（macOS 自带）
- 无需安装任何第三方包
