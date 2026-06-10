---
name: alex-wechat-classifier
description: >
  Use when the user wants to classify WeChat articles by keyword — scan markdown files
  in 公众号文件暂存, group matching ones into keyword subfolders, and generate
  a formatted Excel summary. Triggers on: 分类, 关键词分类, 按XX归类,
  整理公众号文章, 汇总文章.
argument-hint: "<关键词>"
---

# /wechat-article-classifier

## 概述

扫描 `~/Desktop/公众号文件暂存/` 下的 `.md` 文件，按关键词匹配标题或正文，将匹配的文章移动到关键词子文件夹（母目录不再保留），并生成格式化 Excel 汇总表。

- 子文件夹：`~/Desktop/公众号文件暂存/{关键词}/`
- Excel：`{关键词}公众号文章汇总.xlsx`

## 工作流程

### Step 1 — 确认关键词

从用户消息中提取关键词。如果用户没给，提示输入。

### Step 2 — 扫描并匹配

```bash
python3 << 'PYEOF'
import re, os, glob

keyword = "用户提供的关键词"
source_dir = os.path.expanduser("~/Desktop/公众号文件暂存")

md_files = glob.glob(os.path.join(source_dir, "*.md"))
matched = []

for f in md_files:
    with open(f, 'r', encoding='utf-8') as fh:
        text = fh.read()
    # 匹配标题（# 行）或正文中的关键词
    title_match = bool(re.search(keyword, text.split('\n')[0] if text else ''))
    body_match = bool(re.search(keyword, text))
    if title_match or body_match:
        matched.append(f)

print(f"扫描: {len(md_files)} 篇 | 匹配: {len(matched)} 篇")
for f in matched:
    print(f"  {os.path.basename(f)}")
PYEOF
```

向用户展示匹配结果，确认是否继续。

### Step 3 — 复制文件并生成 Excel

```bash
python3 << 'PYEOF'
import re, os, glob, shutil
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

keyword = "用户提供的关键词"
source_dir = os.path.expanduser("~/Desktop/公众号文件暂存")
target_dir = os.path.join(source_dir, keyword)

# ── 扫描匹配 ──
md_files = glob.glob(os.path.join(source_dir, "*.md"))
matched = []
for f in md_files:
    with open(f, 'r', encoding='utf-8') as fh:
        text = fh.read()
    if re.search(keyword, text):
        matched.append(f)

if not matched:
    print(f"⚠️ 没有找到包含「{keyword}」的文章")
    exit(0)

# ── 移动文件到子文件夹 ──
is_new = not os.path.exists(target_dir)
os.makedirs(target_dir, exist_ok=True)
for f in matched:
    shutil.move(f, target_dir)
    print(f"📋 移动: {os.path.basename(f)}")

if is_new:
    print(f"📁 新建分类: {keyword}/")
else:
    print(f"📁 纳入已有分类: {keyword}/")

# ── 解析元数据（扫描子文件夹内所有 md，覆盖已有 + 新增） ──
all_md = glob.glob(os.path.join(target_dir, "*.md"))
articles = []
seen_urls = set()
for f in all_md:
    with open(f, 'r', encoding='utf-8') as fh:
        text = fh.read()
    
    # 按内容模式提取，不依赖行号
    title = ""
    tm = re.search(r'^#\s+(.+?)$', text, re.MULTILINE)
    if tm: title = tm.group(1).strip()
    if not title: title = os.path.basename(f)
    
    def extract_field(label):
        m = re.search(rf'\*\*{label}[：:]\*\*\s*(.+)', text)
        return m.group(1).strip() if m else ""
    
    author = extract_field('作者')
    account = extract_field('公众号')
    pub_date = extract_field('发布时间')
    url = extract_field('原始链接')
    
    # 去重：同链接只保留一份
    if url and url in seen_urls:
        continue
    if url:
        seen_urls.add(url)
    
    # 字数统计（跳过元数据头，从 --- 分隔线之后开始）
    body_start = text.find('---\n')
    if body_start > 0:
        body_text = text[body_start+4:]
    else:
        body_text = text
    char_count = len(body_text.replace('\n', '').replace(' ', ''))
    
    # 日期格式化: 20230727 → 2023年7月27日
    display_date = pub_date
    dm = re.match(r'(\d{4})(\d{2})(\d{2})', pub_date)
    if dm:
        display_date = f"{dm.group(1)}年{int(dm.group(2))}月{int(dm.group(3))}日"
    
    articles.append({
        'title': title,
        'date': display_date,
        'author': author,
        'account': account,
        'chars': char_count,
        'url': url,
    })

# ── 按发布时间排序 ──
articles.sort(key=lambda a: a['date'])

# ── 生成 Excel ──
wb = Workbook()
ws = wb.active
ws.title = f"{keyword}文章汇总"

headers = ['序号', '发布时间', '文章名称', '作者', '公众号名称', '文章字符数量', '原文链接']
col_widths = [6, 18, 45, 14, 18, 14, 50]

# 样式
HEADER_FONT = Font(name='黑体', size=11, bold=True)
BODY_FONT = Font(name='宋体', size=10)
HEADER_FILL = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

# 写表头
for col, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=h)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = CENTER
    cell.border = THIN_BORDER

# 写数据
for i, a in enumerate(articles, 1):
    vals = [i, a['date'], a['title'], a['author'], a['account'], a['chars'], a['url']]
    for col, val in enumerate(vals, 1):
        cell = ws.cell(row=i+1, column=col, value=val)
        cell.font = BODY_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER
    ws.row_dimensions[i+1].height = 40

# 列宽
for col, w in enumerate(col_widths, 1):
    from openpyxl.utils import get_column_letter
    ws.column_dimensions[get_column_letter(col)].width = w

ws.freeze_panes = 'A2'

excel_path = os.path.join(target_dir, f"{keyword}公众号文章汇总.xlsx")
wb.save(excel_path)

print(f"\n✅ 分类{'更新' if not is_new else ''}完成")
print(f"📁 文件夹: {target_dir}")
print(f"📄 文章: {len(articles)} 篇{' (含既有)' if not is_new else ''}")
print(f"📊 Excel: {excel_path}")
PYEOF
```

### Step 4 — 反馈结果

```
✅ 分类完成
────────────────────
关键词：字体
文件夹：~/Desktop/公众号文件暂存/字体/
文章：3 篇
Excel：字体公众号文章汇总.xlsx
────────────────────
  1. 2023年7月27日 | 闵佳凤 | 勤思知律
  2. 2024年7月20日 | 李士林 | 知产前沿
  3. 2025年2月10日 | 未知 | 税律周
```

用 `open` 打开目标文件夹。

## 依赖

- Python 3 + `openpyxl`（需预装：`pip3 install openpyxl`）
