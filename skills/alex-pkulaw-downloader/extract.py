#!/usr/bin/env python3
"""Extract PKULaw judgment from saved HTML or CDP-fetched HTML."""
import re, os, sys, html as html_module

def extract(html, url="", output_dir=None):
    if output_dir is None:
        output_dir = os.path.expanduser("~/Desktop/北大法宝判例")
    os.makedirs(output_dir, exist_ok=True)

    # ── 1. divFullText ──
    idx = html.find('id="divFullText"')
    if idx < 0:
        return None, "divFullText not found"
    ts = html.rfind('<div', idx - 100, idx)
    te = html.find('>', idx) + 1
    depth, pos = 1, te
    while depth > 0:
        no = html.find('<div', pos)
        nc = html.find('</div>', pos)
        if nc < 0:
            break
        if no > 0 and no < nc:
            depth += 1
            pos = no + 4
        else:
            depth -= 1
            pos = nc + 6

    jt_raw = html[te:pos]
    jt = re.sub(r'<br\s*/?>', '\n', jt_raw)
    jt = re.sub(r'</?p[^>]*>', '\n\n', jt)
    jt = re.sub(r'</?div[^>]*>', '\n', jt)
    jt = re.sub(r'<[^>]+>', '', jt)
    jt = html_module.unescape(jt)
    jt = re.sub(r'[ \t]+', ' ', jt)
    jt = re.sub(r'\n{3,}', '\n\n', jt)
    jt = re.sub(r'^[ \t]+', '', jt, flags=re.MULTILINE)
    jt = jt.strip()

    # ── 2. body text (metadata) ──
    body = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL)
    body = re.sub(r'<style[^>]*>.*?</style>', '', body, flags=re.DOTALL)
    body = re.sub(r'<br\s*/?>', '\n', body)
    body = re.sub(r'<[^>]+>', '', body)
    body = html_module.unescape(body)

    # ── 3. metadata with field boundaries ──
    FLDS = r'案由|案\s*号|审理法官|文书类型|公开类型|审理法院|审结日期|案件类型|审理程序|权责关键词|相关企业'

    def gf(label):
        m = re.search(rf'({label})[：:]\s*(.*?)(?=\n\s*(?:{FLDS})[：:]|\n\s*\n|\Z)', body, re.DOTALL)
        if m:
            return re.sub(r'\s+', ' ', m.group(2).strip())
        return ""

    cn = gf(r'案\s*号')
    court = gf('审理法院')
    date_raw = gf('审结日期')
    dm = re.match(r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})', date_raw)
    date_str = f"{dm.group(1)}年{int(dm.group(2))}月{int(dm.group(3))}日" if dm else date_raw
    procedure = gf('审理程序')
    doc_type = gf('文书类型')
    case_cause = re.sub(r'\s+案\s*号.*$', '', gf('案由')).strip()

    # ── 4. title ──
    title = ""
    tm = re.search(r'<title>(.+?)(?:-北大法宝|</title>)', html)
    if tm:
        title = html_module.unescape(tm.group(1).strip())
    if not title:
        title = cn

    # ── 5. 裁判规则 ──
    rule_text = ""
    rm = re.search(r'裁判规则\s*\n(.*?)(?=\n(?:中国裁判文书网|复制全文|\Z))', body, re.DOTALL)
    if rm:
        rule_text = rm.group(1).strip()

    # ── 6. build Markdown ──
    L = [f"# {title}", ""]
    for k, v in [
        ('案号', cn), ('审理法院', court), ('审结日期', date_str),
        ('审理程序', procedure), ('文书类型', doc_type), ('案由', case_cause),
    ]:
        L.append(f"**{k}：** {v}")
    L += ["", f"**来源：** 北大法宝", f"**原始链接：** {url}", "", "---", ""]

    if rule_text:
        L += ["## 裁判规则", ""]
        for sub in ['关键词', '核心问题', '裁判要点']:
            sm = re.search(
                rf'{sub}[：:]\s*(.+?)(?=\n\s*(?:关键词|核心问题|裁判要点|\Z))',
                rule_text, re.DOTALL
            )
            if sm:
                val = re.sub(r'\s+', ' ', sm.group(1).strip())
                L += [f"**{sub}：** {val}", ""]
        L += ["---", ""]

    if jt:
        L += ["## 判决书全文", ""]
        sections = [
            '当事人', '审理经过', '原告诉称', '被告辩称', '一审法院查明',
            '一审法院认为', '二审被上诉人辩称', '本院查明', '二审上诉人诉称',
            '本院认为', '裁判结果', '落款',
        ]
        # Check if structured judgment: ≥3 markers at LINE START (not embedded)
        found_count = sum(1 for s in sections if re.search(rf'(?:^|\n){s}', jt))
        if found_count >= 3:
            last = 0
            for i, sec in enumerate(sections):
                pos = jt.find(sec, last)
                if pos < 0:
                    continue
                nxt = len(jt)
                for ns in sections[i + 1:]:
                    np = jt.find(ns, pos + len(sec))
                    if np > 0 and np < nxt:
                        nxt = np
                content = jt[pos:nxt].strip()
                if len(content) > len(sec) + 5:
                    L += [f"### {sec}", "", content, ""]
                last = nxt
        else:
            # Raw judgment without structured sections — output as-is
            L += [jt]

    md = '\n'.join(L)

    # ── 7. save ──
    def sn(s):
        return re.sub(r'[\\/*?:"<>|]', '', s).strip()[:60]

    cn_s = sn(cn) if cn else "未知案号"
    t_s = sn(title)[:40] if title else "未知"
    filename = f"{cn_s} {t_s}.md".replace('  ', ' ')
    filepath = os.path.join(output_dir, filename)

    # remove old versions with same case number
    for f in os.listdir(output_dir):
        if cn_s in f and f != filename:
            os.remove(os.path.join(output_dir, f))

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(md)

    return filepath, f"{len(md)} chars, judgment {len(jt)} chars"


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: extract.py <html_file> [url]")
        sys.exit(1)
    html_file = sys.argv[1]
    url = sys.argv[2] if len(sys.argv) > 2 else ""
    with open(html_file, 'r', encoding='utf-8') as f:
        html = f.read()
    result, info = extract(html, url)
    print(f"✅ {result}")
    print(f"📊 {info}")
