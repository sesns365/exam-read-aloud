"""把考卷（PDF／Word／圖片）轉成「版面完全不變、可點按報讀」的單一 HTML 檔。

用法：
  python -I build.py <輸入檔...> -o <輸出.html> --work <工作資料夾> [--title 標題] [--dump 報讀文字.txt] [--dpi 200]

輸入：
  .pdf                         有文字層的頁面直接用 PDF 文字；沒有文字層（掃描）的頁面自動 OCR
  .docx .doc .rtf .odt         先用 Word 轉成 PDF（版面不變）再處理
  .png .jpg .jpeg .bmp .tif .tiff .webp .gif   每張圖一頁，用 Windows 內建 OCR（zh-Hant-TW）
  多個檔案依指令中的順序接成多頁。

OCR 校正流程（圖片、掃描檔）：
  第一次執行會在 <工作資料夾>/ocr_lines.json 寫出每一行的位置和辨識文字（text 欄位），並先用原始辨識結果產生草稿 HTML。
  對照頁面圖片（<工作資料夾>/page_N.png）把 text 改成正確文字後，用同樣指令再執行一次，就會用校正後的文字產生正式 HTML。
  - 只改 text，不要改 id、bbox。
  - 圈號選項要寫成 ①②③④（或 ○１、○２…），代號選項寫成「A 東邊 B 西邊」（字母前後留空格）。
  - 閱讀文章新段落的開頭加兩個半形空白。
  - 整行刪除：text 設為 ""。辨識漏掉的整行：自行新增一筆 {"id": "p1-new1", "page": 1, "bbox": [x0,y0,x1,y1], "text": "..."}（像素座標），放在閱讀順序正確的位置。
  - 陣列的順序就是閱讀順序（同一題的行要相鄰），必要時可以搬動整筆的位置。

可點區塊：
  藍框（題號、整題範圍）＝ 整題一次唸完（題目＋全部選項）
  橘框（題目文字）       ＝ 只唸題目
  橘框（每個選項）       ＝ 只唸該選項（①②③④、○＋全形數字、代號框 A–D）
"""
import pymupdf, sys, json, re, base64, argparse, os, subprocess, difflib, io

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')
HERE = os.path.dirname(os.path.abspath(__file__))
ap = argparse.ArgumentParser()
ap.add_argument('inputs', nargs='+')
ap.add_argument('-o', '--out', required=True)
ap.add_argument('--work', required=True)
ap.add_argument('--title', default=None)
ap.add_argument('--dump', default=None)
ap.add_argument('--dpi', type=int, default=200)
ap.add_argument('--max-px', type=int, default=2400, help='圖片最長邊超過此值時縮小')
ap.add_argument('--template', default=os.path.join(HERE, '..', 'assets', 'template.html'))
args = ap.parse_args()
os.makedirs(args.work, exist_ok=True)

IMG_EXT = ('.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp', '.gif')
WORD_EXT = ('.docx', '.doc', '.rtf', '.odt')
FW = {'１': 1, '２': 2, '３': 3, '４': 4, '５': 5}
CIRC = {'①': 1, '②': 2, '③': 3, '④': 4, '⑤': 5}
LETTERS = 'ABCDEFGHㄅㄆㄇㄈㄉＡＢＣＤＥＦＧＨ'   # 代號：英文字母、注音符號或全形字母（同一行要有兩個以上才算代號選項）
CIRC_ALL = '①②③④⑤⑥⑦⑧⑨⑩'
CJK = lambda ch: '一' <= ch <= '鿿'
SECTION_RE = re.compile(r'^[一二三四五六七八九十]+[、．.]')
BR = r'[（(]\s*[）)]'   # 作答括號 (   )
Q_RE = re.compile(rf'^(\d+\s*[.．、]|{BR}\s*\d+\s*[.．、])')
SUB_RE = re.compile(rf'^({BR}\s*([（(][１２３４５６７８９1-9][）)]|[{CIRC_ALL}])|[（(][１２３４５６７８９1-9][）)])')
PREFIX_RE = re.compile(rf'^\s*(\d+\s*[.．、]\s*{BR}|{BR}\s*\d+\s*[.．、]|\d+\s*[.．、]|{BR}\s*[（(].[）)]|{BR}\s*[{CIRC_ALL}]|[（(][１２３４５６７８９1-9][）)])\s*')


def cl(s):
    """去掉空白；但英文單字之間保留一個空格（Made in Taiwan）。"""
    s = re.sub(r'\s+', ' ', s).strip()
    return re.sub(r'(?<![A-Za-z]) | (?![A-Za-z])', '', s)


def strip_zhuyin(s):
    """去掉緊接在國字後面、以文字輸入的注音（例如「上ㄕㄤˋ課ㄎㄜˋ」→「上課」），避免被唸出來。"""
    # 只刪「兩個以上注音符號」或「注音＋聲調」；單獨一個注音（ㄅ、ㄆ 代號）保留
    Z = r'[ㄅ-ㄯㆠ-ㆿ]'
    s = re.sub(rf'(?<=[一-鿿])\s*({Z}{{2,}}[ˊˇˋ˙]?|{Z}[ˊˇˋ˙])', '', s)
    return re.sub(r'(?<=[一-鿿])˙', '', s)


def union(bbs):
    return [min(b[0] for b in bbs), min(b[1] for b in bbs), max(b[2] for b in bbs), max(b[3] for b in bbs)]


def digit(c):
    return FW.get(c) or (int(c) if c.isdigit() else c)


def encode_img(png_bytes_or_pix):
    """回傳 data URL（PNG 與 JPEG 取較小者）。"""
    if isinstance(png_bytes_or_pix, pymupdf.Pixmap):
        png = png_bytes_or_pix.tobytes('png'); jpg = png_bytes_or_pix.tobytes('jpg', jpg_quality=82)
    else:
        from PIL import Image
        im = png_bytes_or_pix
        b1 = io.BytesIO(); im.save(b1, 'PNG', optimize=True); png = b1.getvalue()
        b2 = io.BytesIO(); im.convert('RGB').save(b2, 'JPEG', quality=82); jpg = b2.getvalue()
    data, mime = (png, 'image/png') if len(png) < len(jpg) else (jpg, 'image/jpeg')
    return f'data:{mime};base64,' + base64.b64encode(data).decode(), len(data)


# ---------- 1. 收集頁面 ----------
pages_src = []   # 每頁：{"kind": "text"|"ocr", "img": dataurl, "w","h", "lines"(text) | "png"(ocr)}
for src in args.inputs:
    ext = os.path.splitext(src)[1].lower()
    if ext in WORD_EXT:
        pdf = os.path.join(args.work, f'word_{len(pages_src)}.pdf')
        print(f'Word 轉 PDF：{src}')
        subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
                        os.path.join(HERE, 'word2pdf.ps1'), src, pdf], check=True)
        if not os.path.exists(pdf):
            sys.exit(f'Word 轉 PDF 失敗：{src}')
        src, ext = pdf, '.pdf'
    if ext == '.pdf':
        doc = pymupdf.open(src)
        for p in doc:
            lines = []
            for b in p.get_text('rawdict')['blocks']:
                if b['type'] != 0:
                    continue
                for l in b['lines']:
                    chars = [(c['c'], c['bbox']) for s in l['spans'] for c in s['chars']]
                    lines.append({'text': ''.join(c for c, _ in chars), 'chars': chars, 'bbox': l['bbox']})
            n_text = sum(len(cl(l['text'])) for l in lines)
            if n_text >= 20:
                pix = p.get_pixmap(dpi=args.dpi, colorspace=pymupdf.csGRAY)
                url, size = encode_img(pix)
                pages_src.append({'kind': 'text', 'img': url, 'size': size, 'w': p.rect.width, 'h': p.rect.height, 'lines': lines})
            else:  # 掃描頁：轉圖後 OCR
                pix = p.get_pixmap(dpi=args.dpi)
                n = len(pages_src) + 1
                path = os.path.join(args.work, f'page_{n}.png'); pix.save(path)
                pages_src.append({'kind': 'ocr', 'png': path})
    elif ext in IMG_EXT:
        from PIL import Image, ImageOps
        im = ImageOps.exif_transpose(Image.open(src))
        if im.mode not in ('RGB', 'L'):
            im = im.convert('RGB')
        scale = args.max_px / max(im.size)
        if scale < 1:
            im = im.resize((round(im.width * scale), round(im.height * scale)), Image.LANCZOS)
        n = len(pages_src) + 1
        path = os.path.join(args.work, f'page_{n}.png'); im.save(path)
        pages_src.append({'kind': 'ocr', 'png': path})
    else:
        sys.exit(f'不支援的檔案類型：{src}')

# ---------- 2. OCR（含校正檔） ----------
ocr_lines_path = os.path.join(args.work, 'ocr_lines.json')
ocr_boxes_path = os.path.join(args.work, 'ocr_boxes.json')


def run_ocr(png):
    out = png + '.ocr.json'
    subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
                    os.path.join(HERE, 'ocr.ps1'), png, out], check=True)
    return json.load(open(out, encoding='utf-8-sig'))


def ocr_to_lines(d, page_no):
    """Windows OCR 結果 → 依閱讀順序排好的行（每字含位置）。"""
    raw = []
    for words in d['lines']:
        if isinstance(words, dict):
            words = [words]
        chars = []
        prev = None
        for w in words:
            x, y, ww, hh = w['x'], w['y'], w['w'], w['h']
            if prev is not None and x - prev[2] > 0.45 * hh:
                chars.append((' ', [prev[2], y, x, y + hh]))
            t = w['t']; cw = ww / max(1, len(t))
            for k, c in enumerate(t):
                chars.append((c, [x + k * cw, y, x + (k + 1) * cw, y + hh]))
            prev = [x, y, x + ww, y + hh]
        if chars:
            raw.append({'chars': chars, 'bbox': union([b for _, b in chars])})
    W = d['width']
    # 分欄：若有相當比例的行從右半邊開始，視為左右兩欄
    right = [l for l in raw if l['bbox'][0] > 0.45 * W]
    two_col = len(right) >= 0.2 * len(raw)

    def col(l):
        if not two_col:
            return 0
        if l['bbox'][2] - l['bbox'][0] > 0.6 * W:
            return -1  # 橫跨兩欄（標題）
        return 0 if (l['bbox'][0] + l['bbox'][2]) / 2 < W / 2 else 1
    out = []
    for c in sorted(set(col(l) for l in raw)):
        ls = sorted([l for l in raw if col(l) == c], key=lambda l: l['bbox'][1])
        rows = []
        for l in ls:  # 同一列（垂直重疊過半）
            if rows:
                r = rows[-1][-1]['bbox']; b = l['bbox']
                ov = min(r[3], b[3]) - max(r[1], b[1])
                if ov > 0.5 * min(r[3] - r[1], b[3] - b[1]):
                    rows[-1].append(l); continue
            rows.append([l])
        for row in rows:
            row.sort(key=lambda l: l['bbox'][0])
            merged = [row[0]]
            for l in row[1:]:
                m = merged[-1]; h = m['bbox'][3] - m['bbox'][1]
                if l['bbox'][0] - m['bbox'][2] < 3 * h:
                    gap = [m['bbox'][2], m['bbox'][1], l['bbox'][0], m['bbox'][3]]
                    m['chars'] = m['chars'] + [(' ', gap)] + l['chars']
                    m['bbox'] = union([m['bbox'], l['bbox']])
                else:
                    merged.append(l)
            out.extend(merged)
    for i, l in enumerate(out):
        l['id'] = f'p{page_no}-{i+1:03d}'
        text = ''.join(c for c, _ in l['chars'])
        l['text'] = re.sub(r'(?<=[^\x00-\x7f]) (?=[^\x00-\x7f])', '', text)
    return out


ocr_pages = [i for i, p in enumerate(pages_src) if p['kind'] == 'ocr']
corrected = None
if ocr_pages:
    if os.path.exists(ocr_lines_path) and os.path.exists(ocr_boxes_path):
        corrected = json.load(open(ocr_lines_path, encoding='utf-8'))
        boxes = json.load(open(ocr_boxes_path, encoding='utf-8'))
        print('使用已校正的 ocr_lines.json')
    else:
        entries, boxes = [], {}
        for i in ocr_pages:
            print(f'第 {i+1} 頁 OCR 辨識中…')
            d = run_ocr(pages_src[i]['png'])
            for l in ocr_to_lines(d, i + 1):
                entries.append({'id': l['id'], 'page': i + 1, 'bbox': [round(v, 1) for v in l['bbox']], 'text': l['text']})
                boxes[l['id']] = [[c] + [round(v, 1) for v in b] for c, b in l['chars']]
        with open(ocr_lines_path, 'w', encoding='utf-8') as f:
            f.write('[\n' + ',\n'.join(json.dumps(e, ensure_ascii=False) for e in entries) + '\n]\n')
        json.dump(boxes, open(ocr_boxes_path, 'w', encoding='utf-8'), ensure_ascii=False)
        corrected = entries
        print(f'\n★ OCR 草稿已輸出：{ocr_lines_path}\n  請對照 page_N.png 校正 text 後，用同樣指令再執行一次。\n')


def align(entry, ocr_chars):
    """把校正後文字對齊到 OCR 字框；對不上的字用前後字框內插。"""
    text = entry['text']
    x0, y0, x1, y1 = entry['bbox']
    h = y1 - y0
    src = [c for c, *_ in ocr_chars]
    res = [None] * len(text)
    sm = difflib.SequenceMatcher(None, src, list(text), autojunk=False)
    # 辨識結果太亂（相似度低）時不採用 OCR 字框，改在整行範圍內依字寬平均分配
    ops = sm.get_opcodes() if src and sm.ratio() >= 0.4 else []
    for tag, i1, i2, j1, j2 in ops:
        if tag == 'equal' or (tag == 'replace' and i2 - i1 == j2 - j1):
            for k in range(j2 - j1):
                res[j1 + k] = list(ocr_chars[i1 + k][1:])
    cw = lambda c: (0.3 if c == ' ' else 0.55 if ord(c) < 0x2000 else 1.0) * h
    j = 0
    while j < len(text):
        if res[j] is not None:
            j += 1; continue
        k = j
        while k < len(text) and res[k] is None:
            k += 1
        need = sum(cw(c) for c in text[j:k])
        left = res[j - 1][2] if j > 0 else None
        right = res[k][0] if k < len(text) else None
        if left is None and right is None:
            left, right = x0, max(x1, x0 + need)
        elif left is None:
            left = max(x0 if x0 < right else right - need, right - need)
        elif right is None:
            right = left + need
        span = max(right - left, 1)
        x = left
        for t in range(j, k):
            wdt = cw(text[t]) / need * span if need else 0
            res[t] = [x, y0, x + wdt, y1]; x += wdt
        j = k
    return [(c, b) for c, b in zip(text, res)]


if ocr_pages:
    for i in ocr_pages:
        pg = pages_src[i]
        from PIL import Image
        im = Image.open(pg['png'])
        pg['w'], pg['h'] = im.size
        pg['img'], pg['size'] = encode_img(im)
        pg['lines'] = []
        for e in corrected:
            if e.get('page') != i + 1 or not e.get('text', '').strip():
                continue
            chars = align(e, boxes.get(e['id'], []))
            pg['lines'].append({'text': e['text'], 'chars': chars, 'bbox': e['bbox']})

# ---------- 3. 切題、產生可點區塊 ----------
pages, dump = [], []
seen_section = False
for pi, pg in enumerate(pages_src):
    W, H, lines = pg['w'], pg['h'], pg['lines']
    blocks, cur = [], None
    for ln in lines:
        t = ln['text']; s = t.strip()
        if not s:
            continue
        if re.fullmatch(r'[-－\s]*\d+[-－\s]*', s) and ln['bbox'][1] > H * 0.96:
            continue  # 頁碼
        if re.search(r'座號.*姓名|班.{0,6}座號|^姓名[:：]', s) or (re.search(r'座號|姓名', s) and not seen_section):
            continue  # 班級座號姓名欄
        kind = None
        if s == '○':
            pass
        elif SECTION_RE.match(s):
            kind = 'section'; seen_section = True
        elif not seen_section and pi == 0:
            kind = 'title' if (cur is None or cur['kind'] != 'title') else None
        elif Q_RE.match(s):
            kind = 'q'
        elif SUB_RE.match(s):
            kind = 'sub'
        elif (t.startswith('  ') or t.startswith('　')) and not s.startswith(('○', '①', '②', '③', '④')) and cur and cur['kind'] in ('q', 'para', 'section'):
            kind = 'para'  # 閱讀題文章段落（行首縮排）
        elif s.startswith(('＊', '*', '※')):
            kind = 'note'
        if kind or cur is None:
            cur = {'kind': kind or 'misc', 'lines': []}
            blocks.append(cur)
        cur['lines'].append(ln)

    out_blocks = []
    for bl in blocks:
        seq = []
        for li, ln in enumerate(bl['lines']):
            for c, bb in ln['chars']:
                seq.append([c, bb, li, 0])
        # 開頭題號（例如「(   ) ①」的 ①）不算選項
        pm = PREFIX_RE.match(''.join(x[0] for x in seq))
        nprefix = len(pm.group(0)) if pm else 0
        # 代號選項（A／B 或 ㄅ／ㄆ）：字母在行首或空白後、後面接中文，且同一行至少兩個
        code_at = set()
        for li in range(len(bl['lines'])):
            idx = [i for i, x in enumerate(seq) if x[2] == li]
            cand = []
            for n, i in enumerate(idx):
                c = seq[i][0]
                # 注音代號不會出現在一般文字中，前面不必是空白
                if c in LETTERS and (n == 0 or seq[idx[n-1]][0].isspace() or c in 'ㄅㄆㄇㄈㄉＡＢＣＤＥＦＧＨ'):
                    nxt = next((seq[j][0] for j in idx[n+1:] if not seq[j][0].isspace() and seq[j][0] not in '.．、'), '')
                    if CJK(nxt):
                        cand.append(i)
            if len({seq[i][0] for i in cand}) >= 2:
                code_at.update(cand)
        owner, pending = 0, []
        for i, item in enumerate(seq):
            c = item[0]
            if i < nprefix:
                item[3] = 0; continue
            if c.isspace():
                item[3] = owner; continue
            if c == '○':
                pending.append(i); continue
            new = None
            if c in CIRC:
                new = CIRC[c]
            elif pending and c in FW:
                new = FW[c]
            elif pending and c in '12345':   # ○1 ○2（○加半形數字）
                new = int(c)
            elif pending and c in 'ㄅㄆㄇㄈㄉ':   # ○ㄅ ○ㄆ（圈注音代號）
                new = 10 + LETTERS.index(c)
            elif i in code_at:
                new = 10 + LETTERS.index(c)
            if new is not None:
                owner = new
            for j in pending:
                seq[j][3] = owner
            pending = []
            item[3] = owner
        for j in pending:
            seq[j][3] = owner

        owners = sorted(set(x[3] for x in seq))
        parts = {}
        for o in owners:
            txt = ''.join(x[0] for x in seq if x[3] == o)
            rects = []
            for li in range(len(bl['lines'])):
                bbs = [x[1] for x in seq if x[3] == o and x[2] == li and not x[0].isspace()]
                if bbs:
                    rects.append(union(bbs))
            parts[o] = {'text': txt, 'rects': rects}

        st = parts.get(0, {'text': ''})['text'].strip()
        st = re.sub(rf'^(\d+)\s*[.．、]\s*{BR}', r'第\1題，', st)
        st = re.sub(rf'^{BR}\s*(\d+)\s*[.．、]', r'第\1題，', st)
        st = re.sub(r'^(\d+)\s*[.．、]', r'第\1題，', st)
        st = re.sub(rf'^{BR}\s*[（(](.)[）)]', lambda m: f'第{digit(m.group(1))}小題，', st)
        st = re.sub(rf'^{BR}\s*([{CIRC_ALL}])', lambda m: f'第{CIRC_ALL.index(m.group(1)) + 1}小題，', st)
        st = re.sub(r'^[（(]([１２３４５６７８９1-9])[）)]', lambda m: f'第{digit(m.group(1))}小題，', st)
        st = re.sub(r'[（(]\s*[_＿]+\s*[）)]', '空格', st)
        st = re.sub(BR, '括號', st)
        st = strip_zhuyin(st)
        st = cl(st).replace('○', '圈圈').replace('ˇ', '勾').replace('✓', '勾').replace('×', '叉').replace('＊', '').replace('※', '')
        st = re.sub(r'_{2,}|＿{2,}', '空格', st)
        st = re.sub(r'(\d)\s*[°∘˚]', r'\1度', st)
        st = re.sub(r'[‧˙•·]{2,}', '，', st)
        st = st.replace('，。', '，').replace('。。', '。').lstrip('。')
        if st and st[-1] not in '。？！：」?!:':
            st += '。'

        opts = []
        for o in owners:
            if o == 0:
                continue
            ot = strip_zhuyin(cl(parts[o]['text'])).replace('○', '')
            if o >= 10:
                lab = LETTERS[o - 10]; ot = ot[1:].lstrip('.．、:：')
            else:
                lab = o; ot = re.sub(r'^[１２３４５①②③④⑤1-5]', '', ot)
            ot = re.sub(r'[‧˙•·]{2,}', '，', ot)
            ot = re.sub(r'[（(]\s*[_＿]+\s*[）)]', '空格', ot)
            ot = re.sub(BR, '括號', ot)
            ot = ot.rstrip('。')
            opts.append({'n': lab, 'say': f'{lab}，{ot}。', 'rects': parts[o]['rects']})

        stem_seq = [x for x in seq if x[3] == 0]
        m = PREFIX_RE.match(''.join(x[0] for x in stem_seq))
        k = len(m.group(0)) if m else 0
        stem_rects = []
        for li in range(len(bl['lines'])):
            bbs = [x[1] for x in stem_seq[k:] if x[2] == li and not x[0].isspace()]
            if bbs:
                stem_rects.append(union(bbs))

        all_rects = [r for o in owners for r in parts[o]['rects']]
        if not all_rects:
            continue
        say_all = st + ''.join(x['say'] for x in opts)
        out_blocks.append({'kind': bl['kind'], 'say': say_all, 'stemSay': st, 'box': union(all_rects),
                           'stem': stem_rects, 'opts': opts})
        dump.append(f'P{pi+1} [{bl["kind"]}] {say_all}')
        for o in opts:
            dump.append(f'      選項{o["n"]}: {o["say"]}')

    pages.append({'w': W, 'h': H, 'img': pg['img'], 'blocks': out_blocks})
    print(f'第 {pi+1} 頁（{"PDF 文字" if pg["kind"] == "text" else "OCR"}）：{len(out_blocks)} 個可點區塊，圖片 {pg["size"]//1024} KB')

title = args.title or os.path.splitext(os.path.basename(args.out))[0]
tpl = open(args.template, encoding='utf-8').read()
html = tpl.replace('__TITLE__', title).replace('/*DATA*/null', json.dumps(pages, ensure_ascii=False))
open(args.out, 'w', encoding='utf-8').write(html)
print(f'完成：{args.out}（{len(html)//1024} KB）')
if args.dump:
    open(args.dump, 'w', encoding='utf-8').write('\n'.join(dump))
    print(f'報讀文字清單：{args.dump}')
