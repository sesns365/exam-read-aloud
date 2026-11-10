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
import pymupdf, sys, json, re, base64, argparse, os, subprocess, difflib, io, unicodedata

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
LETTERS = 'ABCDEFGHㄅㄆㄇㄈㄉㄊㄋㄌㄍㄎㄏＡＢＣＤＥＦＧＨ'   # 代號：英文字母、注音符號或全形字母（同一行要有兩個以上才算代號選項）
CIRC_ALL = '①②③④⑤⑥⑦⑧⑨⑩'
CJK = lambda ch: '一' <= ch <= '鿿'
SECTION_RE = re.compile(r'^[一二三四五六七八九十]+[、．.]')
BR = r'[（(]\s*[）)]'   # 作答括號 (   )
Q_RE = re.compile(rf'^(\d+\s*[.．、]|{BR}\s*\d+\s*[.．、]|[0-9０-９]+\s*{BR})')
SUB_RE = re.compile(rf'^({BR}\s*([（(][１２３４５６７８９1-9][）)]|[{CIRC_ALL}])|[（(][１２３４５６７８９1-9][）)])')
PREFIX_RE = re.compile(rf'^\s*(\d+\s*[.．、]\s*{BR}|[0-9０-９]+\s*{BR}|{BR}\s*\d+\s*[.．、]|\d+\s*[.．、]|{BR}\s*[（(].[）)]|{BR}\s*[{CIRC_ALL}]|[（(][１２３４５６７８９1-9][）)])\s*')


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


WINGDINGS2_CIRC = {0xF06A + i: '①②③④⑤⑥⑦⑧⑨⑩'[i] for i in range(10)}   # Wingdings 2 的圈號


def norm_char(ch, font):
    """字型特殊字元轉成一般字：Wingdings 2 圈號、數學粗體／斜體字母、直書標點。回傳 '' 表示丟掉。"""
    o = ord(ch)
    if 0xE000 <= o <= 0xF8FF:   # 符號字型（Wingdings 等）的私用區字元
        return WINGDINGS2_CIRC.get(o, '') if 'wingdings2' in font.lower().replace(' ', '').replace('-', '') else ''
    if 0x1D400 <= o <= 0x1D7FF or 0xFE10 <= o <= 0xFE19 or 0xFE30 <= o <= 0xFE4F:
        return unicodedata.normalize('NFKC', ch)
    if o in (0x0305, 0x0304, 0x20D7):   # 線段上方的橫線、向量箭頭（組合字元）
        return ''
    return ch


def pdf_page_lines(p):
    """讀出 PDF 頁面的文字行（每字含位置與字級）。看不見的白色字（隱藏答案）會被丟掉。
    直書頁會重組成直行；分數（上下兩小行）與上標會合回所在的那一行。"""
    lines = []
    for b in p.get_text('rawdict')['blocks']:
        if b['type'] != 0:
            continue
        for l in b['lines']:
            chars = []
            for s in l['spans']:
                if s.get('color', 0) == 0xFFFFFF or s.get('alpha', 255) == 0:
                    continue   # 白色／透明的字：看不見，多半是藏起來的答案
                for c in s['chars']:
                    ch = norm_char(c['c'], s['font'])
                    if ch:
                        chars.append((ch, c['bbox'], s['size']))
            if chars:
                lines.append({'chars': chars, 'bbox': union([b for _, b, _ in chars])})
    if is_vertical(lines):
        W = p.rect.width
        hs, _ = border_segments(p)
        cuts = [(y, x0, x1) for y, x0, x1 in hs if x1 - x0 > 0.6 * W]
        lines = vertical_lines(lines, cuts)
        for l in lines:
            l['vertical'] = True
    else:
        lines = merge_stacked(lines)
    for l in lines:
        l['chars'] = [(c, b) for c, b, *_ in l['chars']]
        l['text'] = ''.join(c for c, _ in l['chars'])
    return lines


def is_vertical(lines):
    """直書判斷：相鄰兩字「往下排」比「往右排」多很多。"""
    seq = [c for l in lines for c in l['chars'] if c[0].strip()]
    down = right = 0
    for (a, ba, _), (b, bb, _) in zip(seq, seq[1:]):
        h = max(ba[3] - ba[1], 1)
        dx, dy = bb[0] - ba[0], bb[1] - ba[1]
        if abs(dx) < 0.5 * h and 0.5 * h < dy < 1.6 * h:
            down += 1
        elif abs(dy) < 0.5 * h and 0 < dx < 1.6 * h:
            right += 1
    return down > 200 and down > 2 * right


def vertical_lines(lines, cuts=()):
    """直書：依字的位置重組閱讀順序。先用頁面上的長橫線（例如上下兩半的分隔線）切成上下幾區，
    每區由右而左一直行一直行讀，同一直行由上而下；直行中的空隙（例如兩個選項之間）切成不同行。
    直行內橫排的數字（如 114）會接在一起。新段落（行首比前一行低一格半以上）在行首補兩個空白。"""
    seq = [c for l in lines for c in l['chars'] if c[0].strip()]
    if not seq:
        return []
    xc = lambda c: (c[1][0] + c[1][2]) / 2
    band = lambda c: sum(1 for y, x0, x1 in cuts if (c[1][1] + c[1][3]) / 2 > y and x0 <= xc(c) <= x1)
    out = []
    for bi in sorted({band(c) for c in seq}):
        cs = sorted([c for c in seq if band(c) == bi], key=xc, reverse=True)
        cols = []
        for c in cs:   # 依 x 分成直行
            h = max(c[2], 1)
            if cols and abs(xc(c) - cols[-1]['x']) < 0.6 * h:
                cols[-1]['cs'].append(c)
                cols[-1]['x'] = sum(xc(x) for x in cols[-1]['cs']) / len(cols[-1]['cs'])
            else:
                cols.append({'x': xc(c), 'cs': [c]})
        segs = []
        for col in cols:
            col['cs'].sort(key=lambda c: (round(c[1][1] / 3), c[1][0]))
            cur = []
            for c in col['cs']:
                h = max(c[2], 1)
                gap = c[1][1] - max(x[1][3] for x in cur) if cur else 0
                if cur and gap > 0.9 * h and not (cur[-1][0] in '(（︵' and gap < 5 * h):
                    segs.append(cur); cur = []
                cur.append(c)
            if cur:
                segs.append(cur)
        prev = None
        for sg in segs:
            bb = union([b for _, b, _ in sg])
            h = max(sg[0][2], 1)
            chars = list(sg)
            if prev and abs(prev[1] - bb[1]) < 6 * h and bb[1] - prev[1] > 1.5 * h and 0 < prev[0] - bb[2] < 3 * h:
                sp = [bb[0], bb[1] - 2 * h, bb[2], bb[1]]
                chars = [(' ', sp, h), (' ', sp, h)] + chars
            out.append({'chars': chars, 'bbox': bb})
            prev = (bb[0], bb[1], bb[2])
    return out


MATHY = re.compile(r'[0-9A-Za-z()+\-−.,]')


def merge_stacked(lines):
    """數學式：Word 方程式的分數（分子、分母是上下兩個小字）和上標（cm³ 的 3）在 PDF 裡被拆成別行。
    做法：先把同一列被拆開的片段接回一行，再把小字放回所在的那一列，依左右順序排好：
    分數變成「分子⁄分母」，帶分數在整數和分數之間加 ⁣，上標 2、3 變成 ²、³。
    整頁都沒有這種小字時，完全不動。"""
    allc = [c for l in lines for c in l['chars'] if c[0].strip()]
    if not allc:
        return lines
    body = sorted(c[2] for c in allc)[len(allc) // 2]
    is_small = lambda c: c[2] <= 0.85 * body and MATHY.match(c[0]) is not None
    big = lambda l: [c for c in l['chars'] if c[0].strip() and c[2] > 0.85 * body]
    if not any(is_small(c) for c in allc):
        return lines
    # 1. 同一列的片段接起來（只接 PDF 順序上相鄰、上下重疊、左右緊接的片段）
    rows = []   # 每列：{'idx', 'lines', 'y0', 'y1', 'x0', 'x1', 'extra'}
    loose = []  # 只有小字的行
    for i, l in enumerate(lines):
        b = big(l)
        if not b:
            if any(is_small(c) for c in l['chars']):
                loose.append((i, l))
            continue
        y0, y1 = min(c[1][1] for c in b), max(c[1][3] for c in b)
        x0, x1 = l['bbox'][0], l['bbox'][2]
        if rows:
            r = rows[-1]
            ov = min(r['y1'], y1) - max(r['y0'], y0)
            if ov > 0.5 * min(r['y1'] - r['y0'], y1 - y0) and -body < x0 - r['x1'] < 1.2 * body:
                r['lines'].append(l); r['x1'] = max(r['x1'], x1)
                r['y0'], r['y1'] = min(r['y0'], y0), max(r['y1'], y1)
                continue
        rows.append({'idx': i, 'lines': [l], 'y0': y0, 'y1': y1, 'x0': x0, 'x1': x1, 'extra': []})
    # 2. 只有小字的行：每個字放回最近、上下範圍涵蓋它的那一列
    left = []
    for i, l in loose:
        rest = []
        for c in l['chars']:
            cx, cy = (c[1][0] + c[1][2]) / 2, (c[1][1] + c[1][3]) / 2
            best = None
            for r in rows:
                hh = r['y1'] - r['y0']
                if r['y0'] - 0.75 * hh <= cy <= r['y1'] + 0.75 * hh and r['x0'] - body <= cx <= r['x1'] + 2 * body:
                    d = abs(cy - (r['y0'] + r['y1']) / 2)
                    if best is None or d < best[0]:
                        best = (d, r)
            (best[1]['extra'].append(c) if best and c[0].strip() else rest.append(c))
        if any(c[0].strip() for c in rest):
            left.append((i, {'chars': rest, 'bbox': union([c[1] for c in rest])}))
    # 3. 每列重新排字
    out = []
    for r in rows:
        items = [c for l in r['lines'] for c in l['chars']] + r['extra']
        if not any(is_small(c) for c in items):
            chars = items   # 這列沒有分數或上標：保持原本順序
        else:
            mid = (r['y0'] + r['y1']) / 2
            normal = [c for c in items if not is_small(c)]
            smalls = sorted([c for c in items if is_small(c)], key=lambda c: c[1][0])
            groups = []
            for c in smalls:   # 左右相連或上下疊在一起的小字算同一組
                if groups and c[1][0] <= max(x[1][2] for x in groups[-1]) + 0.25 * body:
                    groups[-1].append(c)
                else:
                    groups.append([c])
            tokens = [(c[1][0], [c]) for c in normal]
            for g in groups:
                up = sorted([c for c in g if (c[1][1] + c[1][3]) / 2 < mid - 0.1 * body], key=lambda c: c[1][0])
                dn = sorted([c for c in g if c not in up], key=lambda c: c[1][0])
                x0 = min(c[1][0] for c in g)
                if up and dn:      # 分數
                    bar = [min(up[0][1][0], dn[0][1][0]), up[-1][1][3], max(up[-1][1][2], dn[-1][1][2]), dn[0][1][1]]
                    tokens.append((x0, up + [('⁄', bar, up[0][2])] + dn))
                elif up:           # 上標（次方）
                    sup = {'2': '²', '3': '³'}
                    tokens.append((x0, [(sup.get(c[0], c[0]), c[1], c[2]) for c in up]))
                else:
                    tokens.append((x0, dn))
            tokens.sort(key=lambda t: t[0])
            chars = []
            for x0, cs in tokens:
                if '⁄' in [c[0] for c in cs]:
                    prev = next((c for c in reversed(chars) if c[0].strip()), None)
                    if prev and prev[0].isdigit() and x0 - prev[1][2] < 0.6 * body:
                        chars.append(('⁣', [x0, cs[0][1][1], x0, cs[0][1][3]], body))   # 帶分數：整數「又」分數
                chars.extend(cs)
        out.append((r['idx'], {'chars': chars, 'bbox': union([c[1] for c in chars])}))
    out += left
    out.sort(key=lambda t: t[0])
    return [l for _, l in out]


def border_segments(page):
    """收集頁面上畫出來的框線（表格、方框），回傳 (水平線, 垂直線)。"""
    hs, vs = [], []

    def add(x0, y0, x1, y1):
        x0, x1 = sorted((x0, x1)); y0, y1 = sorted((y0, y1))
        if y1 - y0 <= 2.5 and x1 - x0 > 3:
            hs.append(((y0 + y1) / 2, x0, x1))
        elif x1 - x0 <= 2.5 and y1 - y0 > 3:
            vs.append(((x0 + x1) / 2, y0, y1))
    for dr in page.get_drawings():
        for it in dr['items']:
            if it[0] == 'l':
                add(it[1].x, it[1].y, it[2].x, it[2].y)
            elif it[0] == 're':
                r = it[1]
                if r.height <= 2.5 or r.width <= 2.5:   # 細長的填色矩形＝線條
                    add(r.x0, r.y0, r.x1, r.y1)
                elif dr.get('type') in ('s', 'fs') and dr.get('color') is not None:  # 有描邊的方框：四條邊（純填色的底色塊不算）
                    add(r.x0, r.y0, r.x1, r.y0); add(r.x0, r.y1, r.x1, r.y1)
                    add(r.x0, r.y0, r.x0, r.y1); add(r.x1, r.y0, r.x1, r.y1)
    return hs, vs


def _covered(segs, pos, a, b, tol=2.0):
    """座標 pos 上的線段（可能由好幾段接成）是否完整覆蓋 a～b。"""
    iv = sorted((s, e) for p, s, e in segs if abs(p - pos) <= tol)
    cur = a + tol
    for s, e in iv:
        if s - tol <= cur:
            cur = max(cur, e)
        if cur >= b - tol:
            return True
    return cur >= b - tol


def box_of(bb, segs, W, H, tol=1.5):
    """找出字所在的方框：上下左右四條框線都要完整圍住（字下的底線、圖中的線條不算）。找不到回傳 None。"""
    hs, vs = segs
    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
    lefts = sorted({round(x, 1) for x, a, b in vs if x <= cx and a - tol <= cy <= b + tol}, reverse=True)[:5]
    rights = sorted({round(x, 1) for x, a, b in vs if x >= cx and a - tol <= cy <= b + tol})[:5]
    tops = sorted({round(y, 1) for y, a, b in hs if y <= cy and a - tol <= cx <= b + tol}, reverse=True)[:5]
    bots = sorted({round(y, 1) for y, a, b in hs if y >= cy and a - tol <= cx <= b + tol})[:5]
    best = None
    for l in lefts:
        for r in rights:
            if r - l > 0.85 * W:
                continue
            for t in tops:
                for bt in bots:
                    if bt - t > 0.5 * H:
                        continue
                    area = (r - l) * (bt - t)
                    if best and area >= best[0]:
                        continue
                    if (_covered(hs, t, l, r) and _covered(hs, bt, l, r)
                            and _covered(vs, l, t, bt) and _covered(vs, r, t, bt)):
                        best = (area, (l, t, r, bt))
    return best[1] if best else None


def math_say(s):
    """數學式唸法：分數、帶分數、次方、運算符號、比較符號、百分率、餘數等。"""
    OPD = r'[0-9A-Za-z)）\]⁄²³]'      # 運算元結尾
    OPS = r'[0-9A-Za-z(（\[⁣]'        # 運算元開頭
    s = re.sub(rf'(?<={OPD})\s*[×✕]\s*(?={OPS})', '乘', s)
    s = re.sub(rf'(?<={OPD})\s*÷\s*(?={OPS})', '除以', s)
    s = re.sub(rf'(?<={OPD})\s*[+＋]\s*(?={OPS})', '加', s)
    s = re.sub(rf'(?<={OPD})\s*[−－]\s*(?={OPS})', '減', s)
    s = re.sub(r'\s*[≦≤]\s*', '小於等於', s)
    s = re.sub(r'\s*[≧≥]\s*', '大於等於', s)
    s = re.sub(r'\s*≠\s*', '不等於', s)
    s = re.sub(r'\s*[≒≈]\s*', '約等於', s)
    s = re.sub(r'\s*≅\s*', '全等於', s)
    s = re.sub(r'\s*[=＝]\s*', '等於', s)
    s = re.sub(r'[＜<]([^＜＞<>0-9]{2,})[＞>]', r'\1', s)            # ＜恭喜你作答完畢＞：括住文字的是裝飾括號
    s = re.sub(r'\s*[<＜]\s*', '小於', s)
    s = re.sub(r'\s*[>＞]\s*', '大於', s)
    s = re.sub(r'(?<=\d)\s*[…⋯]+\s*(?=\d)', '餘', s)              # 190÷37=5…5
    # 分數與帶分數
    s = re.sub(r'(\d+)⁣\s*([0-9.]+)⁄([0-9.]+)', r'\1又\3分之\2', s)
    s = re.sub(r'([0-9A-Za-z.]+)⁄([0-9A-Za-z.]+)', r'\2分之\1', s)
    s = s.replace('⁣', '').replace('⁄', '分之')
    s = re.sub(r'(?<![\d/])(\d+)又(\d+)\s*/\s*(\d+)(?![\d/])', r'\1又\3分之\2', s)
    s = re.sub(r'(?<![\d/.])(\d+)\s*/\s*(\d+)(?![\d/])', r'\2分之\1', s)
    # 次方與面積、體積單位
    for u, zh in (('km', '公里'), ('cm', '公分'), ('mm', '公釐'), ('m', '公尺')):
        s = re.sub(rf'(?<![A-Za-z]){u}\s*²', '平方' + zh, s)
        s = re.sub(rf'(?<![A-Za-z]){u}\s*³', '立方' + zh, s)
    s = s.replace('²', '的平方').replace('³', '的立方')
    s = re.sub(r'(\d+(?:\.\d+)?)\s*[%％]', r'百分之\1', s)
    s = s.replace('π', '圓周率').replace('√', '根號').replace('∠', '角').replace('△', '三角形').replace('□', '方格')
    return s


def say_clean(s):
    """一般文字整理成朗讀用（不處理題號）。"""
    s = re.sub(r'[（(]\s*[_＿]+\s*[）)]', '空格', s)
    s = re.sub(r'[（(]\s*[）)]', '括號', s)
    s = strip_zhuyin(s)
    s = math_say(cl(s)).replace('○', '圈圈').replace('ˇ', '勾').replace('✓', '勾').replace('×', '叉').replace('＊', '').replace('※', '')
    s = re.sub(r'_{2,}|＿{2,}', '空格', s)
    s = re.sub(r'(\d)\s*[°∘˚]', r'\1度', s)
    s = re.sub(r'[‧˙•·]+', '，', s).strip('，')
    return s


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
            lines = pdf_page_lines(p)
            n_text = sum(len(cl(l['text'])) for l in lines)
            if n_text >= 20:
                pix = p.get_pixmap(dpi=args.dpi, colorspace=pymupdf.csGRAY)
                url, size = encode_img(pix)
                pages_src.append({'kind': 'text', 'img': url, 'size': size, 'w': p.rect.width, 'h': p.rect.height, 'lines': lines,
                                  'segs': border_segments(p), 'vertical': any(l.get('vertical') for l in lines)})
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
        elif Q_RE.match(s) or (pg.get('vertical') and (re.match(r'^[０-９]{1,2}(?=[一-鿿「（(])', s) or re.match(rf'^{BR}\s*[一-鿿「]', s))):
            kind = 'q'
        elif SUB_RE.match(s):
            kind = 'sub'
        elif (t.startswith('  ') or t.startswith('　')) and not s.startswith(('○', '①', '②', '③', '④')) and cur and cur['kind'] in (('para', 'section') if pg.get('vertical') else ('q', 'para', 'section')):
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
        line_start_codes = []
        for li in range(len(bl['lines'])):
            idx = [i for i, x in enumerate(seq) if x[2] == li]
            cand = []
            for n, i in enumerate(idx):
                c = seq[i][0]
                # 注音代號不會出現在一般文字中，前面不必是空白
                if c in LETTERS and (n == 0 or seq[idx[n-1]][0].isspace() or c in 'ㄅㄆㄇㄈㄉㄊㄋㄌㄍㄎㄏＡＢＣＤＥＦＧＨ'):
                    nxt = next((seq[j][0] for j in idx[n+1:] if not seq[j][0].isspace() and seq[j][0] not in '.．、'), '')
                    if CJK(nxt):
                        cand.append(i)
            if len({seq[i][0] for i in cand}) >= 2:
                code_at.update(cand)
            elif cand and cand[0] == idx[0]:
                line_start_codes.append(cand[0])
        if pg.get('vertical') and len({seq[i][0] for i in line_start_codes}) >= 2:   # 每行開頭一個代號（直書成語選填）
            code_at.update(line_start_codes)
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
            elif pending and c in 'ㄅㄆㄇㄈㄉㄊㄋㄌㄍㄎㄏ':   # ○ㄅ ○ㄆ（圈注音代號）
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

        # 方框／表格格子：題目文字中落在框線內的字，每個框各自成為一個可點區塊
        # （含題號的那個框仍算題目本身，例如連連看左框；佔題目七成以上的大框不拆）
        box_rect = {}
        if pg.get('segs'):
            stem_idx = [i for i, x in enumerate(seq) if x[3] == 0 and i >= nprefix and not x[0].isspace()]
            groups = {}
            for i in stem_idx:
                b = box_of(seq[i][1], pg['segs'], W, H)
                if b:
                    groups.setdefault(b, []).append(i)
            pre_box = {box_of(seq[i][1], pg['segs'], W, H) for i in range(nprefix) if not seq[i][0].isspace()}
            cands = [(b, idx) for b, idx in groups.items() if b not in pre_box and len(idx) < 0.7 * len(stem_idx)
                     and (len(say_clean(''.join(seq[i][0] for i in idx))) >= 2
                          or any(CJK(seq[i][0]) for i in idx))]   # 單一國字（如表頭「花」）也算；只有數字符號的圖上標記不算
            def inline(idx):   # 方框前後同一行都有框外的字 → 是夾在句子中的格子（填空），不拆
                li = seq[idx[0]][2]
                outside = [i for i in stem_idx if seq[i][2] == li and box_of(seq[i][1], pg['segs'], W, H) is None]
                return any(i < idx[0] for i in outside) and any(i > idx[-1] for i in outside)
            cands = [(b, idx) for b, idx in cands if not inline(idx)]
            if not (cands and len(stem_idx) - sum(len(idx) for _, idx in cands) >= 2):
                cands = []
            # 選項的字若跨到別的方框（例如表頭「③種子」下面一列的格子），其他方框也各自獨立
            for o in sorted({x[3] for x in seq if 0 < x[3] < 1000}):
                oidx = [i for i, x in enumerate(seq) if x[3] == o and not x[0].isspace()]
                obox = {}
                for i in oidx:
                    obox.setdefault(box_of(seq[i][1], pg['segs'], W, H), []).append(i)
                first = box_of(seq[oidx[0]][1], pg['segs'], W, H) if oidx else None
                cands += [(b, idx) for b, idx in obox.items() if b is not None and b != first]
                if first is not None:   # 選項在框內：框外的字（圖上的說明）不算這個選項
                    for i in obox.get(None, []):
                        seq[i][3] = 0
            for n, (b, idx) in enumerate(sorted(cands, key=lambda c: (c[0][1], c[0][0]))):
                for i in idx:
                    seq[i][3] = 1000 + n
                box_rect[1000 + n] = b

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
        st = re.sub(rf'^([0-9０-９]+)\s*{BR}', lambda m: f'第{int(m.group(1))}題，', st)
        st = re.sub(rf'^{BR}\s*[（(](.)[）)]', lambda m: f'第{digit(m.group(1))}小題，', st)
        st = re.sub(rf'^{BR}\s*([{CIRC_ALL}])', lambda m: f'第{CIRC_ALL.index(m.group(1)) + 1}小題，', st)
        st = re.sub(r'^[（(]([１２３４５６７８９1-9])[）)]', lambda m: f'第{digit(m.group(1))}小題，', st)
        if pg.get('vertical'):
            st = re.sub(rf'^{BR}\s*', '', st)   # 直書：題號是圖形、只剩作答括號時，括號不唸
            st = re.sub(r'^([０-９]{1,2})(?=[一-鿿「（(])', lambda m: f'第{int(m.group(1))}題，', st)   # 直書：「２原本…」
        st = re.sub(r'[（(]\s*[_＿]+\s*[）)]', '空格', st)
        st = re.sub(BR, '括號', st)
        st = strip_zhuyin(st)
        st = math_say(cl(st)).replace('○', '圈圈').replace('ˇ', '勾').replace('✓', '勾').replace('×', '叉').replace('＊', '').replace('※', '')
        st = re.sub(r'_{2,}|＿{2,}', '空格', st)
        st = re.sub(r'(\d)\s*[°∘˚]', r'\1度', st)
        st = re.sub(r'[‧˙•·]{2,}', '，', st)
        st = st.replace('，。', '，').replace('。。', '。').lstrip('。').rstrip('，')
        if st and st[-1] not in '。？！：」?!:':
            st += '。'

        opts = []
        cells = {}
        for o in owners:
            if o < 1000:
                continue
            ct = say_clean(parts[o]['text']).rstrip('。')
            if ct:
                cells[o] = ct
        for o in owners:
            if o == 0:
                continue
            if o >= 1000:  # 方框／表格格子
                if o not in cells:
                    continue
                ct = cells[o]
                b = box_rect[o]
                # 只有數值的格子（例如「3本」）：前面加上同一列第一格，唸成「A 磁鐵：不加鐵片，3本」
                if len(ct) <= 4 and re.search(r'[0-9０-９?？]', ct):
                    row = [p for p, rb in box_rect.items() if p in cells and p != o and rb[0] < b[0]
                           and min(rb[3], b[3]) - max(rb[1], b[1]) > 0.5 * (b[3] - b[1])]
                    if row:
                        ct = cells[min(row, key=lambda p: box_rect[p][0])] + '，' + ct
                opts.append({'n': '格', 'cell': True, 'say': ct + '。', 'rects': parts[o]['rects']})
                continue
            ot = math_say(strip_zhuyin(cl(parts[o]['text']))).replace('○', '')
            if o >= 10:
                lab = LETTERS[o - 10]; ot = ot[1:].lstrip('.．、:：')
            else:
                lab = o; ot = re.sub(r'^[１２３４５①②③④⑤1-5]', '', ot)
            ot = re.sub(r'[‧˙•·]{2,}', '，', ot)
            ot = re.sub(r'[（(]\s*[_＿]+\s*[）)]', '空格', ot)
            ot = re.sub(BR, '括號', ot)
            ot = ot.rstrip('。，')
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
