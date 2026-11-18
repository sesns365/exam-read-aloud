"""把多份報讀版考卷做成一個「首頁」，學生點選考卷名稱就能開啟。

用法：
  python -I make_index.py <報讀版1.html> <報讀版2.html> ... -o <首頁.html> [--title 首頁標題] [--names 名稱1 名稱2 ...]

- 考卷依指令中的順序排列。
- 名稱預設讀取每個 HTML 的 <title>（去掉「（報讀版）」）；用 --names 可以自訂（數量要和檔案一樣多）。
- 連結用「相對於首頁的路徑」，所以首頁和考卷要維持同樣的資料夾相對位置
  （本機：放同一個資料夾；部署：首頁放在上層資料夾，考卷放各自的子資料夾 index.html）。
- 每張卡片右側有 🔊 按鈕，可以唸出考卷名稱。
"""
import argparse, html, os, re, sys, urllib.parse

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
ap = argparse.ArgumentParser()
ap.add_argument('files', nargs='+')
ap.add_argument('-o', '--out', required=True)
ap.add_argument('--title', default='報讀版考卷')
ap.add_argument('--names', nargs='*', default=None)
args = ap.parse_args()
if args.names and len(args.names) != len(args.files):
    sys.exit('--names 的數量要和考卷檔案一樣多')

out_dir = os.path.dirname(os.path.abspath(args.out))


def read_title(path):
    try:
        with open(path, encoding='utf-8') as f:
            m = re.search(r'<title>(.*?)</title>', f.read(20000), re.S)
        if m:
            t = html.unescape(m.group(1)).strip()
            return re.sub(r'\s*[（(]報讀版[）)]\s*$', '', t) or None
    except OSError:
        pass
    return None


items = []
for i, f in enumerate(args.files):
    name = (args.names[i] if args.names else None) or read_title(f) \
        or re.sub(r'_報讀版$', '', os.path.splitext(os.path.basename(f))[0])
    rel = os.path.relpath(os.path.abspath(f), out_dir).replace(os.sep, '/')
    items.append((name, urllib.parse.quote(rel, safe='/')))

cards = '\n'.join(
    f'    <li><a class="card" href="{href}"><span class="no">{i}</span><span class="name">{html.escape(n)}</span></a>'
    f'<button class="say" type="button" data-say="{html.escape(n)}" aria-label="唸出「{html.escape(n)}」">🔊</button></li>'
    for i, (n, href) in enumerate(items, 1))
title = html.escape(args.title)

page = f'''<!doctype html>
<html lang="zh-Hant-TW">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
  :root {{ --bar:#1f3b57; --card:#fff; --text:#1d2733; --muted:#5b6876; --line:#d5dbe1; --accent:#ff8c00; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:#e9ecef; color:var(--text); font-family:"Microsoft JhengHei","PingFang TC",sans-serif; }}
  header {{ background:var(--bar); color:#fff; padding:20px 16px; text-align:center; }}
  h1 {{ margin:0; font-size:clamp(24px,5vw,34px); letter-spacing:.05em; }}
  .hint {{ margin:8px 0 0; opacity:.85; font-size:16px; }}
  ul {{ list-style:none; margin:24px auto; padding:0 16px; max-width:760px; display:flex; flex-direction:column; gap:14px; }}
  li {{ display:flex; gap:10px; }}
  .card {{ flex:1; display:flex; align-items:center; gap:14px; min-height:72px; padding:14px 18px; background:var(--card);
          border:2px solid var(--line); border-radius:14px; color:inherit; text-decoration:none; font-size:clamp(20px,4vw,26px); font-weight:bold; }}
  .card:hover, .card:focus-visible {{ border-color:var(--accent); outline:none; box-shadow:0 0 0 4px rgba(255,140,0,.25); }}
  .no {{ flex:none; width:42px; height:42px; border-radius:50%; background:var(--bar); color:#fff; display:grid; place-items:center; font-size:20px; }}
  .say {{ flex:none; width:72px; border:2px solid var(--line); border-radius:14px; background:var(--card); font-size:28px; cursor:pointer; }}
  .say:hover, .say:focus-visible {{ border-color:var(--accent); outline:none; }}
</style>
</head>
<body>
<header>
  <h1>{title}</h1>
  <p class="hint">點考卷名稱開始作答，按 🔊 可以聽名稱</p>
</header>
<main>
  <ul>
{cards}
  </ul>
</main>
<script>
  document.querySelectorAll('.say').forEach(function (b) {{
    b.addEventListener('click', function () {{
      if (!window.speechSynthesis) return;
      speechSynthesis.cancel();
      var u = new SpeechSynthesisUtterance(b.dataset.say);
      u.lang = 'zh-TW';
      try {{ var v = localStorage.getItem('voice'); if (v) {{ var m = speechSynthesis.getVoices().find(function (x) {{ return x.name === v; }}); if (m) u.voice = m; }} }} catch (e) {{}}
      speechSynthesis.speak(u);
    }});
  }});
</script>
</body>
</html>
'''
with open(args.out, 'w', encoding='utf-8') as f:
    f.write(page)
print(f'首頁已產生：{args.out}（{len(items)} 份考卷）')
for n, h in items:
    print(f'  {n} → {h}')
