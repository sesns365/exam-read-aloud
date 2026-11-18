"""把多份報讀版考卷做成一個「首頁」，學生點選考卷名稱就能開啟。

用法：
  python -I make_index.py <報讀版1.html> <報讀版2.html> ... -o <首頁.html> [--title 首頁標題] [--names 名稱1 名稱2 ...]

- 考卷依指令中的順序排列。
- 名稱預設讀取每個 HTML 的 <title>（去掉「（報讀版）」）；用 --names 可以自訂（數量要和檔案一樣多）。
- 連結用「相對於首頁的路徑」，所以首頁和考卷要維持同樣的資料夾相對位置
  （本機：放同一個資料夾；部署：首頁放在上層資料夾，考卷放各自的子資料夾 index.html）。
- 每張卡片右側有 🔊 按鈕，可以唸出考卷名稱。

密碼（選用）：
  --password 密碼            點考卷時要先輸入密碼才能開啟
  --lock 1 3                 只有第 1、3 份要密碼（不寫就是全部都要）
  上鎖的考卷連結會用密碼加密後才寫進首頁，看首頁原始碼也找不到考卷網址。
  這是「避免學生提早或自行打開」用的鎖，不是高強度保密：知道考卷網址的人仍可直接開啟，
  所以部署時考卷資料夾名稱要加一段亂碼（見 SKILL.md）。
"""
import argparse, hashlib, html, os, re, secrets, sys, urllib.parse

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
ap = argparse.ArgumentParser()
ap.add_argument('files', nargs='+')
ap.add_argument('-o', '--out', required=True)
ap.add_argument('--title', default='報讀版考卷')
ap.add_argument('--names', nargs='*', default=None)
ap.add_argument('--password', default=None)
ap.add_argument('--lock', nargs='*', type=int, default=None, help='要上鎖的考卷序號（從 1 開始）；不寫就全部上鎖')
args = ap.parse_args()
if args.lock and not args.password:
    sys.exit('--lock 要搭配 --password')
ITER = 100000


def encrypt(text, pw):
    """PBKDF2-SHA256 推出金鑰，再用 SHA-256(金鑰＋計數) 當金鑰流做 XOR（網頁端用 WebCrypto 解）。"""
    salt = secrets.token_bytes(16)
    key = hashlib.pbkdf2_hmac('sha256', pw.encode('utf-8'), salt, ITER, 32)
    data = ('OK:' + text).encode('utf-8')
    stream = b''.join(hashlib.sha256(key + i.to_bytes(4, 'big')).digest() for i in range(len(data) // 32 + 1))
    return salt.hex(), bytes(a ^ b for a, b in zip(data, stream)).hex()
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

locked = set()
if args.password:
    locked = set(args.lock) if args.lock else set(range(1, len(items) + 1))
    bad = [n for n in locked if not 1 <= n <= len(items)]
    if bad:
        sys.exit(f'--lock 的序號超出範圍：{bad}')


def card(i, n, href):
    if i in locked:
        salt, enc = encrypt(href, args.password)
        a = (f'<a class="card" href="#" data-salt="{salt}" data-enc="{enc}">'
             f'<span class="no">{i}</span><span class="name">{html.escape(n)}</span><span class="lock" aria-label="需要密碼">🔒</span></a>')
    else:
        a = f'<a class="card" href="{href}"><span class="no">{i}</span><span class="name">{html.escape(n)}</span></a>'
    return (f'    <li>{a}<button class="say" type="button" data-say="{html.escape(n)}" '
            f'aria-label="唸出「{html.escape(n)}」">🔊</button></li>')


cards = '\n'.join(card(i, n, h) for i, (n, h) in enumerate(items, 1))
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
  .name {{ flex:1; }}
  .lock {{ font-size:24px; }}
  dialog {{ border:none; border-radius:16px; padding:24px; width:min(92vw,420px); box-shadow:0 10px 40px rgba(0,0,0,.3); }}
  dialog::backdrop {{ background:rgba(0,0,0,.45); }}
  dialog h2 {{ margin:0 0 6px; font-size:24px; }}
  dialog p {{ margin:0 0 14px; color:var(--muted); font-size:18px; }}
  dialog input {{ width:100%; font-size:28px; padding:10px 14px; border:2px solid var(--line); border-radius:12px; letter-spacing:.15em; }}
  dialog input:focus {{ border-color:var(--accent); outline:none; }}
  .err {{ color:#c62828; min-height:1.5em; margin:8px 0 0; font-size:18px; font-weight:bold; }}
  .btns {{ display:flex; gap:10px; margin-top:12px; }}
  .btns button {{ flex:1; font-size:22px; padding:12px; border-radius:12px; border:2px solid var(--line); background:#fff; cursor:pointer; }}
  .btns .ok {{ background:var(--bar); border-color:var(--bar); color:#fff; font-weight:bold; }}
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
<dialog id="pw">
  <form method="dialog" id="pwForm">
    <h2>🔒 <span id="pwName"></span></h2>
    <p>請輸入老師給的密碼</p>
    <input id="pwInput" type="password" autocomplete="off" aria-label="密碼">
    <div class="err" id="pwErr" role="alert"></div>
    <div class="btns"><button type="button" id="pwCancel">取消</button><button class="ok" type="submit">開啟考卷</button></div>
  </form>
</dialog>
<script>
  var ITER = {ITER};
  function hex(h) {{ var a = new Uint8Array(h.length / 2); for (var i = 0; i < a.length; i++) a[i] = parseInt(h.substr(i * 2, 2), 16); return a; }}
  async function unlock(pw, salt, enc) {{
    var te = new TextEncoder(), subtle = crypto.subtle;
    var base = await subtle.importKey('raw', te.encode(pw), 'PBKDF2', false, ['deriveBits']);
    var key = new Uint8Array(await subtle.deriveBits({{ name: 'PBKDF2', salt: hex(salt), iterations: ITER, hash: 'SHA-256' }}, base, 256));
    var data = hex(enc), out = new Uint8Array(data.length);
    for (var b = 0; b * 32 < data.length; b++) {{
      var blk = new Uint8Array(36); blk.set(key); new DataView(blk.buffer).setUint32(32, b);
      var ks = new Uint8Array(await subtle.digest('SHA-256', blk));
      for (var j = 0; j < 32 && b * 32 + j < data.length; j++) out[b * 32 + j] = data[b * 32 + j] ^ ks[j];
    }}
    var txt = new TextDecoder('utf-8', {{ fatal: false }}).decode(out);
    return txt.indexOf('OK:') === 0 ? txt.slice(3) : null;
  }}
  var dlg = document.getElementById('pw'), cur = null;
  document.querySelectorAll('.card[data-enc]').forEach(function (c) {{
    c.addEventListener('click', function (e) {{
      e.preventDefault(); cur = c;
      document.getElementById('pwName').textContent = c.querySelector('.name').textContent;
      document.getElementById('pwErr').textContent = '';
      document.getElementById('pwInput').value = '';
      dlg.showModal(); document.getElementById('pwInput').focus();
    }});
  }});
  document.getElementById('pwCancel').addEventListener('click', function () {{ dlg.close(); }});
  document.getElementById('pwForm').addEventListener('submit', async function (e) {{
    e.preventDefault();
    var err = document.getElementById('pwErr');
    if (!window.crypto || !crypto.subtle) {{ err.textContent = '這個瀏覽器不支援密碼功能，請改用 Chrome 或 Edge'; return; }}
    var href = await unlock(document.getElementById('pwInput').value, cur.dataset.salt, cur.dataset.enc);
    if (href) {{ dlg.close(); location.href = href; }}
    else {{ err.textContent = '密碼不對，請再試一次'; document.getElementById('pwInput').select(); }}
  }});
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
print(f'首頁已產生：{args.out}（{len(items)} 份考卷' + (f'，{len(locked)} 份需要密碼' if locked else '') + '）')
for n, h in items:
    print(f'  {n} → {h}' + ('  🔒' if items.index((n, h)) + 1 in locked else ''))
