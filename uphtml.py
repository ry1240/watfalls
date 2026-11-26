#!/usr/bin/env python3
"""100falls.xlsx の Sheet1 から index.html を生成する。

使い方: python uphtml.py [入力.xlsx] [出力.html]
生成するページはスマートフォン専用（地図はピンチ/ボタンで拡大縮小、滝リストは表形式）。
列構成: A=地域 B=都道府県 C=滝名 D=Googlemapリンク E=高さ F=active G=訪瀑日
（見出し行なし。地域・都道府県は結合セルなので上から引き継ぐ）
"""
import json, os, re, sys, datetime
try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl が見つかりません。先に  pip install openpyxl  を実行してください。")

src = sys.argv[1] if len(sys.argv) > 1 else "100falls.xlsx"
dst = sys.argv[2] if len(sys.argv) > 2 else "index.html"


def fmt_dates(v):
    if v is None or v == "":
        return []
    if isinstance(v, (datetime.datetime, datetime.date)):
        return [f"{v.year}/{v.month}/{v.day}"]
    return [s.strip() for s in str(v).split(",") if s.strip()]


if not os.path.exists(src):
    sys.exit(f"入力ファイルが見つかりません: {os.path.abspath(src)}\n(カレントフォルダ: {os.getcwd()})")
wb = openpyxl.load_workbook(src, data_only=True)
if "Sheet1" not in wb.sheetnames:
    sys.exit(f"Sheet1 がありません。シート名: {wb.sheetnames}")
ws = wb["Sheet1"]
rows, region, pref = [], "", ""
for a, b, c, d, e, f, g in ws.iter_rows(min_col=1, max_col=7, values_only=True):
    if not c:
        continue
    region = a or region
    pref = b or pref
    url = (d or "").strip() if isinstance(d, str) else ""
    m = re.search(r"@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)", url)
    rows.append({
        "lat": float(m.group(1)) if m else None, "lon": float(m.group(2)) if m else None,
        "region": region, "pref": pref, "name": str(c).strip(),
        "url": (d or "").strip() if isinstance(d, str) else "",
        "height": str(e).strip() if e is not None else "",
        "active": str(f).strip().lower() == "active" if f else False,
        "dates": fmt_dates(g),
    })

data = json.dumps(rows, ensure_ascii=False).replace("</", "<\\/")
TEMPLATE = r'''<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>日本の滝100選 </title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Shippori+Mincho:wght@600;800&family=Zen+Kaku+Gothic+New:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{--bg:#eef3f4;--surface:#fff;--ink:#12303a;--sub:#587380;--line:#d3e0e3;--accent:#1d6f8a;--on:#1d6f8a;--off:#cfdde1;--tint:#e2f0f4;
 box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px);color-scheme:light}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#0e1d23;--surface:#142830;--ink:#e3eef1;--sub:#90abb4;--line:#25414b;--accent:#6cc3dc;--on:#6cc3dc;--off:#27434d;--tint:#183843;color-scheme:dark}}
:root[data-theme="dark"]{--bg:#0e1d23;--surface:#142830;--ink:#e3eef1;--sub:#90abb4;--line:#25414b;--accent:#6cc3dc;--on:#6cc3dc;--off:#27434d;--tint:#183843;color-scheme:dark}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:400 15px/1.7 "Zen Kaku Gothic New","Hiragino Sans","Yu Gothic",Meiryo,sans-serif}
.wrap{max-width:480px;margin:0 auto;padding:20px 12px 40px}
h1{font:800 28px/1.2 "Shippori Mincho","Hiragino Mincho ProN","Yu Mincho",serif;margin:0 0 4px;letter-spacing:.04em}
.lead{color:var(--sub);margin:0 0 20px}
.top{display:flex;flex-direction:column;align-items:center;gap:14px;margin-bottom:20px}
.hero{display:flex;flex-direction:column;align-items:center;gap:8px;text-align:center}
.pie{width:120px;height:120px;flex:none}
.pie .trk{fill:var(--off)}
.pie .sec{fill:var(--on)}
.count{font-family:"Shippori Mincho",serif;line-height:1}
.count b{font-size:46px;font-weight:800;color:var(--accent)}
.count span{font-size:24px;font-weight:600;color:var(--sub);margin-left:4px}
.map{position:relative;width:100%;min-width:0}
.map svg{display:block;width:100%;height:auto;touch-action:pan-y;user-select:none;-webkit-user-select:none;-webkit-tap-highlight-color:transparent;background:var(--surface);border-radius:8px;outline:1px solid var(--line);outline-offset:-1px}
.zoom{position:absolute;top:8px;left:8px;display:flex;flex-direction:column;gap:6px}
.zoom button{width:44px;height:44px;padding:0;font:700 22px/1 "Zen Kaku Gothic New",sans-serif;color:var(--ink);background:var(--surface);border:1px solid var(--line);border-radius:8px;opacity:.94;touch-action:manipulation;cursor:pointer}
.zoom #zrst{font-size:12px;font-weight:500}
.land{fill:var(--tint);stroke:var(--line);stroke-width:1;stroke-linejoin:round;vector-effect:non-scaling-stroke}
.box{fill:none;stroke:var(--line);stroke-dasharray:3 3;vector-effect:non-scaling-stroke}
.lbl{font:500 11px "Zen Kaku Gothic New",sans-serif;fill:var(--sub)}
.pin path{fill:#ea4335}
.pin.off path{fill:var(--sub);opacity:.45}
.pin{cursor:pointer}
.pin circle{fill:#fff}
.pin .hit{fill:transparent}
.pin:focus-visible{outline:none}
.pin:focus-visible path{stroke:var(--ink);stroke-width:1.5}
.cap{margin:8px 0 0;font-size:14px;color:var(--sub);min-height:3.2em}
.cap b{color:var(--ink);font-size:16px;margin-right:6px}
.wrap{padding-left:max(12px,env(safe-area-inset-left));padding-right:max(12px,env(safe-area-inset-right))}
.tools{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:20px}
.tools input,.tools select,.chip{font:inherit;color:var(--ink);background:var(--surface);border:1px solid var(--line);border-radius:6px;padding:6px 10px}
.tools input{flex:1 1 180px;min-width:0}
.chip{cursor:pointer;padding:4px 12px}
.tools input,.tools select{font-size:16px;min-height:44px}
#chips{flex-wrap:nowrap;overflow-x:auto;padding-bottom:6px;-webkit-overflow-scrolling:touch}
.chip{flex:none;white-space:nowrap;min-height:40px;padding:6px 16px}
.chip[aria-pressed="true"]{background:var(--accent);border-color:var(--accent);color:var(--bg)}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
section{margin-bottom:28px}
h2{font:700 18px/1.3 "Shippori Mincho",serif;margin:0 0 8px;display:flex;justify-content:space-between;align-items:baseline;border-bottom:2px solid var(--ink);padding-bottom:4px}
h2 small{font:400 13px "Zen Kaku Gothic New",sans-serif;color:var(--sub)}
.tbl{overflow-x:auto;max-width:100%}
table{border-collapse:collapse;width:100%;background:var(--surface)}
th,td{text-align:left;padding:8px 5px;border-bottom:1px solid var(--line);vertical-align:top;font-size:12px;line-height:1.5}
th{font-weight:500;color:var(--sub);font-size:13px;white-space:nowrap}
td.pref{color:var(--sub);white-space:nowrap;vertical-align:middle;text-align:center;font-size:11px;padding:8px 3px;border-right:1px solid var(--line)}
table{border-top:1px solid var(--line);table-layout:fixed}
td.nm{font-weight:700;font-size:14px;line-height:1.4}
td:last-child a{display:inline-block;padding:4px 2px}
.ck{display:inline-block;min-width:1em;margin-right:6px;color:#1a73e8;font-weight:700;line-height:1}
@media (prefers-color-scheme:dark){.ck{color:#6cb4ff}} 
a{color:var(--accent)}
.act{color:var(--accent);font-weight:700}
.none{color:var(--sub)}
.empty{padding:32px 0;text-align:center;color:var(--sub)}

</style>
</head>
<body>
<div class="wrap">
<h1>日本の滝100選</h1>
<p class="lead">訪瀑の記録</p>
<div class="top">
<div class="hero">
<svg class="pie" viewBox="0 0 120 120" role="img" id="pie"><circle class="trk" cx="60" cy="60" r="60"/><path class="sec" id="sec"/></svg>
<div class="count"><b id="n">0</b><span id="of"></span></div>
</div>
<div class="map"><div class="zoom"><button type="button" id="zin" aria-label="拡大">＋</button><button type="button" id="zout" aria-label="縮小">－</button><button type="button" id="zrst" aria-label="全体表示">全体</button></div><svg viewBox="0 0 515 590" role="img" aria-label="百名滝の位置を示す日本地図" id="map"></svg><p class="cap" id="cap"></p></div>
</div>
<div class="tools">
<input id="q" type="search" placeholder="滝名・都道府県で検索" aria-label="検索">
<select id="st" aria-label="状態"><option value="">すべて</option><option value="a">行った</option><option value="n">まだ</option></select>
</div>
<div class="tools" id="chips" role="group" aria-label="地域"></div>
<main id="out"></main>
</div>
<script>
const DATA = __DATA__;
const esc = s => String(s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
DATA.forEach((d, i) => { d.i = i; });
const regions = [...new Set(DATA.map(d => d.region))];
const state = {q: "", st: "", region: ""};

const total = DATA.length, done = DATA.filter(d => d.active).length;
document.getElementById("n").textContent = done;
document.getElementById("of").textContent = "/" + total;
document.getElementById("pie").setAttribute("aria-label", "訪瀑済 " + done + "/" + total);
(function () {
  const sec = document.getElementById("sec");
  if (!done) return;
  if (done >= total) { sec.outerHTML = '<circle class="sec" cx="60" cy="60" r="60"/>'; return; }
  const t = 2 * Math.PI * done / total;
  sec.setAttribute("d", "M60 60L60 0A60 60 0 " + (t > Math.PI ? 1 : 0) + " 1 " + (60 + 60 * Math.sin(t)).toFixed(2) + " " + (60 - 60 * Math.cos(t)).toFixed(2) + "Z");
})();

const LAND = [
 [[141.95,45.5],[142.6,44.85],[143.4,44.3],[144.4,44.05],[145.0,44.35],[145.3,44.3],[144.95,43.95],[145.25,43.55],[145.8,43.38],[145.2,43.15],[144.4,42.95],[143.3,42.95],[143.25,41.93],[142.1,42.3],[141.6,42.65],[141.0,42.33],[140.7,42.55],[140.45,42.2],[140.95,41.8],[140.5,41.65],[140.1,41.42],[140.1,41.87],[139.85,42.45],[140.35,43.35],[141.0,43.2],[141.3,43.35],[141.65,43.95],[141.7,45.4]],
 [[140.85,41.5],[141.45,41.4],[141.55,40.5],[141.95,39.65],[141.9,39.25],[141.6,38.3],[141.0,38.2],[140.95,37.0],[140.6,36.3],[140.85,35.7],[140.4,35.1],[139.88,34.9],[140.1,35.3],[139.95,35.65],[139.7,35.4],[139.65,35.15],[139.15,35.25],[139.1,34.7],[138.85,34.6],[138.75,35.0],[138.5,34.95],[138.22,34.6],[137.7,34.65],[137.0,34.6],[136.9,35.05],[136.7,34.6],[136.9,34.3],[136.2,34.07],[135.77,33.43],[135.37,33.73],[135.1,34.2],[135.4,34.6],[135.2,34.7],[134.7,34.75],[133.9,34.55],[133.3,34.45],[132.4,34.35],[132.2,34.15],[131.5,34.0],[130.9,33.95],[131.4,34.4],[132.1,34.9],[133.0,35.5],[133.3,35.57],[134.2,35.5],[134.8,35.65],[135.7,35.5],[136.0,35.65],[136.1,36.1],[136.6,36.6],[136.9,37.05],[136.9,37.4],[137.35,37.5],[137.2,36.8],[137.85,37.05],[138.25,37.2],[139.0,37.9],[139.5,38.5],[139.8,38.75],[139.7,39.9],[140.0,40.2],[139.95,40.6],[140.35,41.26],[140.75,40.9],[141.25,40.95],[141.1,41.4]],
 [[134.6,34.25],[134.7,33.9],[134.2,33.25],[133.55,33.5],[133.0,32.72],[132.5,33.2],[132.0,33.35],[132.45,33.65],[132.7,33.9],[133.0,34.1],[133.6,34.0],[134.05,34.35]],
 [[130.95,33.95],[131.5,33.6],[131.6,33.25],[131.9,33.25],[131.9,32.95],[131.7,32.6],[131.45,31.9],[131.35,31.4],[131.0,31.3],[130.7,31.0],[130.65,31.4],[130.55,31.6],[130.55,31.15],[130.3,31.27],[130.3,31.6],[130.2,32.0],[130.4,32.2],[130.55,32.5],[130.65,32.75],[130.4,32.9],[130.3,33.15],[130.0,32.9],[129.75,32.58],[129.7,33.15],[129.9,33.45],[130.35,33.6],[130.5,33.8]],
 [[138.5,38.3],[138.2,38.3],[138.25,37.95],[138.5,37.8]],
 [[135.0,34.6],[134.85,34.2],[134.7,34.3],[134.95,34.6]],
 [[130.65,30.45],[130.45,30.45],[130.4,30.3],[130.55,30.2],[130.67,30.3]]
];
const OKI = [
 [[128.27,26.87],[128.05,26.65],[127.9,26.4],[127.85,26.15],[127.65,26.08],[127.65,26.3],[127.75,26.5],[127.95,26.72],[128.2,26.85]],
 [[125.45,24.85],[125.25,24.9],[125.2,24.72],[125.4,24.7]],
 [[124.35,24.55],[124.1,24.5],[124.15,24.35],[124.3,24.4]],
 [[123.95,24.4],[123.8,24.45],[123.7,24.38],[123.8,24.25],[123.95,24.3]]
];
const pj = (lon, lat) => [(lon - 129.2) * 28.75 + 14, (45.8 - lat) * 36 + 14];
const pjo = (lon, lat) => [(lon - 122.6) * 23.4 + 345, (27.1 - lat) * 26 + 472];
const poly = (pts, f) => '<polygon class="land" points="' + pts.map(q => f(q[0], q[1]).map(v => v.toFixed(1)).join(",")).join(" ") + '"/>';
const PIN = '<path d="M0 0C-3-7-7-10-7-15A7 7 0 0 1 7-15C7-10 3-7 0 0Z"/><circle cy="-15" r="2.6"/>';
const cap = document.getElementById("cap"), HINT = "ピンをタップすると滝の情報が表示されます。ピンチまたは＋－ボタンで拡大縮小できます。";
function showPin(g) {
  const d = DATA[+g.dataset.i];
  cap.innerHTML = "<b>" + esc(d.name) + "</b>" + esc(d.pref) + (d.height ? " ／ " + esc(d.height.replace(/,/g, " / ")) : "") + (d.active ? " ／ active" : "") +
    (d.url ? ' <a href="' + esc(d.url) + '" target="_blank" rel="noopener">Googleマップで開く</a>' : "");
}
const VW = 515, VH = 590, MAXZ = 8;
const mapEl = document.getElementById("map");
const vb = {x: 0, y: 0, z: 1};
function applyVB() {
  const w = VW / vb.z, h = VH / vb.z;
  vb.x = Math.min(Math.max(vb.x, 0), VW - w);
  vb.y = Math.min(Math.max(vb.y, 0), VH - h);
  mapEl.setAttribute("viewBox", vb.x.toFixed(2) + " " + vb.y.toFixed(2) + " " + w.toFixed(2) + " " + h.toFixed(2));
  mapEl.style.touchAction = vb.z > 1.01 ? "none" : "pan-y";
  const s = (0.8 / Math.sqrt(vb.z)).toFixed(3);
  mapEl.querySelectorAll(".pin").forEach(g => g.setAttribute("transform", "translate(" + g.dataset.x + " " + g.dataset.y + ") scale(" + s + ")"));
}
function zoomAt(f, cx, cy) {
  const r = mapEl.getBoundingClientRect();
  const px = (cx - r.left) / r.width, py = (cy - r.top) / r.height;
  const sx = vb.x + px * VW / vb.z, sy = vb.y + py * VH / vb.z;
  vb.z = Math.min(MAXZ, Math.max(1, vb.z * f));
  vb.x = sx - px * VW / vb.z; vb.y = sy - py * VH / vb.z;
  applyVB();
}
function panBy(dx, dy) {
  const r = mapEl.getBoundingClientRect();
  vb.x -= dx * (VW / vb.z) / r.width; vb.y -= dy * (VH / vb.z) / r.height;
  applyVB();
}
const ptrs = new Map();
let gest = {moved: 0, multi: false, c: null};
const pinch = () => { const a = [...ptrs.values()]; return [(a[0].x + a[1].x) / 2, (a[0].y + a[1].y) / 2, Math.hypot(a[0].x - a[1].x, a[0].y - a[1].y)]; };
mapEl.addEventListener("pointerdown", e => {
  if (!ptrs.size) gest = {moved: 0, multi: false, c: null};
  ptrs.set(e.pointerId, {x: e.clientX, y: e.clientY});
  if (ptrs.size === 2) { gest.multi = true; gest.c = pinch(); }
});
mapEl.addEventListener("pointermove", e => {
  const p = ptrs.get(e.pointerId);
  if (!p) return;
  if (ptrs.size === 1) {
    const dx = e.clientX - p.x, dy = e.clientY - p.y;
    gest.moved += Math.abs(dx) + Math.abs(dy);
    p.x = e.clientX; p.y = e.clientY;
    if (vb.z > 1.01 && gest.moved > 6) panBy(dx, dy);
  } else if (ptrs.size === 2) {
    p.x = e.clientX; p.y = e.clientY;
    const c = pinch(), o = gest.c;
    if (o && o[2] > 0 && c[2] > 0) zoomAt(c[2] / o[2], c[0], c[1]);
    if (o) panBy(c[0] - o[0], c[1] - o[1]);
    gest.c = c;
  }
});
function endPtr(e) {
  if (!ptrs.has(e.pointerId)) return;
  const tap = e.type === "pointerup" && ptrs.size === 1 && !gest.multi && gest.moved < 8;
  ptrs.delete(e.pointerId);
  if (tap) {
    const el = document.elementFromPoint(e.clientX, e.clientY);
    const g = el && el.closest ? el.closest(".pin") : null;
    if (g) showPin(g);
  }
}
mapEl.addEventListener("pointerup", endPtr);
mapEl.addEventListener("pointercancel", endPtr);
const zc = () => { const r = mapEl.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; };
document.getElementById("zin").onclick = () => zoomAt(1.6, ...zc());
document.getElementById("zout").onclick = () => zoomAt(1 / 1.6, ...zc());
document.getElementById("zrst").onclick = () => { vb.z = 1; vb.x = 0; vb.y = 0; applyVB(); };
mapEl.addEventListener("keydown", e => { const g = e.target.closest && e.target.closest(".pin"); if (g && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); showPin(g); } });
function renderMap(list) {
  const pin = d => {
    if (d.lat == null) return "";
    const xy = d.lat < 28 ? pjo(d.lon, d.lat) : pj(d.lon, d.lat);
    return '<g class="pin' + (d.active ? "" : " off") + '" data-i="' + d.i + '" data-x="' + xy[0].toFixed(1) + '" data-y="' + xy[1].toFixed(1) + '" tabindex="0" role="button" aria-label="' + esc(d.name) + '" transform="translate(' + xy[0].toFixed(1) + " " + xy[1].toFixed(1) + ') scale(.8)"><title>' + esc(d.name) + '</title><circle class="hit" cy="-9" r="13"/>' + PIN + "</g>";
  };
  cap.textContent = HINT;
  document.getElementById("map").innerHTML =
    LAND.map(p => poly(p, pj)).join("") +
    '<rect class="box" x="335" y="458" width="170" height="104" rx="4"/><text class="lbl" x="343" y="474">沖縄</text>' +
    OKI.map(p => poly(p, pjo)).join("") +
    list.filter(d => !d.active).map(pin).join("") + list.filter(d => d.active).map(pin).join("");
  applyVB();
}

const chips = document.getElementById("chips");
["", ...regions].forEach(r => {
  const b = document.createElement("button");
  b.className = "chip"; b.type = "button"; b.textContent = r || "全地域";
  b.setAttribute("aria-pressed", String(r === state.region));
  b.onclick = () => { state.region = r; [...chips.children].forEach(c => c.setAttribute("aria-pressed", String(c === b))); render(); };
  chips.appendChild(b);
});
document.getElementById("q").oninput = e => { state.q = e.target.value.trim(); render(); };
document.getElementById("st").onchange = e => { state.st = e.target.value; render(); };

const ck = d => '<span class="ck"' + (d.active ? ' role="img" aria-label="訪瀑済">\u2714\uFE0E' : ' aria-hidden="true">\u3000') + '</span>';
function render() {
  const q = state.q.toLowerCase();
  const list = DATA.filter(d =>
    (!state.region || d.region === state.region) &&
    (!state.st || (state.st === "a") === d.active) &&
    (!q || (d.name + d.pref).toLowerCase().includes(q)));
  renderMap(list);
  if (!list.length) { document.getElementById("out").innerHTML = '<p class="empty">条件に合う滝がありません。</p>'; return; }
  let html = "";
  regions.filter(r => list.some(d => d.region === r)).forEach(r => {
    const items = list.filter(d => d.region === r);
    let prev = "";
    html += '<section><h2>' + esc(r) + '<small>' + items.filter(d => d.active).length + " / " + items.length + '</small></h2><div class="tbl"><table><colgroup><col style="width:52px"><col><col style="width:44px"><col style="width:76px"><col style="width:40px"></colgroup><tbody>';
    items.forEach((d, i) => {
      let pc = "";
      if (d.pref !== prev) { let n = 1; while (items[i + n] && items[i + n].pref === d.pref) n++; pc = '<td class="pref" rowspan="' + n + '">' + esc(d.pref) + '</td>'; }
      prev = d.pref;
      html += '<tr class="' + (d.active ? "v" : "") + '">' + pc + '<td class="nm">' + ck(d) + esc(d.name) + '</td><td>' +
        esc((d.height || "–").replace(/,/g, " / ")) + '</td><td>' +
        (d.dates.length ? d.dates.map(esc).join("<br>") : '<span class="none">–</span>') + '</td><td>' +
        (d.url ? '<a href="' + esc(d.url) + '" target="_blank" rel="noopener">MAP</a>' : '<span class="none">–</span>') + '</td></tr>';
    });
    html += '</tbody></table></div></section>';
  });
  document.getElementById("out").innerHTML = html;
}
render();
</script>
</body>
</html>
'''

html = TEMPLATE.replace("__DATA__", data)
with open(dst, "w", encoding="utf-8") as fh:
    fh.write(html)
print(f"OK: {len(rows)} falls -> {os.path.abspath(dst)}")