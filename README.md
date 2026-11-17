<!DOCTYPE html>
<html lang="zh-HK">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>計劃書回報分析</title>
<script src="https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
<style>
  :root{--navy:#0f172a;--navy2:#1e293b;--gold:#d97706;--gold2:#f59e0b;--soft:#fffbeb;--line:#e2e8f0;--mut:#64748b;--bg:#f1f5f9;--red:#dc2626}
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,"PingFang HK","Microsoft JhengHei",sans-serif;background:var(--bg);color:var(--navy);-webkit-font-smoothing:antialiased}
  header{background:radial-gradient(circle at 85% 0%,#334155 0%,transparent 55%),linear-gradient(135deg,#0f172a,#1e293b);color:#fff;padding:40px 20px 70px}
  .hd{max-width:1100px;margin:0 auto}
  .eyebrow{display:inline-block;font-size:12px;letter-spacing:2px;color:var(--gold2);border:1px solid rgba(245,158,11,.5);border-radius:999px;padding:4px 12px;margin-bottom:14px}
  header h1{margin:0 0 8px;font-size:30px;letter-spacing:1px}
  header p{margin:0;color:#cbd5e1;font-size:13px}
  main{max-width:1100px;margin:-44px auto 0;padding:0 16px 20px}
  section{background:#fff;border-radius:18px;padding:22px;margin-bottom:20px;box-shadow:0 10px 30px rgba(15,23,42,.07);border:1px solid rgba(226,232,240,.8)}
  section.res{border-top:4px solid var(--gold2)}
  h2{margin:0 0 14px;font-size:18px}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:14px;margin-bottom:16px}
  label{font-size:12px;color:var(--mut);font-weight:600;display:block;margin-bottom:5px}
  input[type=text],select,textarea{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:10px;font-size:14px;font-family:inherit;background:#f8fafc;color:var(--navy)}
  input[type=text]:focus,select:focus,textarea:focus{outline:0;border-color:var(--gold2);background:#fff;box-shadow:0 0 0 3px rgba(245,158,11,.18)}
  textarea{height:200px;font-family:Menlo,Consolas,monospace;font-size:12px}
  .bar{display:flex;flex-wrap:wrap;gap:10px;align-items:center}
  .btn{background:linear-gradient(135deg,#f59e0b,#d97706);color:#fff;border:0;border-radius:999px;padding:10px 18px;font-size:14px;font-weight:600;cursor:pointer;box-shadow:0 4px 12px rgba(217,119,6,.25)}
  .btn.dark{background:var(--navy);box-shadow:0 4px 12px rgba(15,23,42,.25)}
  .btn.sm{padding:7px 14px;font-size:13px}
  .cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:12px}
  .card{background:#f8fafc;border:1px solid var(--line);border-radius:14px;padding:14px 16px}
  .card small{color:var(--mut);font-size:12px;font-weight:600}
  .card b{display:block;font-size:21px;margin-top:6px}
  .card.hi{background:linear-gradient(135deg,#0f172a,#1e293b);border-color:transparent}
  .card.hi small{color:#cbd5e1}
  .card.hi b{color:var(--gold2)}
  .chips{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 16px}
  .chip{font-size:13px;background:#fef3c7;color:#92400e;border-radius:999px;padding:6px 12px;font-weight:600}
  .chip b{color:var(--navy)}
  .tip{background:var(--soft);border:1px solid #fde68a;border-radius:10px;padding:10px 14px;font-size:13px;margin-bottom:14px;line-height:1.8}
  .scroll{overflow:auto;border:1px solid var(--line);border-radius:14px}
  .tbl{border-collapse:separate;border-spacing:0;width:100%;min-width:760px;font-size:14px}
  .tbl th{background:var(--navy2);color:#fff;padding:12px 8px;text-align:center;font-weight:600;line-height:1.5;position:sticky;top:0}
  .tbl td{padding:11px 8px;text-align:center;border-bottom:1px solid var(--line);white-space:nowrap}
  .tbl tbody tr:nth-child(even) td{background:#f8fafc}
  .tbl td.k{font-weight:700}
  .tbl tbody tr td.g{background:#fff7ed;color:#9a3412;font-weight:800}
  .tbl td.neg{color:var(--red)}
  .full{max-height:480px}
  details{margin-top:16px}
  summary{cursor:pointer;font-weight:700;font-size:14px;margin-bottom:10px}
  .lg{display:flex;flex-wrap:wrap;gap:14px;font-size:12px;color:var(--mut);margin-top:8px}
  .lg i{display:inline-block;width:18px;height:4px;border-radius:2px;margin-right:6px;vertical-align:middle}
  .note{font-size:12px;color:var(--mut);line-height:1.9}
  footer{font-size:12px;color:var(--mut);text-align:center;padding:6px 20px 34px}
</style>
</head>
<body>
<header>
  <div class="hd">
    <span class="eyebrow">POLICY RETURN ANALYZER</span>
    <h1>計劃書回報分析</h1>
    <p>匯入計劃書，自動計算單利、複利、回報及倍數，並匯出 Excel 給客戶。</p>
  </div>
</header>
<main>
  <section>
    <div class="grid">
      <div><label>客戶稱呼</label><input type="text" id="client" placeholder="例如：李先生"></div>
      <div><label>計劃名稱</label><input type="text" id="plan" placeholder="匯入後自動填寫"></div>
      <div><label>保單貨幣</label><select id="cur"><option>US$</option><option>HK$</option><option>RMB¥</option></select></div>
      <div><label>投保年齡</label><input type="text" id="age" inputmode="numeric" value="0"></div>
      <div><label>重點年度（以逗號分隔）</label><input type="text" id="ms" value="10,20,30,40,50,60,70,80,90,100"></div>
    </div>
    <div class="bar">
      <input type="file" id="pdf" accept="application/pdf" style="display:none">
      <button class="btn" id="pick">📄 匯入計劃書 PDF</button>
      <button class="btn dark" id="exp">⬇ 匯出 Excel</button>
      <span id="msg" style="font-size:13px;color:#b45309"></span>
    </div>
    <details id="dataBox">
      <summary>📋 數據（可核對、修改或手動貼上）</summary>
      <div class="grid">
        <div><label>不提取方案：每行「保單年度　累計已繳保費　退保總額」</label><textarea id="tNw" spellcheck="false"></textarea></div>
        <div><label>提取方案：每行「保單年度　累計已繳保費　當年提取　提取後退保總額」</label><textarea id="tWd" spellcheck="false"></textarea></div>
      </div>
      <button class="btn sm" id="calc">重新計算</button>
    </details>
  </section>
  <div id="out"></div>
</main>
<footer>本工具只供參考，所有數字以保險公司發出的計劃書為準，非保證金額並非保證。PDF 只在你的瀏覽器內讀取，不會上傳。</footer>

<script>
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const num = s => { const v = parseFloat(String(s == null ? '' : s).replace(/[,\s$]/g, '')); return isNaN(v) ? 0 : v; };
const f0 = v => Math.round(v).toLocaleString('en-US');
const pc = v => v == null ? '—' : (v * 100).toFixed(2) + '%';
const mx = v => v == null ? '—' : v.toFixed(2) + ' 倍';
const cur = () => $('cur').value;
let A = { nw: [], wd: [] };

const NOTES = [
  '倍數 ＝ 總利益 ÷ 累計已繳保費。不提取方案的總利益是退保總額；提取方案的總利益是累計提取金額加提取後的保單餘額。',
  '單利（年均）＝（倍數 − 1）÷ 保單年度。',
  '複利（IRR）即內部回報率：假設保費在每個保單年度開始時繳付，提取金額及退保總額在保單年度終結時收取。',
  '退保總額包括非保證金額（復歸紅利及終期分紅），實際金額可能較高或較低。'
];
const DISCLAIMER = '本文件只供參考，並非保險合約的一部分。所有數字以保險公司發出的計劃書為準，非保證金額並非保證。';

// ---------- 讀取 PDF ----------
async function readPdf(file) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer(), cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/', cMapPacked: true }).promise;
  const lines = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();
    const its = tc.items.filter(x => x.str && x.str.trim()).map(x => ({ s: x.str, x: x.transform[4], y: x.transform[5] }));
    its.sort((a, b) => b.y - a.y || a.x - b.x);
    let row = null;
    its.forEach(it => {
      if (!row || Math.abs(row.y - it.y) > 3) { row = { y: it.y, a: [] }; lines.push(row); }
      row.a.push(it);
    });
  }
  return lines.map(r => r.a.sort((a, b) => a.x - b.x).map(x => x.s).join(' '));
}

// 從文字行找出數字表格，並按數字規律辨認各欄（不依賴欄位次序）
function findTables(lines) {
  const NUMRE = /^-?\d[\d,]*(\.\d+)?$/, buckets = {}, tables = [];
  lines.forEach(l => {
    const nums = l.split(/\s+/).filter(t => NUMRE.test(t)).map(t => parseFloat(t.replace(/,/g, '')));
    if (nums.length >= 6 && nums.length <= 16) (buckets[nums.length] = buckets[nums.length] || []).push(nums);
  });
  const isSeq = v => v.every((x, i) => i === 0 || x === 1 || (x - v[i - 1] >= 1 && x - v[i - 1] <= 3));
  Object.keys(buckets).forEach(key => {
    const R = buckets[key], L = +key, idx = [];
    for (let c = 0; c < L; c++) idx.push(c);
    const yc = idx.find(c => R[0][c] === 1 && isSeq(R.map(r => r[c])));
    if (yc === undefined) return;
    const segs = []; let seg = [];
    R.forEach((r, i) => { if (i > 0 && r[yc] === 1) { segs.push(seg); seg = []; } seg.push(r); });
    segs.push(seg);
    segs.forEach(S => {
      const n = S.length; if (n < 5) return;
      const col = c => S.map(r => r[c]), mid = Math.floor(n / 2);
      const seq = idx.filter(c => isSeq(col(c)));
      // 累計已繳保費：首年大於 0、只升不跌、後段不變
      const pcs = idx.filter(c => !seq.includes(c)).filter(c => { const v = col(c); return v[0] > 0 && v[n - 1] === v[mid] && v.every((x, i) => i === 0 || x >= v[i - 1]); });
      if (!pcs.length) return;
      const pm = pcs.reduce((a, b) => S[n - 1][b] < S[n - 1][a] ? b : a);
      const rest = idx.filter(c => !seq.includes(c) && c !== pm);
      // 退保總額：等於另外三欄（保證現金價值＋復歸紅利＋終期分紅）之和
      let T = -1, tri = [];
      rest.forEach(t => {
        if (Math.max.apply(null, col(t)) <= 0) return;
        if (T >= 0 && S[n - 1][t] <= S[n - 1][T]) return;
        const o = rest.filter(c => c !== t);
        for (let a = 0; a < o.length; a++) for (let b = a + 1; b < o.length; b++) for (let c = b + 1; c < o.length; c++) {
          if (S.every(r => Math.abs(r[o[a]] + r[o[b]] + r[o[c]] - r[t]) <= 2)) { T = t; tri = [o[a], o[b], o[c]]; return; }
        }
      });
      let wc = -1;
      if (T < 0) {
        const mxv = Math.max.apply(null, rest.map(c => S[n - 1][c]));
        const cand = rest.filter(c => S[n - 1][c] === mxv);
        const sum = c => col(c).reduce((a, b) => a + b, 0);
        if (!cand.length) return;
        T = cand.reduce((a, b) => sum(b) < sum(a) ? b : a);
      } else if (L <= 10) {
        const w = rest.filter(c => c !== T && !tri.includes(c)).find(c => { const v = col(c); return Math.min.apply(null, v) === 0 && Math.max.apply(null, v) > 0; });
        if (w !== undefined) wc = w;
      }
      tables.push({ hasWd: wc >= 0, rows: S.map(r => ({ y: r[yc], prem: r[pm], wd: wc >= 0 ? r[wc] : 0, val: r[T] })) });
    });
  });
  return tables.sort((a, b) => b.rows.length - a.rows.length);
}

const toTxt = (rows, hasWd) => rows.map(r => (hasWd ? [r.y, r.prem, r.wd, r.val] : [r.y, r.prem, r.val]).join('\t')).join('\n');
function parseTxt(txt, hasWd) {
  const need = hasWd ? 4 : 3, out = [];
  String(txt).split(/\n/).forEach(l => {
    const a = l.trim().split(/[\t ;；|]+/).map(t => t.replace(/,/g, '')).filter(t => /^-?\d+(\.\d+)?$/.test(t)).map(Number);
    if (a.length >= need && a[0] >= 1) out.push(hasWd ? { y: a[0], prem: a[1], wd: a[2], val: a[3] } : { y: a[0], prem: a[1], wd: 0, val: a[2] });
  });
  return out.sort((a, b) => a.y - b.y);
}

// ---------- 計算 ----------
function irr(fl) {
  const f = r => fl.reduce((s, x) => s + x[1] / Math.pow(1 + r, x[0]), 0);
  let lo = -0.99, hi = 1;
  if (!(f(lo) > 0 && f(hi) < 0)) return null;
  for (let k = 0; k < 80; k++) { const m = (lo + hi) / 2; if (f(m) > 0) lo = m; else hi = m; }
  return (lo + hi) / 2;
}
function analyze(rows, age) {
  const out = [], fl = []; let prev = 0, W = 0;
  rows.forEach(r => {
    const inc = r.prem - prev; prev = r.prem;
    if (inc) fl.push([r.y - 1, -inc]);
    if (r.wd) { fl.push([r.y, r.wd]); W += r.wd; }
    const ben = W + r.val, N = r.y, mult = r.prem > 0 ? ben / r.prem : null;
    out.push({ y: N, age: age + N, prem: r.prem, wd: r.wd, W, val: r.val, ben, gain: ben - r.prem, mult,
      simple: mult == null ? null : (mult - 1) / N,
      irr: (r.prem > 0 && ben > 0) ? irr(fl.concat([[N, r.val]])) : null });
  });
  return out;
}
function baseInfo() {
  const b = A.nw.length ? A.nw : A.wd;
  let pay = 0, prev = 0, total = 0;
  b.forEach(r => { if (r.prem > prev) pay++; prev = r.prem; total = Math.max(total, r.prem); });
  const be = b.find(r => r.ben >= r.prem);
  return { annual: b[0].prem, pay, total, be };
}
function miles(rows) {
  const set = String($('ms').value).split(/[,，\s]+/).map(Number).filter(x => x > 0);
  const m = rows.filter(r => set.includes(r.y));
  return m.length ? m : rows.filter((r, i) => r.y % 10 === 0 || i === rows.length - 1);
}
function multChips() {
  const out = [];
  [1, 2, 3, 4, 5, 6, 7, 8, 10, 20, 50, 100].forEach(m => {
    const r = A.nw.find(x => x.mult != null && x.mult >= m);
    if (r) out.push({ m, y: r.y, age: r.age });
  });
  return out;
}
function wdText() {
  const w = A.wd.filter(r => r.wd > 0);
  if (!w.length) return '';
  const e = A.wd[A.wd.length - 1], c = cur();
  return '由第 ' + w[0].y + ' 年（' + w[0].age + ' 歲）起提取，至第 ' + w[w.length - 1].y + ' 年共提取 ' + w.length + ' 次，累計提取 ' + c + f0(e.W) +
    '；第 ' + e.y + ' 年保單餘額 ' + c + f0(e.val) + '，總利益 ' + c + f0(e.ben) + '（已繳保費的 ' + mx(e.mult) + '）。';
}

// 欄位定義：[標題, 取值, 類型]；n=數字 m=金額 x=倍數 p=百分比
const C_Y = ['保單年度', r => r.y, 'n'], C_A = ['年齡', r => r.age, 'n'], C_P = ['累計已繳保費', r => r.prem, 'm'];
const C_G = ['總回報', r => r.gain, 'm'], C_X = ['倍數', r => r.mult, 'x'], C_S = ['單利（年均）', r => r.simple, 'p'], C_I = ['複利（IRR）', r => r.irr, 'p'];
const COLS_NW = [C_Y, C_A, C_P, ['退保總額', r => r.val, 'm'], C_G, C_X, C_S, C_I];
const COLS_WD = [C_Y, C_A, C_P, ['當年提取', r => r.wd, 'm'], ['累計提取', r => r.W, 'm'], ['提取後保單餘額', r => r.val, 'm'], ['總利益（累計提取＋餘額）', r => r.ben, 'm'], C_G, C_X, C_S, C_I];
const show = (v, t) => t === 'n' ? String(v) : t === 'm' ? cur() + f0(v) : t === 'x' ? mx(v) : pc(v);

// ---------- 畫面 ----------
function tbl(cols, rows, full) {
  return '<div class="scroll' + (full ? ' full' : '') + '"><table class="tbl"><thead><tr>' + cols.map(c => '<th>' + c[0] + '</th>').join('') + '</tr></thead><tbody>' +
    rows.map(r => '<tr>' + cols.map((c, k) => {
      const v = c[1](r);
      const cls = k === 0 ? 'k' : (c[2] === 'x' || c[2] === 'p') ? 'g' : (c[2] === 'm' && v < 0) ? 'neg' : '';
      return '<td' + (cls ? ' class="' + cls + '"' : '') + '>' + show(v, c[2]) + '</td>';
    }).join('') + '</tr>').join('') + '</tbody></table></div>';
}
function chart() {
  const b = A.nw.length ? A.nw : A.wd, S = [['累計已繳保費', '#94a3b8', b.map(r => [r.y, r.prem])]];
  if (A.nw.length) S.push(['不提取｜退保總額', '#d97706', A.nw.map(r => [r.y, r.val])]);
  if (A.wd.length) S.push(['提取｜累計提取＋保單餘額', '#0f172a', A.wd.map(r => [r.y, r.ben])]);
  let X1 = 0, Y1 = 0;
  S.forEach(s => s[2].forEach(p => { X1 = Math.max(X1, p[0]); Y1 = Math.max(Y1, p[1]); }));
  if (X1 <= 1 || Y1 <= 0) return '';
  const W = 700, H = 260, L = 10, T = 26, B = 24;
  const X = y => L + (y - 1) / (X1 - 1) * (W - 2 * L), Y = v => H - B - v / Y1 * (H - T - B);
  let g = '';
  for (let k = 0; k <= 4; k++) { const yy = Y(Y1 * k / 4).toFixed(1); g += '<line x1="' + L + '" x2="' + (W - L) + '" y1="' + yy + '" y2="' + yy + '" stroke="#e2e8f0"/>'; }
  S.forEach(s => { g += '<polyline fill="none" stroke="' + s[1] + '" stroke-width="2.5" stroke-linejoin="round" points="' + s[2].map(p => X(p[0]).toFixed(1) + ',' + Y(p[1]).toFixed(1)).join(' ') + '"/>'; });
  g += '<text x="' + L + '" y="16" font-size="12" fill="#64748b">最高 ' + cur() + f0(Y1) + '</text>' +
    '<text x="' + L + '" y="' + (H - 6) + '" font-size="12" fill="#64748b">第 1 年</text>' +
    '<text x="' + (W - L) + '" y="' + (H - 6) + '" font-size="12" fill="#64748b" text-anchor="end">第 ' + X1 + ' 年</text>';
  return '<section><h2>價值走勢</h2><svg viewBox="0 0 ' + W + ' ' + H + '" style="width:100%;height:auto">' + g + '</svg><div class="lg">' +
    S.map(s => '<span><i style="background:' + s[1] + '"></i>' + s[0] + '</span>').join('') + '</div></section>';
}
function render() {
  const age = Math.round(num($('age').value));
  A.nw = analyze(parseTxt($('tNw').value, false), age);
  A.wd = analyze(parseTxt($('tWd').value, true), age);
  if (!A.nw.length && !A.wd.length) {
    $('out').innerHTML = '<section><div class="tip" style="margin:0">請按「匯入計劃書 PDF」，或展開「數據」手動貼上數字，然後按「重新計算」。</div></section>';
    return;
  }
  const c = cur(), I = baseInfo();
  let h = '<section class="res"><h2>保單概覽</h2><div class="cards">' +
    '<div class="card"><small>年繳保費</small><b>' + c + f0(I.annual) + '</b></div>' +
    '<div class="card"><small>繳費年期</small><b>' + I.pay + ' 年</b></div>' +
    '<div class="card hi"><small>總保費</small><b>' + c + f0(I.total) + '</b></div>' +
    '<div class="card hi"><small>回本年度</small><b>' + (I.be ? '第 ' + I.be.y + ' 年' : '—') + '</b></div></div></section>';
  h += chart();
  if (A.nw.length) {
    h += '<section class="res"><h2>不提取方案（複利累積）</h2><div class="chips">' +
      multChips().map(x => '<span class="chip">' + (x.m === 1 ? '回本' : x.m + ' 倍') + '｜<b>第 ' + x.y + ' 年（' + x.age + ' 歲）</b></span>').join('') + '</div>' +
      tbl(COLS_NW, miles(A.nw)) + '<details><summary>查看逐年數字</summary>' + tbl(COLS_NW, A.nw, true) + '</details></section>';
  }
  if (A.wd.length) {
    const t = wdText();
    h += '<section class="res"><h2>提取方案</h2>' + (t ? '<div class="tip">' + t + '</div>' : '') +
      tbl(COLS_WD, miles(A.wd)) + '<details><summary>查看逐年數字</summary>' + tbl(COLS_WD, A.wd, true) + '</details></section>';
  }
  h += '<section><h2>計算說明</h2><div class="note">' + NOTES.map(x => '・' + x).join('<br>') + '</div></section>';
  $('out').innerHTML = h;
}

// ---------- 事件 ----------
$('pick').addEventListener('click', () => $('pdf').click());
$('pdf').addEventListener('change', async ev => {
  const f = ev.target.files[0]; if (!f) return;
  if (!window.pdfjsLib) { $('msg').textContent = 'PDF 元件未能載入，請檢查網絡後重新整理。'; return; }
  $('msg').textContent = '讀取中…';
  try {
    const lines = await readPdf(f), tables = findTables(lines);
    const nw = tables.find(t => !t.hasWd), wd = tables.find(t => t.hasWd);
    $('tNw').value = nw ? toTxt(nw.rows, false) : '';
    $('tWd').value = wd ? toTxt(wd.rows, true) : '';
    const all = lines.join('\n');
    const pm = all.match(/建[议議][书書]摘要[：:]\s*([^\n]+)/);
    if (pm) $('plan').value = pm[1].trim();
    const cm = all.match(/保[单單][货貨][币幣][：:]\s*(美元|港元|人民[币幣])/);
    if (cm) $('cur').value = cm[1] === '美元' ? 'US$' : cm[1] === '港元' ? 'HK$' : 'RMB¥';
    if (nw || wd) {
      $('msg').textContent = '已讀取：不提取方案 ' + (nw ? nw.rows.length + ' 年' : '找不到') + '；提取方案 ' + (wd ? wd.rows.length + ' 年' : '找不到') + '。請對照計劃書核對數字。';
    } else {
      $('msg').textContent = '未能從此 PDF 辨認表格，請展開「數據」手動貼上。';
      $('dataBox').open = true;
    }
    render();
  } catch (e) { $('msg').textContent = '讀取失敗：' + e.message; }
  ev.target.value = '';
});
$('calc').addEventListener('click', render);
['age', 'ms'].forEach(id => $(id).addEventListener('input', render));
$('cur').addEventListener('change', render);

// ---------- 匯出 Excel ----------
const XC = { navy: 'FF0F172A', navy2: 'FF1E293B', soft: 'FFFEF3C7', light: 'FFF8FAFC', line: 'FFCBD5E1', pay: 'FFFFF7ED', payTxt: 'FF9A3412', red: 'FFDC2626', mut: 'FF64748B', white: 'FFFFFFFF' };
const XHEAD = { bold: true, color: XC.white, fill: XC.navy2 };
function xCell(ws, r, c1, c2, v, o) {
  o = o || {};
  const s = { style: 'thin', color: { argb: XC.line } };
  for (let c = c1; c <= c2; c++) {
    const cell = ws.getCell(r, c);
    if (c === c1) cell.value = v;
    cell.font = { name: 'Microsoft JhengHei', size: o.size || 11, bold: !!o.bold, italic: !!o.italic, color: { argb: o.color || XC.navy } };
    cell.alignment = { horizontal: o.align || 'center', vertical: 'middle', wrapText: true };
    if (o.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: o.fill } };
    if (!o.noBorder) cell.border = { top: s, left: s, bottom: s, right: s };
    if (o.fmt) cell.numFmt = o.fmt;
  }
  if (c2 > c1) ws.mergeCells(r, c1, r, c2);
}
function xFmt(t) {
  const q = '"' + cur().replace(/"/g, '') + '"';
  return t === 'm' ? q + '#,##0;[Red]-' + q + '#,##0' : t === 'x' ? '0.00"倍"' : t === 'p' ? '0.00%' : '0';
}
// 寫入一個表格，回傳下一個空白列
function xTable(ws, r, cols, rows) {
  cols.forEach((c, k) => xCell(ws, r, k + 1, k + 1, c[0], XHEAD));
  ws.getRow(r).height = 36; r++;
  rows.forEach((row, i) => {
    const z = i % 2 ? XC.light : undefined;
    cols.forEach((c, k) => {
      const v = c[1](row), hi = c[2] === 'x' || c[2] === 'p';
      xCell(ws, r, k + 1, k + 1, v == null ? '—' : v, { fmt: xFmt(c[2]), bold: k === 0 || hi, fill: hi ? XC.pay : z, color: hi ? XC.payTxt : undefined });
    });
    ws.getRow(r).height = 22; r++;
  });
  return r;
}
function xToday() { const d = new Date(); return d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日'; }
function xTitle(ws, title, NC) {
  xCell(ws, 1, 1, NC, title, { bold: true, size: 16, color: XC.white, fill: XC.navy, noBorder: true });
  ws.getRow(1).height = 38;
  xCell(ws, 2, 1, NC, '製表日期：' + xToday(), { size: 9, color: XC.mut, fill: XC.light, noBorder: true });
  ws.getRow(2).height = 20;
}
function xSummary(wb) {
  const NC = 11, ws = wb.addWorksheet('方案摘要', { views: [{ showGridLines: false }] });
  ws.columns = Array.from({ length: NC }, () => ({ width: 17 }));
  ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  const client = $('client').value.trim(), plan = $('plan').value.trim(), I = baseInfo();
  xTitle(ws, (client ? client + '｜' : '') + '保單回報分析', NC);
  let r = 3;
  const sec = t => { r++; xCell(ws, r, 1, NC, t, { bold: true, size: 12, fill: XC.soft, align: 'left' }); ws.getRow(r).height = 26; r++; };
  const info = (label, v, fmt, h) => {
    xCell(ws, r, 1, 3, label, { bold: true, fill: XC.light, align: 'left' });
    xCell(ws, r, 4, NC, v, { align: 'left', fmt });
    ws.getRow(r).height = h || 22; r++;
  };
  sec('一、保單資料');
  if (client) info('客戶', client);
  if (plan) info('計劃名稱', plan);
  info('保單貨幣', cur());
  info('投保年齡', Math.round(num($('age').value)) + ' 歲');
  info('年繳保費', I.annual, xFmt('m'));
  info('繳費年期', I.pay + ' 年');
  info('總保費', I.total, xFmt('m'));
  if (I.be) info('回本年度', '第 ' + I.be.y + ' 年（' + I.be.age + ' 歲）');
  if (A.nw.length) {
    sec('二、不提取方案（複利累積）｜重點年度');
    const ch = multChips().filter(x => x.m > 1).map(x => x.m + ' 倍：第 ' + x.y + ' 年（' + x.age + ' 歲）').join('　｜　');
    if (ch) info('資金倍增時間', ch, undefined, 40);
    r = xTable(ws, r, COLS_NW, miles(A.nw));
  }
  if (A.wd.length) {
    sec((A.nw.length ? '三' : '二') + '、提取方案｜重點年度');
    const t = wdText();
    if (t) info('提取安排', t, undefined, 40);
    r = xTable(ws, r, COLS_WD, miles(A.wd));
  }
  sec('計算說明');
  NOTES.concat([DISCLAIMER]).forEach(t => { xCell(ws, r, 1, NC, t, { size: 9, color: XC.mut, align: 'left', noBorder: true }); ws.getRow(r).height = 20; r++; });
}
function xFull(wb, name, title, cols, rows) {
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 4, showGridLines: false }] });
  ws.columns = cols.map(() => ({ width: 17 }));
  ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  xTitle(ws, title, cols.length);
  xTable(ws, 4, cols, rows);
}
$('exp').addEventListener('click', async () => {
  if (!window.ExcelJS) { alert('Excel 元件未能載入，請檢查網絡後重新整理頁面再試。'); return; }
  if (!A.nw.length && !A.wd.length) { alert('請先匯入計劃書或貼上數據。'); return; }
  const b = $('exp'), old = b.textContent;
  b.textContent = '匯出中…';
  try {
    const wb = new ExcelJS.Workbook();
    xSummary(wb);
    if (A.nw.length) xFull(wb, '不提取方案（逐年）', '不提取方案｜逐年回報分析', COLS_NW, A.nw);
    if (A.wd.length) xFull(wb, '提取方案（逐年）', '提取方案｜逐年回報分析', COLS_WD, A.wd);
    const buf = await wb.xlsx.writeBuffer();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    a.download = '保單回報分析.xlsx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 3000);
  } catch (e) { alert('匯出失敗：' + e.message); }
  b.textContent = old;
});

render();
</script>
</body>
</html>
