/* 南山発注 ― 店舗用 発注ページ
   ・品目は 発注品目リスト_*.csv から読み込む（Excelで編集 → 「CSV UTF-8」で保存）
   ・発注内容はサーバーに送らない。入力途中の内容だけ端末に一時保存し、送信後に消す。 */
'use strict';

// ===== 設定 =====
const STORES = [
  { name: '南山亭上田本店',   short: '上田本店' },
  { name: '南山亭長野東口店', short: '長野東口店' },
  { name: '南山亭松本駅前店', short: '松本駅前店' },
  { name: '南山亭佐久平店',   short: '佐久平店' },
  { name: 'NAMSAN上田秋和店', short: 'NAMSAN秋和店' },
];

const SHEETS = [
  {
    id: 'center',
    label: 'センター発注書',
    to: '本社（セントラルキッチン）へ',
    title: 'センター発注書',
    file: 'センター発注書',
    csv: '発注品目リスト_センター.csv',
    layout: 'center',
    orientation: 'portrait',
    memoLabel: '連絡事項',
    excludeStores: [],
  },
  {
    id: 'supplies',
    label: '本社消耗品 発注表',
    to: '本社（セントラルキッチン）へ',
    title: '本社消耗品 発注表',
    file: '消耗品',
    csv: '発注品目リスト_消耗品.csv',
    layout: 'supplies',
    orientation: 'landscape',
    memoLabel: '備考欄',
    caution: '（注）レンタル品については、利用後は本社に返却をお願いします。\nキッチンシューズ（かかと無し）は注文を頂いてからの発注になるので、納品まで時間がかかります。',
    excludeStores: [],
  },
  {
    id: 'honten',
    label: '南山本店発注書',
    to: '上田本店へ',
    title: '南山亭本店発注書',
    file: '本店発注書',
    csv: '発注品目リスト_本店.csv',
    layout: 'honten',
    orientation: 'landscape',
    memoLabel: '連絡事項',
    excludeStores: ['南山亭上田本店'],   // 上田本店は自店に発注しない
  },
];

// 発注時に選べる単位（全品目共通）。増やす・減らすときはここを直す
const UNITS = ['kg', 'c/s', 'p/c', '本', '枚', '個', '缶', '袋', '箱', '束', '足', 'セット'];

const BIG_QTY = 100;          // この数量以上は確認メッセージを出す
const PDF_SCALE = 2.5;        // PDF画像の解像度（大きいほどきれい・重い）
const DRAFT_PREFIX = 'namsan-hacchu-draft-';

// ===== 状態 =====
const itemsBySheet = {};      // id -> [{key, block, name, kind, unit, noteLabel}]
const loadErrors = {};
let current = null;           // 表示中の用紙
let draft = null;             // 入力途中の内容
let renderedFormId = null;
let pdfFile = null, pdfUrl = null;

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));

// ===== 文字の整形 =====
const toHalf = s => String(s ?? '').replace(/[０-９．]/g, c => c === '．' ? '.' : String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
const kana = s => s.normalize('NFKC').toLowerCase().replace(/[ぁ-ゖ]/g, c => String.fromCharCode(c.charCodeAt(0) + 0x60));
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const safeFile = s => String(s).replace(/[\\/:*?"<>|\s]+/g, '');

function parseQty(v) {
  const s = toHalf(v).trim();
  if (s === '') return { ok: true, n: 0 };
  if (!/^\d+(\.\d+)?$/.test(s)) return { ok: false, n: 0 };
  return { ok: true, n: Number(s) };
}
function fmtNum(n) { return String(Math.round(n * 1000) / 1000); }

// ===== CSV =====
function parseCSV(text) {
  text = text.replace(/^﻿/, '');
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; }
      else f += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(f); f = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(f); rows.push(row); row = []; f = '';
    } else f += c;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows.filter(r => r.some(x => x.trim() !== ''));
}

async function loadSheet(sheet) {
  const res = await fetch(encodeURI(sheet.csv), { cache: 'no-store' });
  if (!res.ok) throw new Error(`${sheet.csv} が見つかりません（${res.status}）`);
  const buf = await res.arrayBuffer();
  let text;
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(buf); }
  catch { text = new TextDecoder('shift_jis').decode(buf); }   // 普通のCSV保存でも読めるように
  const rows = parseCSV(text);
  if (!rows.length) throw new Error(`${sheet.csv} が空です`);
  const head = rows[0].map(h => h.trim());
  const col = name => head.indexOf(name);
  const iName = col('品名');
  if (iName < 0) throw new Error(`${sheet.csv} の1行目に「品名」の列がありません`);
  const iBlock = col('ブロック'), iUnit = col('単位'), iSpec = col('規格'), iKind = col('種類'), iNote = col('記入欄');
  const get = (r, i) => (i >= 0 && r[i] != null ? r[i].trim() : '');
  const seen = {};
  return rows.slice(1).map(r => {
    const name = get(r, iName);
    if (!name) return null;
    const it = {
      block: get(r, iBlock) || '品目',
      name,
      kind: get(r, iKind),
      unit: iSpec >= 0 ? get(r, iSpec) : get(r, iUnit),
      noteLabel: get(r, iNote),
    };
    let key = [it.name, it.kind, it.unit].join('|');
    seen[key] = (seen[key] || 0) + 1;
    if (seen[key] > 1) key += '#' + seen[key];
    it.key = key;
    return it;
  }).filter(Boolean);
}

// ===== 一時保存 =====
const emptyDraft = () => ({ store: '', date: '', person: '', memo: '', qty: {}, note: {}, unit: {} });

// 単位：センター・本店はCSVの単位が初期値。消耗品は「規格どおり」（空）が初期値
const isSupplies = () => current.layout === 'supplies';
function unitOf(it) {
  const u = draft.unit[it.key];
  if (u != null && u !== '') return u;
  return isSupplies() ? '' : it.unit;
}
function unitOptions(it) {
  const list = UNITS.slice();
  if (!isSupplies() && it.unit && !list.includes(it.unit)) list.unshift(it.unit);
  const sel = unitOf(it);
  return (isSupplies() ? `<option value="">規格どおり</option>` : '') +
    list.map(u => `<option value="${esc(u)}"${u === sel ? ' selected' : ''}>${esc(u)}</option>`).join('');
}
function readDraft(id) {
  try { const s = localStorage.getItem(DRAFT_PREFIX + id); if (s) return Object.assign(emptyDraft(), JSON.parse(s)); } catch {}
  return emptyDraft();
}
function writeDraft() {
  if (!current) return;
  try { localStorage.setItem(DRAFT_PREFIX + current.id, JSON.stringify(draft)); } catch {}
}
function clearDraft(id) { try { localStorage.removeItem(DRAFT_PREFIX + id); } catch {} }
function hasDraft(id) {
  const d = readDraft(id);
  return Object.values(d.qty).some(v => parseQty(v).n > 0) || d.memo || Object.values(d.note).some(Boolean);
}

// ===== 画面切り替え（#/ , #/form/ID , #/preview/ID） =====
function route() {
  const [, screen, id] = (location.hash || '#/').split('/');
  const sheet = SHEETS.find(s => s.id === id);
  if ((screen === 'form' || screen === 'preview') && sheet && itemsBySheet[sheet.id]) {
    if (!current || current.id !== sheet.id) { current = sheet; draft = readDraft(sheet.id); }
    if (screen === 'preview') {
      if (!validate(true)) { location.replace('#/form/' + sheet.id); return; }
      showPreview();
    } else showForm();
  } else {
    showHome();
  }
}
let formScroll = 0, shownId = null;
function show(id) {
  if (shownId === 'screen-form' && id !== 'screen-form') formScroll = window.scrollY;
  const back = id === 'screen-form' && shownId === 'screen-preview';
  $$('.screen').forEach(s => s.hidden = s.id !== id);
  shownId = id;
  window.scrollTo(0, back ? formScroll : 0);   // 確認画面から戻ったときは元の位置へ
}
document.addEventListener('click', e => {
  const a = e.target.closest('[data-action]');
  if (!a) return;
  if (a.dataset.action === 'home') location.hash = '#/';
  if (a.dataset.action === 'form' && current) location.hash = '#/form/' + current.id;
});

// ===== ホーム =====
function showHome() {
  const box = $('#sheet-buttons');
  box.innerHTML = SHEETS.map(s => `
    <button class="sheet-btn" data-sheet="${s.id}" ${itemsBySheet[s.id] ? '' : 'disabled'}>
      <strong>${esc(s.label)}</strong>
      <span>${esc(s.to)}</span>
      ${hasDraft(s.id) ? '<span class="draft">● 入力途中あり</span>' : ''}
    </button>`).join('');
  $$('.sheet-btn', box).forEach(b => b.onclick = () => location.hash = '#/form/' + b.dataset.sheet);
  const errs = Object.values(loadErrors);
  const err = $('#load-error');
  err.hidden = !errs.length;
  if (errs.length) err.textContent = '品目リストを読み込めませんでした。\n' + errs.join('\n') +
    (location.protocol === 'file:' ? '\n\nPCで確認するときは「PC確認用」フォルダーの「PCで確認.bat」から開いてください。' : '');
  show('screen-home');
}

// ===== 入力画面 =====
function showForm() {
  if (renderedFormId !== current.id) renderForm();
  show('screen-form');
}

function renderForm() {
  renderedFormId = current.id;
  $('#form-title').textContent = current.label;
  $('#form-to').textContent = current.to;
  $('#memo-label').textContent = current.memoLabel;

  const stores = STORES.filter(s => !current.excludeStores.includes(s.name));
  $('#f-store').innerHTML = '<option value="">選択してください</option>' +
    stores.map(s => `<option value="${esc(s.name)}">${esc(s.name)}</option>`).join('');
  if (!stores.some(s => s.name === draft.store)) draft.store = '';
  $('#f-store').value = draft.store;
  $('#f-date').value = draft.date;
  $('#f-person').value = draft.person;
  $('#f-memo').value = draft.memo;
  $('#f-search').value = '';
  $('#f-only').setAttribute('aria-pressed', 'false');

  const items = itemsBySheet[current.id];
  const blocks = groupBlocks(items);
  $('#item-list').innerHTML = blocks.map(b => `
    <section class="block">
      <h2>${esc(b.name)}</h2>
      <div class="rows">${b.items.map(itemRow).join('')}</div>
    </section>`).join('') + '<p class="empty" id="no-hit" hidden>該当する品目がありません</p>';
  $$('.item').forEach(refreshRow);
  updateCount();
  $$('.invalid').forEach(el => el.classList.remove('invalid'));
}

function itemRow(it) {
  const q = draft.qty[it.key] ?? '';
  return `
    <div class="item" data-key="${esc(it.key)}" data-search="${esc(kana(it.name + ' ' + it.kind))}">
      <div class="name">${esc(it.name)}${it.kind ? `<small>${esc(it.kind)}</small>` : ''}
        ${isSupplies() && it.unit ? `<span class="unit">${esc(it.unit)}</span>` : ''}
        <label class="unitsel">単位 <select aria-label="${esc(it.name)} 単位">${unitOptions(it)}</select></label></div>
      <div class="stepper">
        <button type="button" data-step="-1" aria-label="減らす">−</button>
        <input type="text" inputmode="decimal" value="${esc(q)}" aria-label="${esc(it.name)} 数量">
        <button type="button" data-step="1" aria-label="増やす">＋</button>
      </div>
      ${it.noteLabel ? `<div class="note"><input type="text" placeholder="${esc(it.noteLabel)}" value="${esc(draft.note[it.key] ?? '')}"></div>` : ''}
    </div>`;
}

function groupBlocks(items) {
  const map = new Map();
  items.forEach(it => { if (!map.has(it.block)) map.set(it.block, []); map.get(it.block).push(it); });
  return Array.from(map, ([name, items]) => ({ name, items }));
}

function refreshRow(row) {
  const v = $('.stepper input', row).value;
  row.classList.toggle('on', parseQty(v).n > 0);
}
function updateCount() {
  const n = Object.values(draft.qty).filter(v => parseQty(v).n > 0).length;
  $('#count-badge').textContent = n + '品目';
}

// 入力イベント
$('#item-list').addEventListener('input', e => {
  const row = e.target.closest('.item'); if (!row) return;
  const key = row.dataset.key;
  if (e.target.closest('.stepper')) { draft.qty[key] = e.target.value; refreshRow(row); updateCount(); }
  else if (e.target.closest('.note')) draft.note[key] = e.target.value;
  writeDraft();
});
$('#item-list').addEventListener('change', e => {
  const sel = e.target.closest('.unitsel select'); if (!sel) return;
  draft.unit[sel.closest('.item').dataset.key] = sel.value;
  writeDraft();
});
$('#item-list').addEventListener('click', e => {
  const b = e.target.closest('[data-step]'); if (!b) return;
  const row = b.closest('.item'), input = $('.stepper input', row);
  const cur = parseQty(input.value).n;
  const next = Math.max(0, cur + Number(b.dataset.step));
  input.value = next > 0 ? fmtNum(next) : '';
  draft.qty[row.dataset.key] = input.value;
  refreshRow(row); updateCount(); writeDraft();
});
[['#f-store', 'store'], ['#f-date', 'date'], ['#f-person', 'person'], ['#f-memo', 'memo']].forEach(([sel, k]) => {
  const el = $(sel);
  const h = () => { draft[k] = el.value; el.classList.remove('invalid'); writeDraft(); };
  el.addEventListener('input', h); el.addEventListener('change', h);
});

// 検索・「発注のみ」
function applyFilter() {
  const q = kana($('#f-search').value.trim());
  const only = $('#f-only').getAttribute('aria-pressed') === 'true';
  let hits = 0;
  $$('.block').forEach(b => {
    let vis = 0;
    $$('.item', b).forEach(r => {
      const ok = (!q || r.dataset.search.includes(q)) && (!only || r.classList.contains('on'));
      r.hidden = !ok; if (ok) vis++;
    });
    b.hidden = vis === 0; hits += vis;
  });
  $('#no-hit').hidden = hits > 0;
}
$('#f-search').addEventListener('input', applyFilter);
$('#f-only').addEventListener('click', e => {
  const b = e.currentTarget;
  b.setAttribute('aria-pressed', b.getAttribute('aria-pressed') === 'true' ? 'false' : 'true');
  applyFilter();
});

$('#btn-clear').addEventListener('click', () => {
  if (!confirm('入力した内容をすべて消します。よろしいですか？')) return;
  resetCurrent();
  renderForm();
});
function resetCurrent() {
  clearDraft(current.id);
  draft = emptyDraft();
  renderedFormId = null;
}

// ===== チェック =====
function orderedItems() {
  return itemsBySheet[current.id].filter(it => parseQty(draft.qty[it.key]).n > 0);
}
function validate(silent) {
  const miss = [];
  if (!draft.store) miss.push(['#f-store', '店舗']);
  if (!draft.date) miss.push(['#f-date', '納品日']);
  if (!draft.person.trim()) miss.push(['#f-person', '担当者']);
  const bad = itemsBySheet[current.id].filter(it => !parseQty(draft.qty[it.key]).ok);
  const ordered = orderedItems();
  if (silent) return !miss.length && !bad.length && ordered.length > 0;

  if (miss.length) {
    miss.forEach(([sel]) => $(sel).classList.add('invalid'));
    alert('次の項目を入力してください：\n' + miss.map(m => '・' + m[1]).join('\n'));
    window.scrollTo({ top: 0, behavior: 'smooth' });
    return false;
  }
  if (bad.length) {
    alert('数量が正しくありません（数字で入力してください）：\n' + bad.map(it => '・' + it.name + (it.kind ? ' ' + it.kind : '')).join('\n'));
    return false;
  }
  if (!ordered.length) { alert('発注する品目がありません。数量を入力してください。'); return false; }
  const big = ordered.filter(it => parseQty(draft.qty[it.key]).n >= BIG_QTY);
  if (big.length && !confirm(`数量が${BIG_QTY}以上の品目があります。\n` +
      big.map(it => `・${it.name}${it.kind ? ' ' + it.kind : ''}：${draft.qty[it.key]}${unitOf(it) ? ' ' + unitOf(it) : ''}`).join('\n') +
      '\n\nこのまま進みますか？')) return false;
  return true;
}
$('#btn-preview').addEventListener('click', () => {
  if (validate(false)) location.hash = '#/preview/' + current.id;
});

// ===== 発注書（A4）を作る =====
const SIZE = { portrait: [793, 1122], landscape: [1122, 793] };
const PAD_V = 52, PAD_H = 56, HEAD = 54;

function fmtDate(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const w = '日月火水木金土'[new Date(y, m - 1, d).getDay()];
  return { text: `${m}月${d}日（${w}）`, file: String(m).padStart(2, '0') + String(d).padStart(2, '0') };
}
function nowText() {
  const n = new Date(), p = x => String(x).padStart(2, '0');
  return `${n.getFullYear()}/${p(n.getMonth() + 1)}/${p(n.getDate())} ${p(n.getHours())}:${p(n.getMinutes())}`;
}
function fileName() {
  const st = STORES.find(s => s.name === draft.store);
  return safeFile(`${current.file}_${st ? st.short : draft.store}_${fmtDate(draft.date).file}納品_${draft.person.trim()}`) + '.pdf';
}

function buildSheet() {
  const [W, H] = SIZE[current.orientation];
  const el = document.createElement('div');
  el.className = 'sheet ' + current.orientation;
  el.innerHTML = `
    <div class="sh-made">作成 ${esc(nowText())}　発注 ${orderedItems().length}品目（網掛け・太字）</div>
    <div class="sh-head">
      <div class="sh-title">${esc(current.title)}</div>
      <div class="sh-fields">
        <span><b>${esc(draft.store)}</b></span>
        <span>納品日：<b>${esc(fmtDate(draft.date).text)}</b></span>
        <span>担当者：<b>${esc(draft.person.trim())}</b></span>
      </div>
    </div>`;
  const innerW = W - PAD_H, innerH = H - PAD_V - HEAD;
  const blocks = groupBlocks(itemsBySheet[current.id]);
  const memo = esc(draft.memo.trim());

  if (current.layout === 'center') {
    const memoH = 92;
    el.insertAdjacentHTML('beforeend', blocksHTML(blocks, innerW, innerH - memoH, unitCols(innerW, blocks.length)));
    el.insertAdjacentHTML('beforeend', `
      <div class="sh-memo" style="height:${memoH - 8}px"><div class="lbl">連絡事項</div><div class="txt">${memo}</div></div>`);
  } else if (current.layout === 'supplies') {
    const memoH = 66;
    el.insertAdjacentHTML('beforeend', blocksHTML(blocks, innerW, innerH - memoH, null));
    el.insertAdjacentHTML('beforeend', `
      <div class="sh-memo" style="height:${memoH - 8}px">
        <div class="lbl">備考欄</div><div class="txt">${memo}</div>
        <div class="txt sh-caution" style="flex:0 0 540px;border-left:1px solid #000;margin:0">${esc(current.caution).replace(/\n/g, '<br>')}</div>
      </div>`);
  } else { // honten
    const tableW = Math.round(innerW * 0.5);
    el.insertAdjacentHTML('beforeend', `
      <div class="sh-blocks" style="height:${innerH}px;align-items:stretch">
        ${blocksHTML(blocks, tableW, innerH, unitCols(tableW, 1), true)}
        <div class="sh-side-memo" style="flex:1"><div class="lbl">連絡事項</div><div class="txt">${memo}</div></div>
      </div>`);
  }
  return { el, W, H };
}

function unitCols(totalW, nBlocks) {
  const w = Math.floor((totalW - 8 * (nBlocks - 1)) / nBlocks);
  return { w, cols: [w - 4 - 40 - 46, 40, 46] };
}

// ブロックを横に並べた表
function blocksHTML(blocks, totalW, availH, uc, bare) {
  const supplies = current.layout === 'supplies';
  // 行の高さ：一番長いブロックが入るように
  const units = blocks.map(b => b.items.reduce((s, it) => s + (it.noteLabel ? 2 : 1), 0));
  const maxU = Math.max(...units);
  const rowH = Math.max(12, Math.min(30, Math.floor((availH - 26) / maxU)));
  const fs = Math.max(8, Math.min(15, Math.floor(rowH * 0.6)));
  const nB = blocks.length;
  const bw = uc ? uc.w : Math.floor((totalW - 8 * (nB - 1)) / nB);
  const tables = blocks.map(b => {
    let cols, headRow, body;
    if (supplies) {
      cols = [0, 54, 66, 44];   // 品名, 種類, 規格, 数量
      cols[0] = bw - 4 - cols[1] - cols[2] - cols[3];
      headRow = `<tr><th colspan="2">品名</th><th>規格</th><th>数量</th></tr>`;
      body = suppliesRows(b.items, rowH);
    } else {
      cols = uc.cols;
      headRow = `<tr><th>品名</th><th>単位</th><th>${current.layout === 'center' ? '発注' : '数量'}</th></tr>`;
      body = b.items.map(it => {
        const q = qtyText(it);
        return `<tr class="${q ? 'on' : ''}" style="height:${rowH}px">
          <td><span class="fit">${esc(it.name)}</span></td><td class="c"><span class="fit">${esc(unitOf(it))}</span></td><td class="q">${esc(q)}</td></tr>`;
      }).join('');
    }
    const colg = '<colgroup>' + cols.map(c => `<col style="width:${c}px">`).join('') + '</colgroup>';
    return `<table class="sh-table" style="width:${bw}px;font-size:${fs}px">${colg}<thead>${headRow}</thead><tbody>${body}</tbody></table>`;
  }).join('');
  return bare ? tables : `<div class="sh-blocks">${tables}</div>`;
}

function qtyText(it) {
  const p = parseQty(draft.qty[it.key]);
  return p.n > 0 ? fmtNum(p.n) : '';
}
// 消耗品の数量欄：単位を選んだときは「2kg」のように単位も付ける
function suppliesQty(it) {
  const q = qtyText(it);
  return q && unitOf(it) ? q + unitOf(it) : q;
}

// 消耗品：同じ品名が続く行は品名をまとめる（元の用紙と同じ形）
function suppliesRows(items, rowH) {
  const groups = [];
  items.forEach(it => {
    const last = groups[groups.length - 1];
    if (last && last.name === it.name && !it.noteLabel && !last.items[0].noteLabel) last.items.push(it);
    else groups.push({ name: it.name, items: [it] });
  });
  let html = '';
  groups.forEach(g => {
    const anyOn = g.items.some(it => qtyText(it));
    if (g.items.length === 1 && !g.items[0].kind) {
      const it = g.items[0], q = qtyText(it), h = rowH * (it.noteLabel ? 2 : 1);
      const note = it.noteLabel
        ? `<span class="fit sh-note">${esc(noteLabelShort(it.noteLabel))}：${esc((draft.note[it.key] || '').trim())}</span>` : '';
      html += `<tr class="${q ? 'on' : ''}" style="height:${h}px">
        <td colspan="2"><span class="fit">${esc(it.name)}</span>${note}</td>
        <td class="c"><span class="fit">${esc(it.unit)}</span></td><td class="q"><span class="fit">${esc(suppliesQty(it))}</span></td></tr>`;
      return;
    }
    g.items.forEach((it, i) => {
      const q = qtyText(it);
      const first = i === 0
        ? `<td class="grp${anyOn ? ' on' : ''}" rowspan="${g.items.length}"><span class="fitv" style="max-height:${rowH * g.items.length - 2}px">${esc(g.name)}</span></td>` : '';
      html += `<tr class="${q ? 'on' : ''}" style="height:${rowH}px">${first}
        <td class="c"><span class="fit">${esc(it.kind)}</span></td>
        <td class="c"><span class="fit">${esc(it.unit)}</span></td><td class="q"><span class="fit">${esc(suppliesQty(it))}</span></td></tr>`;
    });
  });
  return html;
}
const noteLabelShort = s => s.replace(/（.*$/, '');

// はみ出す文字を自動で小さくする
function fitText(root) {
  $$('.fit', root).forEach(sp => {
    sp.style.display = 'block'; sp.style.overflow = 'hidden'; sp.style.whiteSpace = 'nowrap';
    let size = parseFloat(getComputedStyle(sp).fontSize);
    while (sp.scrollWidth > sp.clientWidth + 0.5 && size > 6) { size -= 0.5; sp.style.fontSize = size + 'px'; }
  });
  $$('.fitv', root).forEach(sp => {
    sp.style.display = 'block'; sp.style.overflow = 'hidden';
    let size = parseFloat(getComputedStyle(sp).fontSize);
    while ((sp.scrollHeight > sp.clientHeight + 0.5 || sp.scrollWidth > sp.clientWidth + 0.5) && size > 6) { size -= 0.5; sp.style.fontSize = size + 'px'; }
  });
}

// ===== 確認画面 =====
let previewSheet = null;
function showPreview() {
  show('screen-preview');
  const { el, W, H } = buildSheet();
  previewSheet = { el, W, H };
  const scaler = $('#preview-scaler'), wrap = $('#preview-wrap');
  scaler.innerHTML = ''; scaler.appendChild(el);
  fitText(el);
  const avail = Math.min(wrap.parentElement.clientWidth, W);
  const s = avail / W;
  scaler.style.transform = `scale(${s})`;
  wrap.style.width = W * s + 'px';
  wrap.style.height = H * s + 'px';
}
window.addEventListener('resize', () => { if (!$('#screen-preview').hidden && previewSheet) {
  const { W, H } = previewSheet, wrap = $('#preview-wrap');
  const s = Math.min(wrap.parentElement.clientWidth, W) / W;
  $('#preview-scaler').style.transform = `scale(${s})`;
  wrap.style.width = W * s + 'px'; wrap.style.height = H * s + 'px';
} });

// 印刷
$('#btn-print').addEventListener('click', () => {
  if (!previewSheet) return;
  const area = $('#print-area');
  area.innerHTML = ''; area.appendChild(previewSheet.el.cloneNode(true));
  let ps = $('#page-style');
  if (!ps) { ps = document.createElement('style'); ps.id = 'page-style'; document.head.appendChild(ps); }
  ps.textContent = `@page { size: A4 ${current.orientation}; margin: 0; }`;
  window.print();
});
window.addEventListener('afterprint', () => { $('#print-area').innerHTML = ''; });

// OK → PDF作成
$('#btn-ok').addEventListener('click', async () => {
  if (!previewSheet) return;
  if (!window.html2canvas || !window.jspdf) { alert('PDF作成の部品を読み込めませんでした。インターネットにつながっているか確認して、ページを開き直してください。'); return; }
  $('#busy').hidden = false;
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:0;top:0;z-index:-10;pointer-events:none;';
  const clone = previewSheet.el.cloneNode(true);
  host.appendChild(clone); document.body.appendChild(host);
  try {
    const { W, H } = previewSheet;
    const canvas = await html2canvas(clone, { scale: PDF_SCALE, backgroundColor: '#ffffff', width: W, height: H, windowWidth: W + 40, windowHeight: H + 40, scrollX: 0, scrollY: 0, logging: false });
    const img = canvas.toDataURL('image/jpeg', 0.92);
    const land = current.orientation === 'landscape';
    const pdf = new window.jspdf.jsPDF({ orientation: land ? 'landscape' : 'portrait', unit: 'mm', format: 'a4', compress: true });
    pdf.addImage(img, 'JPEG', 0, 0, land ? 297 : 210, land ? 210 : 297);
    pdf.setProperties({ title: fileName().replace(/\.pdf$/, '') });
    const blob = pdf.output('blob');
    const name = fileName();
    pdfFile = new File([blob], name, { type: 'application/pdf' });
    if (pdfUrl) URL.revokeObjectURL(pdfUrl);
    pdfUrl = URL.createObjectURL(blob);
    openSendPanel(name);
  } catch (err) {
    alert('PDFの作成に失敗しました。\n' + (err && err.message ? err.message : err));
  } finally {
    host.remove();
    $('#busy').hidden = true;
  }
});

// ===== LINEへ送る =====
function canShareFile() {
  try { return !!(navigator.canShare && navigator.share && navigator.canShare({ files: [pdfFile] })); } catch { return false; }
}
function openSendPanel(name) {
  $('#send-filename').textContent = name;
  const share = canShareFile();
  $('#btn-share').hidden = !share;
  $('#share-help').hidden = !share;
  $('#save-help').hidden = share;
  const save = $('#btn-save');
  save.href = pdfUrl; save.download = name;
  save.className = share ? 'sub big' : 'primary big';
  $('#send-panel').hidden = false;
}
$('#btn-share').addEventListener('click', async () => {
  try {
    await navigator.share({ files: [pdfFile], title: pdfFile.name });
    afterSent();
  } catch (err) {
    if (err && err.name === 'AbortError') return;    // 共有をキャンセルした
    alert('共有できませんでした。「PDFを保存」から保存して、LINEで送ってください。');
    $('#save-help').hidden = false;
  }
});
$('#btn-save').addEventListener('click', () => {
  // 保存後はLINEアプリに移るので、戻ってきたときにクリアを確認する
  setTimeout(() => { addDoneButton(); }, 300);
});
function addDoneButton() {
  if ($('#btn-done')) return;
  const b = document.createElement('button');
  b.id = 'btn-done'; b.className = 'sub big'; b.textContent = '送信が終わった（入力をクリア）';
  b.onclick = () => afterSent(true);
  $('#btn-send-close').before(b);
}
function afterSent(noAsk) {
  if (noAsk || confirm('送信できましたか？\n「OK」で入力内容をクリアしてホームに戻ります。')) {
    resetCurrent();
    closeSendPanel();
    location.hash = '#/';
  }
}
function closeSendPanel() {
  $('#send-panel').hidden = true;
  const d = $('#btn-done'); if (d) d.remove();
}
$('#btn-send-close').addEventListener('click', closeSendPanel);

// ===== 起動 =====
(async function init() {
  await Promise.all(SHEETS.map(async s => {
    try { itemsBySheet[s.id] = await loadSheet(s); }
    catch (e) { loadErrors[s.id] = `・${s.label}：${e.message || e}`; }
  }));
  window.addEventListener('hashchange', route);
  route();
})();
