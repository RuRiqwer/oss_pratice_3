'use strict';
// HWPX(OWPML)와 DOCX(OOXML)는 모두 ZIP + XML이다.
// JSZip으로 압축을 풀고 브라우저 내장 DOMParser로 XML을 읽는다.

const $ = (id) => document.getElementById(id);
const PX = 1 / 75; // HWPUNIT(1/7200 inch) -> px(96dpi)
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml' };

// ---- XML 헬퍼: 접두사(hp:, hh:)에 의존하지 않도록 localName으로 찾는다 ----
const kids = (el, n) => Array.from(el.children).filter((c) => c.localName === n);
const first = (el, n) => (el ? kids(el, n)[0] : undefined);
const all = (el, n) => Array.from(el.getElementsByTagNameNS('*', n));
const num = (v) => Number(v) || 0;

function parseXml(text, label) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error(`${label} 파싱 실패`);
  return doc;
}
const findEntry = (zip, path) => {
  const p = path.replace(/^\.?\//, '').toLowerCase();
  return Object.values(zip.files).find((f) => !f.dir && f.name.toLowerCase() === p);
};
async function readXml(zip, path) {
  const f = findEntry(zip, path);
  if (!f) throw new Error(`${path} 파일이 없습니다`);
  return parseXml(await f.async('string'), path);
}
const baseName = (p) => p.split('/').pop().replace(/\.[^.]+$/, '');

// ---- 1단계: header.xml 의 글자/문단 스타일 ----
function parseHeader(doc) {
  const charPr = new Map(), paraPr = new Map();
  for (const el of all(doc, 'charPr')) {
    const ul = first(el, 'underline'), st = first(el, 'strikeout');
    charPr.set(el.getAttribute('id'), {
      size: num(el.getAttribute('height')) / 100, // 1000 = 10pt
      color: el.getAttribute('textColor'),
      bold: !!first(el, 'bold'),
      italic: !!first(el, 'italic'),
      underline: !!ul && ul.getAttribute('type') !== 'NONE',
      strike: !!st && st.getAttribute('shape') !== 'NONE',
    });
  }
  for (const el of all(doc, 'paraPr')) {
    const align = all(el, 'align')[0]?.getAttribute('horizontal');
    const ls = all(el, 'lineSpacing')[0];
    paraPr.set(el.getAttribute('id'), {
      align: { CENTER: 'center', RIGHT: 'right', JUSTIFY: 'justify', LEFT: 'left' }[align],
      lineHeight: ls && ls.getAttribute('type') === 'PERCENT' ? num(ls.getAttribute('value')) / 100 : 0,
    });
  }
  return { charPr, paraPr };
}

// ---- 2단계: content.hpf (섹션 순서, 이미지 ID -> 파일) ----
async function loadManifest(zip) {
  const f = findEntry(zip, 'Contents/content.hpf');
  const items = new Map(); let spine = [];
  if (f) {
    const doc = parseXml(await f.async('string'), 'content.hpf');
    for (const it of all(doc, 'item')) items.set(it.getAttribute('id'), it.getAttribute('href'));
    spine = all(doc, 'itemref').map((r) => items.get(r.getAttribute('idref'))).filter((h) => /section\d+\.xml$/i.test(h || ''));
  }
  if (!spine.length) {
    spine = Object.keys(zip.files).filter((n) => /^contents\/section\d+\.xml$/i.test(n)).sort();
  }
  return { items, spine };
}
async function loadImages(zip, manifest, resources) {
  const urls = new Map();
  for (const f of Object.values(zip.files)) {
    if (f.dir || !/^bindata\//i.test(f.name)) continue;
    const ext = f.name.split('.').pop().toLowerCase();
    const blob = new Blob([await f.async('uint8array')], { type: MIME[ext] || '' });
    const url = URL.createObjectURL(blob);
    resources.add(url);
    urls.set(baseName(f.name), url);
  }
  for (const [id, href] of manifest.items) {
    const u = href && urls.get(baseName(href));
    if (u) urls.set(id, u);
  }
  return urls;
}

// ---- 3단계: 렌더링 (textContent / append(text)만 사용 -> XSS 방지) ----
function renderText(t, span) {
  for (const n of t.childNodes) {
    if (n.nodeType === Node.TEXT_NODE) span.append(n.data);
    else if (n.localName === 'lineBreak') span.append(document.createElement('br'));
    else if (n.localName === 'tab') span.append('\u2003');
    else if (n.localName === 'nbSpace') span.append('\u00A0');
    else if (n.localName === 'fwSpace') span.append('\u3000');
  }
}
function applyChar(span, c) {
  if (!c) return;
  if (c.size) span.style.fontSize = c.size + 'pt';
  if (/^#[0-9a-f]{6}$/i.test(c.color || '')) span.style.color = c.color;
  if (c.bold) span.style.fontWeight = 'bold';
  if (c.italic) span.style.fontStyle = 'italic';
  const deco = [c.underline && 'underline', c.strike && 'line-through'].filter(Boolean).join(' ');
  if (deco) span.style.textDecoration = deco;
}
function missing(msg) {
  const s = document.createElement('span'); s.className = 'missing'; s.textContent = msg; return s;
}
function renderPic(pic, ctx) {
  const ref = all(pic, 'img')[0]?.getAttribute('binaryItemIDRef');
  const url = ctx.urls.get(ref);
  if (!url) return missing(`[이미지 없음: ${ref}]`);
  const w = num(first(pic, 'curSz')?.getAttribute('width')) || num(first(pic, 'sz')?.getAttribute('width'));
  const img = new Image();
  img.src = url; img.alt = '문서 이미지'; img.style.maxWidth = '100%';
  if (w) img.style.width = w * PX + 'px';
  img.onerror = () => img.replaceWith(missing('[표시할 수 없는 이미지 형식]'));
  return img;
}
function renderTable(tbl, ctx) {
  const t = document.createElement('table'); t.className = 'hwpx-tbl';
  const sz = first(tbl, 'sz');
  if (sz) t.style.width = num(sz.getAttribute('width')) * PX + 'px';
  for (const tr of kids(tbl, 'tr')) {
    const row = t.insertRow();
    for (const tc of kids(tr, 'tc')) {
      const td = row.insertCell();
      const span = first(tc, 'cellSpan');
      if (span) { td.colSpan = num(span.getAttribute('colSpan')) || 1; td.rowSpan = num(span.getAttribute('rowSpan')) || 1; }
      const cw = num(first(tc, 'cellSz')?.getAttribute('width'));
      if (cw) td.style.width = cw * PX + 'px';
      for (const p of kids(first(tc, 'subList') || tc, 'p')) td.append(renderParagraph(p, ctx));
    }
  }
  return t;
}
function renderParagraph(p, ctx) {
  const box = document.createElement('div'); box.className = 'hwpx-p';
  const pp = ctx.styles.paraPr.get(p.getAttribute('paraPrIDRef'));
  if (pp?.align) box.style.textAlign = pp.align;
  if (pp?.lineHeight) box.style.lineHeight = pp.lineHeight;
  for (const run of kids(p, 'run')) {
    const cp = ctx.styles.charPr.get(run.getAttribute('charPrIDRef'));
    for (const node of run.children) {
      if (node.localName === 't') {
        const span = document.createElement('span');
        applyChar(span, cp); renderText(node, span); box.append(span);
      } else if (node.localName === 'tbl') box.append(renderTable(node, ctx));
      else if (node.localName === 'pic') box.append(renderPic(node, ctx));
    }
  }
  if (!box.childNodes.length) box.append(document.createElement('br')); // 빈 문단도 줄 높이 유지
  return box;
}
function renderSection(doc, ctx) {
  const sheet = document.createElement('section'); sheet.className = 'sheet';
  const pg = all(doc, 'pagePr')[0], mg = first(pg, 'margin');
  if (pg) sheet.style.width = num(pg.getAttribute('width')) * PX + 'px';
  if (mg) sheet.style.padding = ['top', 'right', 'bottom', 'left'].map((k) => num(mg.getAttribute(k)) * PX + 'px').join(' ');
  for (const p of kids(doc.documentElement, 'p')) sheet.append(renderParagraph(p, ctx));
  return sheet;
}

// ---- 4단계: DOCX의 스타일 / 관계 파일 / 본문 ----
const TWIP = 96 / 1440; // DOCX의 1/20pt -> px
const EMU = 96 / 914400; // DrawingML 이미지 크기 -> px
const attr = (el, name) => Array.from(el?.attributes || []).find((a) => a.localName === name)?.value ?? '';
const on = (el) => !/^(0|false|off)$/i.test(attr(el, 'val'));

function docxCharProps(pr) {
  const c = {};
  if (!pr) return c;
  for (const [tag, prop] of [['b', 'bold'], ['i', 'italic'], ['strike', 'strike']]) {
    const el = first(pr, tag);
    if (el) c[prop] = on(el);
  }
  const size = first(pr, 'sz'), color = first(pr, 'color'), u = first(pr, 'u');
  const fonts = first(pr, 'rFonts');
  if (size) c.size = num(attr(size, 'val')) / 2;
  if (/^[0-9a-f]{6}$/i.test(attr(color, 'val'))) c.color = '#' + attr(color, 'val');
  if (u) c.underline = !/^(none|0|false|off)$/i.test(attr(u, 'val'));
  const font = attr(fonts, 'eastAsia') || attr(fonts, 'ascii') || attr(fonts, 'hAnsi');
  if (font) c.font = font;
  return c;
}
function docxParaProps(pr) {
  const p = {};
  if (!pr) return p;
  const jc = first(pr, 'jc'), spacing = first(pr, 'spacing'), ind = first(pr, 'ind');
  if (jc) p.align = { both: 'justify', distribute: 'justify', center: 'center', right: 'right', end: 'right' }[attr(jc, 'val')] || 'left';
  for (const side of ['before', 'after']) {
    if (attr(spacing, side) !== '') p[side] = num(attr(spacing, side)) * TWIP;
  }
  if (attr(spacing, 'line') !== '') {
    const line = num(attr(spacing, 'line'));
    p.lineHeight = /^(exact|atLeast)$/.test(attr(spacing, 'lineRule')) ? line * TWIP + 'px' : String(line / 240);
  }
  for (const side of ['left', 'right']) {
    if (attr(ind, side) !== '') p[side] = num(attr(ind, side)) * TWIP;
  }
  if (attr(ind, 'firstLine') !== '' || attr(ind, 'hanging') !== '') {
    p.indent = (num(attr(ind, 'firstLine')) - num(attr(ind, 'hanging'))) * TWIP;
  }
  return p;
}
async function loadDocxStyles(zip) {
  const entries = new Map();
  let defaults = { char: { size: 11 }, para: {} }, defaultId = '';
  const f = findEntry(zip, 'word/styles.xml');
  if (f) {
    const doc = parseXml(await f.async('string'), 'word/styles.xml');
    const dd = all(doc, 'docDefaults')[0];
    defaults = {
      char: { ...defaults.char, ...docxCharProps(first(first(dd, 'rPrDefault'), 'rPr')) },
      para: docxParaProps(first(first(dd, 'pPrDefault'), 'pPr')),
    };
    for (const el of all(doc, 'style')) {
      const id = attr(el, 'styleId');
      entries.set(id, { basedOn: attr(first(el, 'basedOn'), 'val'), char: docxCharProps(first(el, 'rPr')), para: docxParaProps(first(el, 'pPr')) });
      if (attr(el, 'type') === 'paragraph' && /^(1|true|on)$/i.test(attr(el, 'default'))) defaultId = id;
    }
  }
  function resolve(id, seen = new Set()) {
    const s = entries.get(id);
    if (!s || seen.has(id)) return { char: {}, para: {} };
    seen.add(id);
    const base = resolve(s.basedOn, seen);
    return { char: { ...base.char, ...s.char }, para: { ...base.para, ...s.para } };
  }
  return { defaults, defaultId, resolve };
}

// relationship의 상대 경로를 ZIP 내부 경로로 바꾼다. 외부 이미지는 내려받지 않는다.
function docxPartPath(target) {
  try { target = decodeURIComponent(target); } catch { return ''; }
  if (/^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith('//')) return '';
  const parts = target.startsWith('/') ? [] : ['word'];
  for (const part of target.replace(/\\/g, '/').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { if (!parts.length) return ''; parts.pop(); }
    else parts.push(part);
  }
  return parts.join('/');
}
async function loadDocxImages(zip, resources) {
  const urls = new Map(), loaded = new Map();
  const f = findEntry(zip, 'word/_rels/document.xml.rels');
  if (!f) return urls;
  const doc = parseXml(await f.async('string'), 'word/_rels/document.xml.rels');
  for (const rel of all(doc, 'Relationship')) {
    if (!/\/image$/.test(attr(rel, 'Type')) || /^external$/i.test(attr(rel, 'TargetMode'))) continue;
    const path = docxPartPath(attr(rel, 'Target')), entry = findEntry(zip, path);
    if (!entry) continue;
    let url = loaded.get(entry.name);
    if (!url) {
      const ext = entry.name.split('.').pop().toLowerCase();
      url = URL.createObjectURL(new Blob([await entry.async('uint8array')], { type: MIME[ext] || '' }));
      resources.add(url); loaded.set(entry.name, url);
    }
    urls.set(attr(rel, 'Id'), url);
  }
  return urls;
}
function renderDocxImage(node, ctx) {
  const fragment = document.createDocumentFragment();
  const refs = [...all(node, 'blip'), ...all(node, 'imagedata')];
  for (const ref of refs) {
    const id = attr(ref, 'embed') || attr(ref, 'id'), url = ctx.urls.get(id);
    if (!url) { fragment.append(missing(`[이미지 없음: ${id || '외부 이미지'}]`)); continue; }
    const img = new Image();
    img.src = url; img.alt = attr(all(node, 'docPr')[0], 'descr') || '문서 이미지';
    img.style.maxWidth = '100%';
    const extent = all(node, 'extent')[0];
    const width = num(attr(extent, 'cx')) * EMU;
    const vmlWidth = /(?:^|;)\s*width\s*:\s*([\d.]+)(pt|px)/i.exec(attr(ref.parentElement, 'style'));
    if (width) img.style.width = width + 'px';
    else if (vmlWidth) img.style.width = vmlWidth[1] + vmlWidth[2];
    img.onerror = () => img.replaceWith(missing('[표시할 수 없는 이미지 형식]'));
    fragment.append(img);
  }
  return fragment;
}
function renderDocxInline(node, parent, ctx, baseChar) {
  const name = node.localName;
  if (['pPr', 'rPr', 'del', 'instrText'].includes(name)) return;
  if (name === 'r') {
    const pr = first(node, 'rPr'), styleId = attr(first(pr, 'rStyle'), 'val');
    const c = { ...baseChar, ...ctx.styles.resolve(styleId).char, ...docxCharProps(pr) };
    const span = document.createElement('span');
    applyChar(span, c);
    // false 속성도 지정하여 상위 문단 스타일을 해제할 수 있게 한다.
    if ('bold' in c) span.style.fontWeight = c.bold ? 'bold' : 'normal';
    if ('italic' in c) span.style.fontStyle = c.italic ? 'italic' : 'normal';
    if (c.font) span.style.fontFamily = c.font;
    for (const child of node.children) renderDocxInline(child, span, ctx, c);
    parent.append(span);
  } else if (name === 't') parent.append(node.textContent);
  else if (name === 'tab') parent.append('\u2003');
  else if (name === 'br' || name === 'cr') parent.append(document.createElement('br'));
  else if (name === 'noBreakHyphen') parent.append('\u2011');
  else if (name === 'softHyphen') parent.append('\u00AD');
  else if (name === 'drawing' || name === 'pict') parent.append(renderDocxImage(node, ctx));
  else if (name === 'AlternateContent') {
    const choice = first(node, 'Choice') || first(node, 'Fallback');
    if (choice) for (const child of choice.children) renderDocxInline(child, parent, ctx, baseChar);
  } else for (const child of node.children) renderDocxInline(child, parent, ctx, baseChar);
}
function renderDocxParagraph(p, ctx) {
  const box = document.createElement('div'); box.className = 'hwpx-p';
  const pr = first(p, 'pPr'), styleId = attr(first(pr, 'pStyle'), 'val') || ctx.styles.defaultId;
  const style = ctx.styles.resolve(styleId);
  const para = { ...ctx.styles.defaults.para, ...style.para, ...docxParaProps(pr) };
  const baseChar = { ...ctx.styles.defaults.char, ...style.char };
  box.style.fontSize = baseChar.size + 'pt';
  if (para.align) box.style.textAlign = para.align;
  if (para.lineHeight) box.style.lineHeight = para.lineHeight;
  for (const [prop, css] of [['before', 'marginTop'], ['after', 'marginBottom'], ['left', 'marginLeft'], ['right', 'marginRight'], ['indent', 'textIndent']]) {
    if (para[prop] !== undefined) box.style[css] = para[prop] + 'px';
  }
  for (const child of p.children) renderDocxInline(child, box, ctx, baseChar);
  if (!box.textContent && !box.querySelector('img, br, .missing')) box.append(document.createElement('br'));
  return box;
}
function setDocxWidth(el, width) {
  const value = num(attr(width, 'w')), type = attr(width, 'type');
  if (value > 0 && type === 'dxa') el.style.width = value * TWIP + 'px';
  else if (value > 0 && type === 'pct') el.style.width = value / 50 + '%';
}
function renderDocxTable(tbl, ctx) {
  const table = document.createElement('table'); table.className = 'hwpx-tbl';
  setDocxWidth(table, first(first(tbl, 'tblPr'), 'tblW'));
  let previous = new Map();
  for (const tr of kids(tbl, 'tr')) {
    const row = table.insertRow(), next = new Map();
    let column = num(attr(first(first(tr, 'trPr'), 'gridBefore'), 'val'));
    for (const tc of kids(tr, 'tc')) {
      const pr = first(tc, 'tcPr'), merge = first(pr, 'vMerge');
      const colSpan = Math.max(1, num(attr(first(pr, 'gridSpan'), 'val')) || 1);
      const above = previous.get(column);
      if (merge && attr(merge, 'val') !== 'restart' && above && above.colSpan === colSpan) {
        above.rowSpan += 1; next.set(column, above); column += colSpan; continue;
      }
      const cell = row.insertCell(); cell.colSpan = colSpan;
      setDocxWidth(cell, first(pr, 'tcW'));
      renderDocxBlocks(tc, cell, ctx);
      if (merge) next.set(column, cell);
      column += colSpan;
    }
    previous = next;
  }
  return table;
}
function renderDocxBlocks(source, target, ctx) {
  for (const node of source.children) {
    if (node.localName === 'p') target.append(renderDocxParagraph(node, ctx));
    else if (node.localName === 'tbl') target.append(renderDocxTable(node, ctx));
    else if (['sdt', 'sdtContent', 'customXml'].includes(node.localName)) renderDocxBlocks(node, target, ctx);
  }
}
async function buildDocx(zip, resources) {
  const doc = await readXml(zip, 'word/document.xml'), body = first(doc.documentElement, 'body');
  if (!body) throw new Error('DOCX 본문(w:body)이 없습니다.');
  const styles = await loadDocxStyles(zip), urls = await loadDocxImages(zip, resources);
  const sheet = document.createElement('section'); sheet.className = 'sheet';
  sheet.style.width = '816px'; sheet.style.padding = '64px';
  const section = first(body, 'sectPr') || all(body, 'sectPr').slice(-1)[0];
  const size = first(section, 'pgSz'), margin = first(section, 'pgMar');
  if (num(attr(size, 'w'))) sheet.style.width = num(attr(size, 'w')) * TWIP + 'px';
  if (margin) sheet.style.padding = ['top', 'right', 'bottom', 'left'].map((side) => num(attr(margin, side)) * TWIP + 'px').join(' ');
  renderDocxBlocks(body, sheet, { styles, urls });
  return { sheets: [sheet], summary: `문단 ${all(body, 'p').length}개, 이미지 ${new Set(urls.values()).size}개` };
}

// ---- 5단계: 형식 자동 판별 / 공통 파일 열기 / 기존 UI ----
function detectFormat(zip) {
  // 확장자만 보지 않고 ZIP의 실제 내부 구조로 문서 형식을 판별한다.
  if (findEntry(zip, 'word/document.xml')) return 'DOCX';
  if (findEntry(zip, 'Contents/header.xml')) return 'HWPX';
  throw new Error('지원하지 않는 문서 형식입니다. HWPX 또는 DOCX 파일을 선택해 주세요.');
}
async function buildHwpx(zip, resources) {
  const styles = parseHeader(await readXml(zip, 'Contents/header.xml'));
  const manifest = await loadManifest(zip);
  if (!manifest.spine.length) throw new Error('HWPX 본문 섹션이 없습니다.');
  const urls = await loadImages(zip, manifest, resources), sheets = [];
  for (const path of manifest.spine) sheets.push(renderSection(await readXml(zip, path), { styles, urls }));
  return { sheets, summary: `섹션 ${manifest.spine.length}개, 이미지 ${new Set(urls.values()).size}개` };
}
let currentUrls = [], openSequence = 0;
async function openDocument(file) {
  const sequence = ++openSequence, resources = new Set();
  const status = $('status'), viewer = $('viewer');
  status.className = ''; status.textContent = `${file.name} 여는 중...`;
  currentUrls.forEach((url) => URL.revokeObjectURL(url)); currentUrls = [];
  viewer.replaceChildren(); $('tree').textContent = ''; $('structure').hidden = true;
  try {
    if (typeof JSZip === 'undefined') throw new Error('JSZip을 불러오지 못했습니다. 인터넷 연결을 확인하고 새로고침해 주세요.');
    let zip;
    try { zip = await JSZip.loadAsync(await file.arrayBuffer()); }
    catch {
      if (/\.hwp$/i.test(file.name)) throw new Error('.hwp(바이너리) 형식은 지원하지 않습니다. .hwpx로 저장해 주세요.');
      if (/\.doc$/i.test(file.name)) throw new Error('.doc(바이너리) 형식은 지원하지 않습니다. .docx로 저장해 주세요.');
      throw new Error('정상적인 HWPX/DOCX 파일이 아닙니다. 파일이 손상되었거나 암호화되었을 수 있습니다.');
    }
    const format = detectFormat(zip);
    const result = format === 'DOCX' ? await buildDocx(zip, resources) : await buildHwpx(zip, resources);
    if (sequence !== openSequence) return; // 연속 업로드 시 마지막으로 선택한 문서만 표시
    viewer.replaceChildren(...result.sheets);
    currentUrls = [...resources]; resources.clear();
    $('tree').textContent = Object.values(zip.files).map((f) => (f.dir ? f.name : `${f.name}  (${f._data?.uncompressedSize ?? '?'} B)`)).join('\n');
    $('structure').hidden = false;
    status.textContent = `${file.name} - ${format}, ${result.summary}`;
  } catch (e) {
    if (sequence !== openSequence) return;
    status.className = 'error'; status.textContent = `열 수 없습니다: ${e.message}`;
  } finally {
    resources.forEach((url) => URL.revokeObjectURL(url));
  }
}

$('file').addEventListener('change', (e) => {
  const file = e.target.files[0];
  e.target.value = ''; // 같은 파일도 다시 선택 가능
  if (file) openDocument(file);
});
const drop = $('drop');
['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => e.dataTransfer.files[0] && openDocument(e.dataTransfer.files[0]));
