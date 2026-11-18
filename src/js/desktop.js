// SPDX-License-Identifier: MIT OR Apache-2.0
// Built by build.sh from lib/. Do not edit by hand.
(function () {
'use strict';
// Fill a Word template with data, working on the Oxi document model (the JSON
// that docx_to_model returns and docx_from_model writes back).
//
// Placeholders are written in the template as {{name}}. A placeholder may be
// split across runs (Word does this when part of it is formatted differently
// or was typed in two goes); the replacement keeps the formatting of the run
// where the placeholder starts.
//
// A table row that mentions {{list.field}}, where `list` is an array in the
// data, is repeated once per entry of that array; inside the copies,
// {{list.field}} reads the entry's field. Blank rows right under it (the
// empty lines of a printed form) are used up first, so the table keeps its
// height until there are more entries than lines; with no entries the row
// stays, emptied.
//
// {{name[3]}} is the third character of the value and {{name[-1]}} the last
// one, for forms that print one character per box (amounts, registration
// numbers); {{name:和暦}} writes a date as 令和8年11月17日. Placeholders inside text boxes are filled too; headers and
// footers are not part of the model and are left as they are.

const PLACEHOLDER = /\{\{\s*([^{}]+?)\s*\}\}/g;

/** Every {"Text": ...} item held directly by the runs of a paragraph. */
function textItems(paragraph) {
  const items = [];
  for (const inline of paragraph.content || []) {
    const run = inline.Run;
    if (!run) continue;
    for (const item of run.content || []) {
      if (typeof item.Text === 'string') items.push(item);
    }
  }
  return items;
}

/**
 * The w:t elements of a paragraph kept as raw XML (text boxes and other
 * drawings), wrapped so they read and write like {"Text": ...} items.
 */
function rawTextItems(p) {
  const items = [];
  for (const r of p.children || []) {
    if (!r || r.name !== 'w:r') continue;
    for (const t of r.children || []) {
      if (!t || t.name !== 't' && t.name !== 'w:t') continue;
      items.push({
        get Text() { return (t.children || []).filter((c) => typeof c === 'string').join(''); },
        set Text(v) {
          t.children = v === '' ? [] : [v];
          const attrs = t.attrs || (t.attrs = []);
          if (!attrs.some((a) => a[0] === 'xml:space')) attrs.push(['xml:space', 'preserve']);
        },
      });
    }
  }
  return items;
}

/** Replace the placeholders spread over `items`, in place. */
function fillItems(items, lookup) {
  if (items.length === 0) return;
  const starts = [];
  let full = '';
  for (const item of items) { starts.push(full.length); full += item.Text; }
  if (!full.includes('{{')) return;
  const matches = [...full.matchAll(PLACEHOLDER)];
  // Right to left, so the offsets of earlier placeholders stay valid.
  for (const m of matches.reverse()) {
    const value = lookup(m[1]);
    if (value === undefined) continue; // unknown name: leave it visible
    const start = m.index, end = m.index + m[0].length;
    const segAt = (pos) => {
      let i = 0;
      while (i + 1 < items.length && starts[i + 1] <= pos) i++;
      return i;
    };
    const s0 = segAt(start), s1 = segAt(end - 1);
    const a = items[s0].Text, from = start - starts[s0];
    if (s0 === s1) {
      items[s0].Text = a.slice(0, from) + value + a.slice(end - starts[s0]);
    } else {
      items[s0].Text = a.slice(0, from) + value;
      for (let i = s0 + 1; i < s1; i++) items[i].Text = '';
      items[s1].Text = items[s1].Text.slice(end - starts[s1]);
    }
  }
}

/** All text of a subtree, to see which placeholders a row mentions. */
function textOf(node) {
  if (node == null) return '';
  if (typeof node === 'string') return '';
  if (Array.isArray(node)) return node.map(textOf).join('');
  if (typeof node.Text === 'string') return node.Text;
  if (node.name === 'w:t' || node.name === 't') {
    return (node.children || []).filter((c) => typeof c === 'string').join('');
  }
  return Object.values(node).map(textOf).join('');
}

/** The list a row repeats over: the first {{list.field}} naming an array. */
function rowList(row, data) {
  for (const m of textOf(row).matchAll(PLACEHOLDER)) {
    const dot = m[1].indexOf('.');
    if (dot > 0 && Array.isArray(data[m[1].slice(0, dot)])) return m[1].slice(0, dot);
  }
  return null;
}

/** A row with no text at all: a blank line of a printed form. */
function blankRow(item) {
  return !!item.Row && textOf(item.Row).trim() === '';
}

/** A copy of a repeating row with every placeholder emptied. */
function clearedRow(item) {
  const copy = structuredClone(item);
  walk(copy, () => '', {});
  return copy;
}

function walk(node, lookup, data) {
  if (node == null || typeof node !== 'object') return;
  if (Array.isArray(node)) { node.forEach((n) => walk(n, lookup, data)); return; }
  if (node.Table) {
    const src = node.Table.rows;
    const rows = [];
    for (let i = 0; i < src.length; i++) {
      const item = src[i];
      const list = item.Row ? rowList(item.Row, data) : null;
      if (!list) { rows.push(item); continue; }
      // A printed form leaves blank rows under the first line of a list. The
      // entries go into those rows first, so the table keeps its height; only
      // more entries than lines make it grow.
      let blanks = 0;
      while (i + 1 + blanks < src.length && blankRow(src[i + 1 + blanks])) blanks++;
      const entries = data[list];
      for (const entry of entries) {
        const copy = structuredClone(item);
        const inner = (name) => {
          if (name.startsWith(list + '.')) {
            const v = entryValue(entry, name.slice(list.length + 1));
            return v === undefined ? undefined : v;
          }
          return lookup(name);
        };
        walk(copy, inner, data);
        rows.push(copy);
      }
      const free = blanks + 1 - entries.length;
      if (free > 0) {
        if (entries.length === 0) rows.push(clearedRow(item));
        for (let k = blanks - free + (entries.length === 0 ? 1 : 0); k < blanks; k++) rows.push(src[i + 1 + k]);
      }
      i += blanks;
    }
    node.Table.rows = rows;
    for (const item of rows) walk(item, lookup, data);
    return;
  }
  if (node.Paragraph) fillItems(textItems(node.Paragraph), lookup);
  else if (node.name === 'w:p' && Array.isArray(node.children)) fillItems(rawTextItems(node), lookup);
  for (const value of Object.values(node)) walk(value, lookup, data);
}

/**
 * A name may carry a format and an index: {{name:和暦}}, {{name[-1]}}.
 *
 * `:和暦` writes a date (2026-11-17) as 令和8年11月17日 and `:年月日` as
 * 2026年11月17日; values that are not dates are left as they are.
 *
 * `[i]` is the i-th character of the value (from 1) and `[-i]` the i-th from
 * the end: one character per box, as printed forms ask for amounts and
 * registration numbers. Separators (",") are left out of the count.
 */
const NAME = /^(.+?)(?::(和暦|年月日))?(?:\[(-?\d+)\])?$/;
const DATE = /^(\d{4})-(\d{2})-(\d{2})/;
function formatValue(value, format) {
  const d = String(value).match(DATE);
  if (!format || !d) return String(value);
  const [y, m, day] = [Number(d[1]), Number(d[2]), Number(d[3])];
  if (format === '年月日') return `${y}年${m}月${day}日`;
  const ymd = y * 10000 + m * 100 + day;
  const [era, first] = ymd >= 20190501 ? ['令和', 2019] : ymd >= 19890108 ? ['平成', 1989] : ['昭和', 1926];
  const n = y - first + 1;
  return `${era}${n === 1 ? '元' : n}年${m}月${day}日`;
}
function pick(value, index) {
  const chars = [...value.replace(/,/g, '')];
  const at = index > 0 ? index - 1 : chars.length + index;
  return index !== 0 && at >= 0 && at < chars.length ? chars[at] : '';
}
/** The base field name of a placeholder, without format or index. */
function baseName(name) {
  return name.match(NAME)[1];
}
function entryValue(entry, name) {
  const [, field, format, index] = name.match(NAME);
  const v = entry[field];
  if (v === undefined || Array.isArray(v)) return undefined;
  const text = formatValue(v, format);
  return index === undefined ? text : pick(text, Number(index));
}

/**
 * Fill `model` (the parsed JSON of docx_to_model) with `data`, in place.
 * `data` maps names to strings or numbers, and list names to arrays of such
 * maps. Returns the names used in the template that `data` does not have.
 */
function fillModel(model, data) {
  const missing = new Set();
  const lookup = (name) => {
    const v = entryValue(data, name);
    if (v === undefined) missing.add(baseName(name));
    return v;
  };
  walk(model, lookup, data);
  return [...missing];
}

/** Fill a .docx template; `api` is the initialised Oxi wasm module. */
function fillDocx(api, templateBytes, data) {
  const model = JSON.parse(api.docx_to_model(templateBytes));
  const missing = fillModel(model, data);
  const docx = api.docx_from_model(JSON.stringify(model), templateBytes);
  return { docx, missing };
}

/**
 * Turn a kintone record (event.record or REST record) into fill data:
 * numbers are formatted with thousands separators, subtables become arrays.
 */
function kintoneRecordData(record) {
  const fmt = (field) => {
    if (field == null) return '';
    const { type, value } = field;
    if (value == null || value === '') return '';
    if (type === 'NUMBER' || type === 'CALC') {
      const n = Number(value);
      return Number.isFinite(n) ? n.toLocaleString('ja-JP') : String(value);
    }
    if (Array.isArray(value)) {
      return value.map((v) => (typeof v === 'object' ? v.name ?? v.code ?? '' : v)).join('、');
    }
    if (typeof value === 'object') return value.name ?? value.code ?? '';
    return String(value);
  };
  const data = {};
  for (const [code, field] of Object.entries(record)) {
    if (field.type === 'SUBTABLE') {
      data[code] = field.value.map((row) => {
        const entry = {};
        for (const [c, f] of Object.entries(row.value)) entry[c] = fmt(f);
        return entry;
      });
    } else if (field.type !== 'FILE') {
      data[code] = fmt(field);
    }
  }
  return data;
}

/**
 * Like kintoneRecordData, but numbers stay numbers, so a spreadsheet cell
 * that is just {{金額}} gets a number and the template's number format.
 */
function kintoneRecordValues(record) {
  const text = kintoneRecordData(record);
  const typed = (field, value) => {
    if (field && (field.type === 'NUMBER' || field.type === 'CALC') && field.value !== '' && field.value != null) {
      const n = Number(field.value);
      if (Number.isFinite(n)) return n;
    }
    return value;
  };
  const data = {};
  for (const [code, field] of Object.entries(record)) {
    if (!(code in text)) continue;
    if (field.type === 'SUBTABLE') {
      data[code] = field.value.map((row, i) => {
        const entry = {};
        for (const [c, f] of Object.entries(row.value)) entry[c] = typed(f, text[code][i][c]);
        return entry;
      });
    } else {
      data[code] = typed(field, text[code]);
    }
  }
  return data;
}
// Fill an Excel template with data, working on the workbook the Oxi engine
// parses (parse_spreadsheet), then recalculate and write it back.
//
// The placeholders are the Word template's: {{name}}, {{list.field}} in a
// row of a list, {{name[-1]}} for one character per cell, {{name:和暦}}.
// A cell that is nothing but one placeholder takes the value as it is, so a
// number stays a number and the cell's number format (#,##0, ¥…) applies;
// placeholders inside longer text are replaced as text.
//
// A row naming {{list.field}} is the first line of the list: entry k goes
// into the row k below it, in the same columns. Rows are not inserted, since
// that would move every formula, merge and print area under them; the
// template provides as many lines as the form prints. Entries beyond the
// last free line are reported, not written over what is below.

const xlText = (cell) => (cell && cell.value && typeof cell.value === 'object' && typeof cell.value.String === 'string' ? cell.value.String : null);
const xlEmpty = (cell) => !cell || cell.value === 'Empty' || (xlText(cell) === '' && !cell.formula);

/** The value for `name` from `scope`: a number or a string, undefined when absent. */
function xlResolve(scope, name) {
  const [, field, format, index] = name.match(NAME);
  const v = scope[field];
  if (v === undefined || Array.isArray(v)) return undefined;
  if (format) return formatValue(v, format);
  if (index !== undefined) return pick(typeof v === 'number' ? String(v) : String(v), Number(index));
  return v;
}

const xlAsText = (v) => (typeof v === 'number' ? v.toLocaleString('ja-JP') : String(v));

/** Set a cell from its template text; returns the names not found. */
function xlFillCell(cell, template, lookup, missing) {
  const whole = template.match(/^\s*\{\{\s*([^{}]+?)\s*\}\}\s*$/);
  if (whole) {
    const v = lookup(whole[1]);
    if (v === undefined) { missing.add(baseName(whole[1])); cell.value = { String: template }; return; }
    if (typeof v === 'number') cell.value = { Number: v };
    else cell.value = v === '' ? 'Empty' : { String: v };
    return;
  }
  cell.value = {
    String: template.replace(PLACEHOLDER, (m, name) => {
      const v = lookup(name);
      if (v === undefined) { missing.add(baseName(name)); return m; }
      return xlAsText(v);
    }),
  };
}

function xlRowAt(sheet, index) {
  let row = sheet.rows.find((r) => r.index === index);
  if (!row) {
    row = { cells: [], height: null, index };
    sheet.rows.push(row);
    sheet.rows.sort((a, b) => a.index - b.index);
  }
  return row;
}

function xlCellAt(row, col, style) {
  let cell = row.cells.find((c) => c.col === col);
  if (!cell) {
    cell = { col, style: structuredClone(style), value: 'Empty' };
    row.cells.push(cell);
    row.cells.sort((a, b) => a.col - b.col);
  }
  return cell;
}

/**
 * Fill `workbook` (the object parse_spreadsheet returns) with `data`, in
 * place. Returns { missing, overflow }: names the template uses that `data`
 * lacks, and for each list that did not fit, how many entries were left out.
 */
function fillWorkbook(workbook, data) {
  const missing = new Set();
  const overflow = {};
  const lookup = (name) => xlResolve(data, name);
  for (const sheet of workbook.sheets) {
    const listRows = [];
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        const t = xlText(cell);
        if (t === null) continue;
        for (const m of t.matchAll(PLACEHOLDER)) {
          const dot = m[1].indexOf('.');
          const list = dot > 0 ? m[1].slice(0, dot) : null;
          if (list && Array.isArray(data[list]) && !listRows.some((l) => l.row === row)) listRows.push({ row, list });
        }
      }
    }
    for (const { row, list } of listRows) {
      const templates = row.cells.filter((c) => xlText(c) !== null && xlText(c).includes('{{'))
        .map((c) => ({ col: c.col, style: c.style, text: xlText(c) }));
      const cols = templates.map((t) => t.col);
      const entries = data[list];
      let written = 0;
      for (let k = 0; k < Math.max(entries.length, 1); k++) {
        const target = k === 0 ? row : xlRowAt(sheet, row.index + k);
        if (k > 0 && !cols.every((col) => xlEmpty(target.cells.find((c) => c.col === col)))) break;
        const entry = entries[k];
        for (const t of templates) {
          const cell = xlCellAt(target, t.col, t.style);
          if (!entry) { cell.value = 'Empty'; continue; }
          xlFillCell(cell, t.text, (name) => (name.startsWith(list + '.') ? xlResolve(entry, name.slice(list.length + 1)) : lookup(name)), missing);
        }
        if (entry) written++;
      }
      if (written < entries.length) overflow[list] = entries.length - written;
    }
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        const t = xlText(cell);
        if (t !== null && t.includes('{{')) xlFillCell(cell, t, lookup, missing);
      }
    }
  }
  return { missing: [...missing], overflow };
}

/**
 * Fill an .xlsx template and return the filled workbook's bytes, with every
 * formula recalculated so totals follow the new values. `api` is the
 * initialised Oxi wasm module (suite build).
 */
function fillXlsx(api, templateBytes, data) {
  const workbook = api.parse_spreadsheet(templateBytes);
  const { missing, overflow } = fillWorkbook(workbook, data);
  const recalculated = api.recalculate_spreadsheet(workbook, Date.now());
  const xlsx = api.edit_xlsx_from_workbook(templateBytes, recalculated);
  return { xlsx, missing, overflow };
}
// Plugin settings layout (kintone keeps settings as string values, at most
// 65,535 characters each and 256 KB in total):
//   templates: JSON [{ name, kind, parts, size }], kind 'docx' or 'xlsx'
//   tpl_<i>_<j>: base64 chunk j of template i
//   output_field: attachment field code that receives generated PDFs ('' = none)
var CHUNK = 60000;

function bytesToBase64(bytes) {
  var s = '';
  for (var i = 0; i < bytes.length; i += 0x8000) {
    s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(s);
}

function base64ToBytes(b64) {
  var s = atob(b64), out = new Uint8Array(s.length);
  for (var i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}

/** Templates with their bytes, from a getConfig() result. */
function readTemplates(config) {
  var list = [];
  try { list = JSON.parse(config.templates || '[]'); } catch (e) { list = []; }
  return list.map(function (t, i) {
    var b64 = '';
    for (var j = 0; j < t.parts; j++) b64 += config['tpl_' + i + '_' + j] || '';
    return { name: t.name, kind: t.kind || 'docx', bytes: base64ToBytes(b64) };
  });
}

/** A setConfig() object for these templates and output field. */
function writeTemplates(templates, outputField) {
  var config = { output_field: outputField || '' }, meta = [];
  templates.forEach(function (t, i) {
    var b64 = bytesToBase64(t.bytes), parts = 0;
    for (var o = 0; o < b64.length; o += CHUNK) config['tpl_' + i + '_' + parts++] = b64.slice(o, o + CHUNK);
    meta.push({ name: t.name, kind: t.kind || 'docx', parts: parts, size: t.bytes.length });
  });
  config.templates = JSON.stringify(meta);
  return config;
}

/** Total size kintone will count for a settings object, in bytes. */
function configBytes(config) {
  var n = 0;
  Object.keys(config).forEach(function (k) { n += new TextEncoder().encode(config[k]).length; });
  return n;
}
// Oxi for kintone (prototype). On the record detail screen it offers:
//   - 帳票: <name> for each Word or Excel template registered in the plugin
//     settings: fill its {{placeholders}} with this record (field codes or
//     the labels shown on the form; a row with {{subtable.field}} takes the
//     subtable's rows), show the PDF, and optionally attach it to the record;
//   - 表示 / 帳票 for each attached .docx or .xlsx, as is or filled.
// The Oxi engine (WebAssembly) is loaded from the public npm package via
// unpkg (jsDelivr refuses files over 50 MB); documents and record data never
// leave the browser.
var ENGINE = 'https://unpkg.com/@oxidocs/wasm@0.10.0/oxidocs_wasm.js';
// The metrics of uncommon faces (HG fonts, UD Digi Kyokasho…) come from the
// font plugin, fetched per face by range request (jsDelivr answers those).
var FONTS = 'https://cdn.jsdelivr.net/npm/@oxidocs/wasm-fonts@0.10.0/index.js';
if (typeof window !== 'undefined' && window.OXI_ENGINE_URL) ENGINE = window.OXI_ENGINE_URL;
if (typeof window !== 'undefined' && window.OXI_FONTS_URL) FONTS = window.OXI_FONTS_URL;
var withFontsOf = null; // the engine's withFonts, once loaded
var enginePromise = null;
var PLUGIN_ID = kintone.$PLUGIN_ID;

function loadEngine(report) {
  if (!enginePromise) {
    var t0 = performance.now();
    report('Oxi エンジンを読み込み中…（初回は約 24MB）');
    enginePromise = import(ENGINE).then(function (mod) {
      return mod.default().then(function () {
        console.log('[oxi] engine ready in ' + Math.round(performance.now() - t0) + ' ms');
        // Font plugins: without them a document in an uncommon face still
        // converts, with fallback metrics, so a failure here is only logged.
        return Promise.all([import(ENGINE.replace(/oxidocs_wasm\.js$/, 'fonts.js')), import(FONTS)])
          .then(function (m) { m[0].useFonts(mod, m[1].default); withFontsOf = m[0].withFonts; })
          .catch(function (e) { console.warn('[oxi] font plugin not loaded', e); })
          .then(function () { return mod; });
      });
    });
    enginePromise.catch(function () { enginePromise = null; });
  }
  return enginePromise;
}

function download(fileKey) {
  var url = kintone.api.url('/k/v1/file', true) + '?fileKey=' + encodeURIComponent(fileKey);
  return fetch(url, { headers: { 'X-Requested-With': 'XMLHttpRequest' }, credentials: 'same-origin' })
    .then(function (res) {
      if (!res.ok) throw new Error('ファイルを取得できません（HTTP ' + res.status + '）');
      return res.arrayBuffer();
    });
}

/** Field labels by code, subtable members included, so {{表示名}} works too. */
function formLabels() {
  var appId = kintone.app.getId();
  if (!appId) return Promise.resolve({});
  return kintone.api(kintone.api.url('/k/v1/app/form/fields', true), 'GET', { app: appId })
    .then(function (resp) {
      var labels = {};
      Object.keys(resp.properties).forEach(function (code) {
        var p = resp.properties[code];
        labels[code] = p.label;
        if (p.type === 'SUBTABLE') {
          Object.keys(p.fields).forEach(function (c) { labels[c] = p.fields[c].label; });
        }
      });
      return labels;
    })
    .catch(function () { return {}; });
}

/** Add each value under its label as well as its code. */
function withLabels(data, labels) {
  var out = {};
  Object.keys(data).forEach(function (code) {
    var v = data[code];
    if (Array.isArray(v)) {
      v = v.map(function (entry) {
        var e = {};
        Object.keys(entry).forEach(function (c) {
          e[c] = entry[c];
          if (labels[c] && !(labels[c] in e)) e[labels[c]] = entry[c];
        });
        return e;
      });
    }
    out[code] = v;
    if (labels[code] && !(labels[code] in out)) out[labels[code]] = v;
  });
  return out;
}

// kintone's stylesheet greys out bare buttons; spell the look out.
var BUTTON = 'background:#fff;color:#333;border:1px solid #ccc;border-radius:4px;padding:4px 12px;font-size:14px;cursor:pointer';

function overlay() {
  var box = document.getElementById('oxi-preview-overlay');
  if (box) return box;
  box = document.createElement('div');
  box.id = 'oxi-preview-overlay';
  box.style.cssText = 'position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.75);display:flex;flex-direction:column;padding:24px;box-sizing:border-box';
  var bar = document.createElement('div');
  // Opaque, so kintone's own header does not show through behind the text.
  bar.style.cssText = 'display:flex;gap:12px;align-items:center;color:#fff;background:#2b2b2b;padding:8px 12px;border-radius:6px 6px 0 0;font-size:14px';
  var status = document.createElement('span');
  status.className = 'oxi-status';
  status.style.flex = '1';
  var save = document.createElement('a');
  save.className = 'oxi-save';
  save.textContent = 'PDF を保存';
  save.style.cssText = 'color:#fff;text-decoration:underline;display:none';
  var attach = document.createElement('button');
  attach.type = 'button';
  attach.className = 'oxi-attach';
  attach.textContent = 'レコードに添付';
  attach.style.cssText = BUTTON + ';display:none';
  var close = document.createElement('button');
  close.type = 'button';
  close.textContent = '閉じる';
  close.style.cssText = BUTTON;
  close.onclick = function () { box.remove(); };
  bar.appendChild(status);
  bar.appendChild(save);
  bar.appendChild(attach);
  bar.appendChild(close);
  var frame = document.createElement('iframe');
  frame.className = 'oxi-frame';
  frame.style.cssText = 'flex:1;border:0;background:#fff';
  box.appendChild(bar);
  box.appendChild(frame);
  document.body.appendChild(box);
  return box;
}

/**
 * Show a PDF made from `source` ({ name, load() -> Promise<bytes> }); with
 * `fill`, the source is a template filled from `record` first. `attachTo` is
 * the attachment field code to offer saving into ('' for none).
 */
function show(source, record, fill, attachTo) {
  var box = overlay();
  var status = box.querySelector('.oxi-status');
  var frame = box.querySelector('.oxi-frame');
  var save = box.querySelector('.oxi-save');
  var attach = box.querySelector('.oxi-attach');
  save.style.display = 'none';
  attach.style.display = 'none';
  function report(text) { status.textContent = source.name + ' — ' + text; }
  report('準備中…');
  Promise.all([loadEngine(report), source.load(), fill ? formLabels() : Promise.resolve({})])
    .then(function (r) {
      var api = r[0], bytes = new Uint8Array(r[1]), note = '';
      var t0 = performance.now();
      var excel = source.kind === 'xlsx';
      if (excel && !api.xlsx_to_pdf) throw new Error('Excel のひな形には、新しい変換エンジンが必要です。プラグインを更新してください。');
      if (fill) {
        report('帳票を作成中…');
        var filled;
        if (excel) {
          filled = fillXlsx(api, bytes, withLabels(kintoneRecordValues(record), r[2]));
          bytes = filled.xlsx;
          var over = Object.keys(filled.overflow).map(function (k) { return k + ' ' + filled.overflow[k] + ' 件'; });
          if (over.length) note += '　行が足りず入らなかった明細: ' + over.join('、');
        } else {
          filled = fillDocx(api, bytes, withLabels(kintoneRecordData(record), r[2]));
          bytes = filled.docx;
        }
        if (filled.missing.length) note += '　見つからない項目: ' + filled.missing.join('、');
      }
      report('PDF に変換中…');
      var title = source.name.replace(/\.(docx|xlsx)$/i, '');
      // Engines from 0.9.2 on write a PDF title, which the viewer shows in
      // place of the blob URL; older ones only have docx_to_pdf.
      var convert = function () {
        return excel ? api.xlsx_to_pdf(bytes, title)
          : api.docx_to_pdf_with_title ? api.docx_to_pdf_with_title(bytes, title) : api.docx_to_pdf(bytes);
      };
      return Promise.resolve(withFontsOf ? withFontsOf(api, convert) : convert()).then(function (pdf) { return { pdf: pdf, t0: t0, note: note, title: title }; });
    })
    .then(function (done) {
      var pdf = done.pdf, note = done.note, title = done.title;
      var ms = Math.round(performance.now() - done.t0);
      var fileName = title + (fill ? '_' + (record.$id ? record.$id.value : '') : '') + '.pdf';
      var url = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }));
      frame.src = url;
      save.href = url;
      save.download = fileName;
      save.style.display = '';
      if (attachTo) {
        attach.style.display = '';
        attach.disabled = false;
        attach.onclick = function () {
          attach.disabled = true;
          report('レコードに添付中…');
          attachPdf(pdf, fileName, record, attachTo)
            .then(function () { report('レコードに添付しました。画面を更新します…'); window.location.reload(); })
            .catch(function (e) {
              console.error('[oxi]', e);
              attach.disabled = false;
              report(attachError(e));
            });
        };
      }
      report('ブラウザ内で作成しました（' + ms + ' ms）' + note);
    })
    .catch(function (e) {
      console.error('[oxi]', e);
      report('失敗: ' + (e && e.message ? e.message : e));
    });
}

/**
 * Upload the PDF and add it to the record's attachment field, keeping the
 * files already there. The update names the revision the screen was opened
 * at: if someone saved the record since, kintone refuses it (GAIA_CO02)
 * instead of overwriting their change with this screen's older values.
 */
function attachPdf(pdf, fileName, record, fieldCode) {
  var form = new FormData();
  form.append('__REQUEST_TOKEN__', kintone.getRequestToken());
  form.append('file', new Blob([pdf], { type: 'application/pdf' }), fileName);
  return fetch(kintone.api.url('/k/v1/file', true), {
    method: 'POST', body: form, credentials: 'same-origin',
    headers: { 'X-Requested-With': 'XMLHttpRequest' },
  })
    .then(function (res) {
      if (!res.ok) return res.text().then(function (t) { throw new Error('アップロード（HTTP ' + res.status + '）: ' + t); });
      return res.json();
    })
    .then(function (uploaded) {
      var keep = ((record[fieldCode] && record[fieldCode].value) || []).map(function (f) { return { fileKey: f.fileKey }; });
      var body = { app: kintone.app.getId(), id: kintone.app.record.getId(), record: {} };
      if (record.$revision && record.$revision.value) body.revision = record.$revision.value;
      body.record[fieldCode] = { value: keep.concat([{ fileKey: uploaded.fileKey }]) };
      return kintone.api(kintone.api.url('/k/v1/record', true), 'PUT', body);
    });
}

/** What to tell the user when attaching failed. */
function attachError(e) {
  var code = e && e.code;
  if (code === 'GAIA_CO02') {
    return 'このレコードは画面を開いたあとに更新されたため、添付していません。画面を読み込み直してから、もう一度作成してください。';
  }
  if (code === 'CB_NO02' || code === 'GAIA_NO01') {
    return 'このレコードを編集する権限がないため、添付できませんでした。「PDF を保存」で手元に保存できます。';
  }
  return '添付に失敗しました: ' + (e && e.message ? e.message : JSON.stringify(e));
}

function button(label, onclick) {
  var b = document.createElement('button');
  b.type = 'button';
  b.textContent = label;
  b.style.marginRight = '8px';
  b.onclick = onclick;
  return b;
}

kintone.events.on('app.record.detail.show', function (event) {
  var space = kintone.app.record.getHeaderMenuSpaceElement();
  if (!space || space.querySelector('.oxi-buttons')) return event;
  var holder = document.createElement('span');
  holder.className = 'oxi-buttons';
  var record = event.record;
  var config = (PLUGIN_ID && kintone.plugin.app.getConfig(PLUGIN_ID)) || {};
  var output = config.output_field && record[config.output_field] ? config.output_field : '';
  readTemplates(config).forEach(function (t) {
    var source = { name: t.name + '.' + t.kind, kind: t.kind, load: function () { return Promise.resolve(t.bytes); } };
    holder.appendChild(button('帳票: ' + t.name, function () { show(source, record, true, output); }));
  });
  Object.keys(record).forEach(function (code) {
    var field = record[code];
    if (field.type !== 'FILE') return;
    field.value.forEach(function (file) {
      var ext = (file.name.match(/\.(docx|xlsx)$/i) || [])[1];
      if (!ext) return;
      var source = { name: file.name, kind: ext.toLowerCase(), load: function () { return download(file.fileKey); } };
      holder.appendChild(button('表示: ' + file.name, function () { show(source, record, false, ''); }));
      holder.appendChild(button('帳票: ' + file.name, function () { show(source, record, true, output); }));
    });
  });
  space.appendChild(holder);
  return event;
});
})();
