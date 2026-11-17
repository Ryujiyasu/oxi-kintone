// SPDX-License-Identifier: MIT OR Apache-2.0
// Built by build.sh from lib/oxi-fill.js and lib/plugin-main.js.
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
// {{list.field}} reads the entry's field. An empty array removes the row.

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

/** Replace the placeholders of one paragraph, in place. */
function fillParagraph(paragraph, lookup) {
  const items = textItems(paragraph);
  if (items.length === 0) return;
  const starts = [];
  let full = '';
  for (const item of items) { starts.push(full.length); full += item.Text; }
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

function walk(node, lookup, data) {
  if (node == null || typeof node !== 'object') return;
  if (Array.isArray(node)) { node.forEach((n) => walk(n, lookup, data)); return; }
  if (node.Table) {
    const rows = [];
    for (const item of node.Table.rows) {
      const list = item.Row ? rowList(item.Row, data) : null;
      if (!list) { rows.push(item); continue; }
      for (const entry of data[list]) {
        const copy = structuredClone(item);
        const inner = (name) => {
          if (name.startsWith(list + '.')) {
            const v = entry[name.slice(list.length + 1)];
            return v === undefined ? undefined : String(v);
          }
          return lookup(name);
        };
        walk(copy, inner, data);
        rows.push(copy);
      }
    }
    node.Table.rows = rows;
    for (const item of rows) walk(item, lookup, data);
    return;
  }
  if (node.Paragraph) fillParagraph(node.Paragraph, lookup);
  for (const value of Object.values(node)) walk(value, lookup, data);
}

/**
 * Fill `model` (the parsed JSON of docx_to_model) with `data`, in place.
 * `data` maps names to strings or numbers, and list names to arrays of such
 * maps. Returns the names used in the template that `data` does not have.
 */
function fillModel(model, data) {
  const missing = new Set();
  const lookup = (name) => {
    const v = data[name];
    if (v === undefined || Array.isArray(v)) { missing.add(name); return undefined; }
    return String(v);
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
// Oxi for kintone (prototype). On the record detail screen, for each attached
// .docx it offers:
//   - 表示: convert the file to PDF and show it;
//   - 帳票: treat the file as a template, fill its {{placeholders}} with this
//     record (field codes or the labels shown on the form; a table row with
//     {{subtable.field}} repeats per subtable row), and show the PDF.
// The Oxi engine (WebAssembly) is loaded from the public npm package via
// unpkg (jsDelivr refuses files over 50 MB); documents and record data never
// leave the browser.
var ENGINE = 'https://unpkg.com/@oxidocs/wasm@0.9.0/oxidocs_wasm.js';
if (typeof window !== 'undefined' && window.OXI_ENGINE_URL) ENGINE = window.OXI_ENGINE_URL;
var enginePromise = null;

function loadEngine(report) {
  if (!enginePromise) {
    var t0 = performance.now();
    report('Oxi エンジンを読み込み中…（初回は約 56MB）');
    enginePromise = import(ENGINE).then(function (mod) {
      return mod.default().then(function () {
        console.log('[oxi] engine ready in ' + Math.round(performance.now() - t0) + ' ms');
        return mod;
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

function overlay() {
  var box = document.getElementById('oxi-preview-overlay');
  if (box) return box;
  box = document.createElement('div');
  box.id = 'oxi-preview-overlay';
  box.style.cssText = 'position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.6);display:flex;flex-direction:column;padding:24px;box-sizing:border-box';
  var bar = document.createElement('div');
  bar.style.cssText = 'display:flex;gap:12px;align-items:center;color:#fff;margin-bottom:8px;font-size:14px';
  var status = document.createElement('span');
  status.className = 'oxi-status';
  status.style.flex = '1';
  var save = document.createElement('a');
  save.className = 'oxi-save';
  save.textContent = 'PDF を保存';
  save.style.cssText = 'color:#fff;display:none';
  var close = document.createElement('button');
  close.type = 'button';
  close.textContent = '閉じる';
  close.onclick = function () { box.remove(); };
  bar.appendChild(status);
  bar.appendChild(save);
  bar.appendChild(close);
  var frame = document.createElement('iframe');
  frame.className = 'oxi-frame';
  frame.style.cssText = 'flex:1;border:0;background:#fff';
  box.appendChild(bar);
  box.appendChild(frame);
  document.body.appendChild(box);
  return box;
}

function show(file, record, fill) {
  var box = overlay();
  var status = box.querySelector('.oxi-status');
  var frame = box.querySelector('.oxi-frame');
  var save = box.querySelector('.oxi-save');
  function report(text) { status.textContent = file.name + ' — ' + text; }
  report('準備中…');
  Promise.all([loadEngine(report), download(file.fileKey), fill ? formLabels() : Promise.resolve({})])
    .then(function (r) {
      var api = r[0], bytes = new Uint8Array(r[1]), note = '';
      var t0 = performance.now();
      if (fill) {
        report('帳票を作成中…');
        var filled = fillDocx(api, bytes, withLabels(kintoneRecordData(record), r[2]));
        bytes = filled.docx;
        if (filled.missing.length) note = '　見つからない項目: ' + filled.missing.join('、');
      }
      report('PDF に変換中…');
      var pdf = api.docx_to_pdf(bytes);
      var ms = Math.round(performance.now() - t0);
      var url = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }));
      frame.src = url;
      save.href = url;
      save.download = file.name.replace(/\.docx$/i, '') + (fill ? '_帳票' : '') + '.pdf';
      save.style.display = '';
      report('ブラウザ内で作成しました（' + ms + ' ms）' + note);
    })
    .catch(function (e) {
      console.error('[oxi]', e);
      report('失敗: ' + (e && e.message ? e.message : e));
    });
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
  Object.keys(record).forEach(function (code) {
    var field = record[code];
    if (field.type !== 'FILE') return;
    field.value.forEach(function (file) {
      if (!/\.docx$/i.test(file.name)) return;
      holder.appendChild(button('表示: ' + file.name, function () { show(file, record, false); }));
      holder.appendChild(button('帳票: ' + file.name, function () { show(file, record, true); }));
    });
  });
  space.appendChild(holder);
  return event;
});
})();
