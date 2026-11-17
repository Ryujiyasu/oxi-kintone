// SPDX-License-Identifier: MIT OR Apache-2.0
// Oxi for kintone (prototype). On the record detail screen it offers:
//   - 帳票: <name> for each Word or Excel template registered in the plugin
//     settings: fill its {{placeholders}} with this record (field codes or
//     the labels shown on the form; a row with {{subtable.field}} takes the
//     subtable's rows), show the PDF, and optionally attach it to the record;
//   - 表示 / 帳票 for each attached .docx or .xlsx, as is or filled.
// The Oxi engine (WebAssembly) is loaded from the public npm package via
// unpkg (jsDelivr refuses files over 50 MB); documents and record data never
// leave the browser.
var ENGINE = 'https://unpkg.com/@oxidocs/wasm@__ENGINE_VERSION__/oxidocs_wasm.js';
if (typeof window !== 'undefined' && window.OXI_ENGINE_URL) ENGINE = window.OXI_ENGINE_URL;
var enginePromise = null;
var PLUGIN_ID = kintone.$PLUGIN_ID;

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
      var pdf = excel ? api.xlsx_to_pdf(bytes, title)
        : api.docx_to_pdf_with_title ? api.docx_to_pdf_with_title(bytes, title) : api.docx_to_pdf(bytes);
      var ms = Math.round(performance.now() - t0);
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
