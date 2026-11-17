// SPDX-License-Identifier: MIT OR Apache-2.0
// Built by build.sh from lib/. Do not edit by hand.
(function () {
'use strict';
// Plugin settings layout (kintone keeps settings as string values, at most
// 65,535 characters each and 256 KB in total):
//   templates: JSON [{ name, parts, size }]
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
    return { name: t.name, bytes: base64ToBytes(b64) };
  });
}

/** A setConfig() object for these templates and output field. */
function writeTemplates(templates, outputField) {
  var config = { output_field: outputField || '' }, meta = [];
  templates.forEach(function (t, i) {
    var b64 = bytesToBase64(t.bytes), parts = 0;
    for (var o = 0; o < b64.length; o += CHUNK) config['tpl_' + i + '_' + parts++] = b64.slice(o, o + CHUNK);
    meta.push({ name: t.name, parts: parts, size: t.bytes.length });
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
// Plugin settings screen: register Word templates and pick the attachment
// field that receives generated PDFs.
var PLUGIN_ID = kintone.$PLUGIN_ID;
var LIMIT = 250 * 1024; // kintone allows 256 KB of settings in total
var root = document.querySelector('.oxi-config');
var current = kintone.plugin.app.getConfig(PLUGIN_ID) || {};
var templates = readTemplates(current);
var errorBox = root.querySelector('.oxi-error');
var usage = root.querySelector('.oxi-usage');

function render() {
  var tbody = root.querySelector('.oxi-templates tbody');
  tbody.textContent = '';
  templates.forEach(function (t, i) {
    var tr = document.createElement('tr');
    var name = document.createElement('td'); name.textContent = t.name; name.style.padding = '4px 8px';
    var size = document.createElement('td'); size.textContent = Math.ceil(t.bytes.length / 1024) + ' KB';
    size.style.cssText = 'padding:4px 8px;text-align:right';
    var del = document.createElement('td');
    var b = document.createElement('button'); b.type = 'button'; b.textContent = '削除';
    b.onclick = function () { templates.splice(i, 1); render(); };
    del.appendChild(b);
    tr.appendChild(name); tr.appendChild(size); tr.appendChild(del);
    tbody.appendChild(tr);
  });
  var used = configBytes(writeTemplates(templates, root.querySelector('.oxi-output').value));
  usage.textContent = '設定の使用量: ' + Math.ceil(used / 1024) + ' KB / ' + Math.floor(LIMIT / 1024) + ' KB';
}

root.querySelector('.oxi-add').onclick = function () {
  errorBox.textContent = '';
  var input = root.querySelector('.oxi-new-file');
  var file = input.files && input.files[0];
  if (!file) { errorBox.textContent = 'ひな形の Word ファイル（.docx）を選んでください。'; return; }
  if (!/\.docx$/i.test(file.name)) { errorBox.textContent = '.docx ファイルを選んでください。'; return; }
  var name = root.querySelector('.oxi-new-name').value.trim() || file.name.replace(/\.docx$/i, '');
  file.arrayBuffer().then(function (buf) {
    var next = templates.concat([{ name: name, bytes: new Uint8Array(buf) }]);
    if (configBytes(writeTemplates(next, root.querySelector('.oxi-output').value)) > LIMIT) {
      errorBox.textContent = 'このひな形を加えると、設定の上限（256 KB）を超えます。画像を小さくするか、ほかのひな形を削除してください。';
      return;
    }
    templates = next;
    input.value = '';
    root.querySelector('.oxi-new-name').value = '';
    render();
  });
};

root.querySelector('.oxi-save').onclick = function () {
  var config = writeTemplates(templates, root.querySelector('.oxi-output').value);
  if (configBytes(config) > LIMIT) { errorBox.textContent = '設定が上限（256 KB）を超えています。'; return; }
  kintone.plugin.app.setConfig(config, function () {
    window.location.href = '../../flow?app=' + kintone.app.getId();
  });
};
root.querySelector('.oxi-cancel').onclick = function () { window.history.back(); };

kintone.api(kintone.api.url('/k/v1/preview/app/form/fields', true), 'GET', { app: kintone.app.getId() })
  .then(function (resp) {
    var select = root.querySelector('.oxi-output');
    Object.keys(resp.properties).forEach(function (code) {
      var p = resp.properties[code];
      if (p.type !== 'FILE') return;
      var o = document.createElement('option');
      o.value = code; o.textContent = p.label + '（' + code + '）';
      select.appendChild(o);
    });
    select.value = current.output_field || '';
    select.onchange = render;
    render();
  })
  .catch(function () { render(); });
})();
