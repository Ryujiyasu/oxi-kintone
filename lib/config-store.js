// SPDX-License-Identifier: MIT OR Apache-2.0
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
