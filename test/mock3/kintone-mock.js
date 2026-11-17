// A stand-in for the parts of the kintone JS API the plugin uses.
window.__calls = [];
var props = {
  顧客名: { type: 'SINGLE_LINE_TEXT', label: '顧客名' }, 見積日: { type: 'DATE', label: '見積日' },
  件名: { type: 'SINGLE_LINE_TEXT', label: '件名' }, 備考: { type: 'MULTI_LINE_TEXT', label: '備考' },
  合計: { type: 'CALC', label: '合計' }, 添付ファイル: { type: 'FILE', label: '添付ファイル' },
  明細: { type: 'SUBTABLE', label: '明細', fields: { 品名: { label: '品名' }, 数量: { label: '数量' }, 単価: { label: '単価' }, 金額: { label: '金額' } } }
};
var handlers = {};
window.kintone = {
  $PLUGIN_ID: 'mockplugin',
  events: { on: function (t, f) { handlers[t] = f; } },
  getRequestToken: function () { return 'tok'; },
  api: Object.assign(function (url, method, body) {
    window.__calls.push({ url: url, method: method, body: body });
    if (method === 'PUT') localStorage.setItem('put', JSON.stringify({ url: url, body: body }));
    if (/form\/fields/.test(url)) return Promise.resolve({ properties: props });
    return Promise.resolve({ revision: '2' });
  }, { url: function (p) { return p + '.json'; } }),
  app: { getId: function () { return 1; }, record: { getId: function () { return 7; }, getHeaderMenuSpaceElement: function () { return document.getElementById('hdr'); } } },
  plugin: { app: {
    getConfig: function () { try { return JSON.parse(localStorage.getItem('cfg') || '{}'); } catch (e) { return {}; } },
    setConfig: function (c, cb) { localStorage.setItem('cfg', JSON.stringify(c)); window.__saved = c; }
  } }
};
window.__handlers = handlers;
