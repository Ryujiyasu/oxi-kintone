// node fill-xlsx.mjs <pkg> <template.xlsx> <out.xlsx> [out.pdf]
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fillXlsx } from '../../lib/oxi-fill-xlsx.js';
import { kintoneRecordValues } from '../../lib/oxi-fill.js';
const [pkg, tpl, out, pdf] = process.argv.slice(2);
const api = await import(pathToFileURL(resolve(pkg, 'oxidocs_wasm.js')).href);
await api.default({ module_or_path: readFileSync(resolve(pkg, 'oxidocs_wasm_bg.wasm')) });
const T = (type, value) => ({ type, value });
const row = (no, item, qty, price) => ({ value: { No: T('NUMBER', no), 品名: T('SINGLE_LINE_TEXT', item), 数量: T('NUMBER', qty), 単価: T('NUMBER', price) } });
const record = {
  顧客名: T('SINGLE_LINE_TEXT', '株式会社テスト商事'), 請求日: T('DATE', '2026-11-18'), 請求番号: T('SINGLE_LINE_TEXT', 'INV-2026-0042'),
  件名: T('SINGLE_LINE_TEXT', '社内研修用ノートPCの導入'), 備考: T('MULTI_LINE_TEXT', 'お振込手数料は貴社ご負担でお願いいたします。'),
  明細: T('SUBTABLE', [row('1', 'ノートPC 14インチ', '10', '120000'), row('2', '初期設定作業', '10', '8000'), row('3', '保守サポート（1年）', '1', '106000')]),
};
const t0 = performance.now();
const { xlsx, missing, overflow } = fillXlsx(api, readFileSync(tpl), kintoneRecordValues(record));
writeFileSync(out, xlsx);
let note = '';
if (pdf && api.xlsx_to_pdf) { writeFileSync(pdf, api.xlsx_to_pdf(xlsx, '請求書')); note = ' + pdf'; }
console.log('missing', missing, 'overflow', overflow, Math.round(performance.now() - t0), 'ms' + note);
const wb = api.parse_spreadsheet(xlsx);
for (const r of wb.sheets[0].rows) {
  const vals = r.cells.filter((c) => c.value !== 'Empty' && !(c.value.String === '')).map((c) => `${'ABCDE'[c.col]}${r.index}=${JSON.stringify(c.value)}`);
  if (vals.length) console.log(vals.join('  '));
}
