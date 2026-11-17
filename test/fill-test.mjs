// node fill-test.mjs <wasm pkg dir> <template.docx> <out.pdf>
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fillDocx, kintoneRecordData } from '../lib/oxi-fill.js';
const [pkg, tplPath, out] = process.argv.slice(2);
const api = await import(pathToFileURL(resolve(pkg, 'oxidocs_wasm.js')).href);
await api.default({ module_or_path: readFileSync(resolve(pkg, 'oxidocs_wasm_bg.wasm')) });
const row = (item, qty, price, amount) => ({ value: {
  品名: { type: 'SINGLE_LINE_TEXT', value: item }, 数量: { type: 'NUMBER', value: qty },
  単価: { type: 'NUMBER', value: price }, 金額: { type: 'CALC', value: amount } } });
const record = {
  顧客名: { type: 'SINGLE_LINE_TEXT', value: '株式会社テスト商事' }, 見積日: { type: 'DATE', value: '2026-11-17' },
  件名: { type: 'SINGLE_LINE_TEXT', value: '社内研修用ノートPCの導入' }, 備考: { type: 'MULTI_LINE_TEXT', value: '納期はご発注後2週間です。' },
  合計: { type: 'CALC', value: '1386000' },
  明細: { type: 'SUBTABLE', value: [row('ノートPC 14インチ', '10', '120000', '1200000'), row('初期設定作業', '10', '8000', '80000'), row('保守サポート（1年）', '1', '106000', '106000')] },
};
const { docx, missing } = fillDocx(api, readFileSync(tplPath), kintoneRecordData(record));
writeFileSync(out, api.docx_to_pdf(docx));
console.log('missing:', missing, '->', out);
