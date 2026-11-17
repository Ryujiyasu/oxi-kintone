// node try.mjs <wasm pkg> <outdir>: fill the three templates with sample-record and write PDFs
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fillDocx, kintoneRecordData, kintoneRecordValues } from '../lib/oxi-fill.js';
import { fillXlsx } from '../lib/oxi-fill-xlsx.js';
import { record } from './sample-record.mjs';
const [pkg, out] = process.argv.slice(2);
const api = await import(pathToFileURL(resolve(pkg, 'oxidocs_wasm.js')).href);
await api.default({ module_or_path: readFileSync(resolve(pkg, 'oxidocs_wasm_bg.wasm')) });
for (const name of ['見積書.docx', '納品書.docx', '請求書.xlsx']) {
  const tpl = readFileSync(name), title = name.replace(/\..*$/, '');
  const filled = name.endsWith('.xlsx') ? fillXlsx(api, tpl, kintoneRecordValues(record)) : fillDocx(api, tpl, kintoneRecordData(record));
  const pdf = name.endsWith('.xlsx') ? api.xlsx_to_pdf(filled.xlsx, title) : api.docx_to_pdf_with_title(filled.docx, title);
  writeFileSync(resolve(out, title + '.pdf'), pdf);
  console.log(name, 'missing', filled.missing, filled.overflow || '', pdf.length, 'bytes');
}
