// node fill-json.mjs <wasm pkg dir> <template.docx> <data.json> <out.pdf>
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fillDocx } from '../lib/oxi-fill.js';
const [pkg, tplPath, dataPath, out] = process.argv.slice(2);
const api = await import(pathToFileURL(resolve(pkg, 'oxidocs_wasm.js')).href);
await api.default({ module_or_path: readFileSync(resolve(pkg, 'oxidocs_wasm_bg.wasm')) });
const t0 = performance.now();
const { docx, missing } = fillDocx(api, readFileSync(tplPath), JSON.parse(readFileSync(dataPath, 'utf8')));
writeFileSync(out.replace(/\.pdf$/, '.docx'), docx);
writeFileSync(out, api.docx_to_pdf(docx));
console.log('missing:', missing, 'ms:', Math.round(performance.now() - t0), '->', out);
