// node roundtrip.mjs <pkg> <in.docx> <outprefix>: original PDF vs model round-trip PDF
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
const [pkg, input, out] = process.argv.slice(2);
const api = await import(pathToFileURL(resolve(pkg, 'oxidocs_wasm.js')).href);
await api.default({ module_or_path: readFileSync(resolve(pkg, 'oxidocs_wasm_bg.wasm')) });
const bytes = readFileSync(input);
writeFileSync(out + '-orig.pdf', api.docx_to_pdf(bytes));
const json = api.docx_to_model(bytes);
writeFileSync(out + '-model.json', json);
const rt = api.docx_from_model(json, bytes);
writeFileSync(out + '-rt.docx', rt);
writeFileSync(out + '-rt.pdf', api.docx_to_pdf(rt));
console.log('ok', input);
