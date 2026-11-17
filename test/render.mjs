import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
const [pkg, input, out] = process.argv.slice(2);
const api = await import(pathToFileURL(resolve(pkg, 'oxidocs_wasm.js')).href);
await api.default({ module_or_path: readFileSync(resolve(pkg, 'oxidocs_wasm_bg.wasm')) });
writeFileSync(out, api.docx_to_pdf(readFileSync(input)));
