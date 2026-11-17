// SPDX-License-Identifier: MIT OR Apache-2.0
// Fill an Excel template with data, working on the workbook the Oxi engine
// parses (parse_spreadsheet), then recalculate and write it back.
//
// The placeholders are the Word template's: {{name}}, {{list.field}} in a
// row of a list, {{name[-1]}} for one character per cell, {{name:和暦}}.
// A cell that is nothing but one placeholder takes the value as it is, so a
// number stays a number and the cell's number format (#,##0, ¥…) applies;
// placeholders inside longer text are replaced as text.
//
// A row naming {{list.field}} is the first line of the list: entry k goes
// into the row k below it, in the same columns. Rows are not inserted, since
// that would move every formula, merge and print area under them; the
// template provides as many lines as the form prints. Entries beyond the
// last free line are reported, not written over what is below.
import { PLACEHOLDER, NAME, formatValue, pick, baseName } from './oxi-fill.js';

const xlText = (cell) => (cell && cell.value && typeof cell.value === 'object' && typeof cell.value.String === 'string' ? cell.value.String : null);
const xlEmpty = (cell) => !cell || cell.value === 'Empty' || (xlText(cell) === '' && !cell.formula);

/** The value for `name` from `scope`: a number or a string, undefined when absent. */
function xlResolve(scope, name) {
  const [, field, format, index] = name.match(NAME);
  const v = scope[field];
  if (v === undefined || Array.isArray(v)) return undefined;
  if (format) return formatValue(v, format);
  if (index !== undefined) return pick(typeof v === 'number' ? String(v) : String(v), Number(index));
  return v;
}

const xlAsText = (v) => (typeof v === 'number' ? v.toLocaleString('ja-JP') : String(v));

/** Set a cell from its template text; returns the names not found. */
function xlFillCell(cell, template, lookup, missing) {
  const whole = template.match(/^\s*\{\{\s*([^{}]+?)\s*\}\}\s*$/);
  if (whole) {
    const v = lookup(whole[1]);
    if (v === undefined) { missing.add(baseName(whole[1])); cell.value = { String: template }; return; }
    if (typeof v === 'number') cell.value = { Number: v };
    else cell.value = v === '' ? 'Empty' : { String: v };
    return;
  }
  cell.value = {
    String: template.replace(PLACEHOLDER, (m, name) => {
      const v = lookup(name);
      if (v === undefined) { missing.add(baseName(name)); return m; }
      return xlAsText(v);
    }),
  };
}

function xlRowAt(sheet, index) {
  let row = sheet.rows.find((r) => r.index === index);
  if (!row) {
    row = { cells: [], height: null, index };
    sheet.rows.push(row);
    sheet.rows.sort((a, b) => a.index - b.index);
  }
  return row;
}

function xlCellAt(row, col, style) {
  let cell = row.cells.find((c) => c.col === col);
  if (!cell) {
    cell = { col, style: structuredClone(style), value: 'Empty' };
    row.cells.push(cell);
    row.cells.sort((a, b) => a.col - b.col);
  }
  return cell;
}

/**
 * Fill `workbook` (the object parse_spreadsheet returns) with `data`, in
 * place. Returns { missing, overflow }: names the template uses that `data`
 * lacks, and for each list that did not fit, how many entries were left out.
 */
export function fillWorkbook(workbook, data) {
  const missing = new Set();
  const overflow = {};
  const lookup = (name) => xlResolve(data, name);
  for (const sheet of workbook.sheets) {
    const listRows = [];
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        const t = xlText(cell);
        if (t === null) continue;
        for (const m of t.matchAll(PLACEHOLDER)) {
          const dot = m[1].indexOf('.');
          const list = dot > 0 ? m[1].slice(0, dot) : null;
          if (list && Array.isArray(data[list]) && !listRows.some((l) => l.row === row)) listRows.push({ row, list });
        }
      }
    }
    for (const { row, list } of listRows) {
      const templates = row.cells.filter((c) => xlText(c) !== null && xlText(c).includes('{{'))
        .map((c) => ({ col: c.col, style: c.style, text: xlText(c) }));
      const cols = templates.map((t) => t.col);
      const entries = data[list];
      let written = 0;
      for (let k = 0; k < Math.max(entries.length, 1); k++) {
        const target = k === 0 ? row : xlRowAt(sheet, row.index + k);
        if (k > 0 && !cols.every((col) => xlEmpty(target.cells.find((c) => c.col === col)))) break;
        const entry = entries[k];
        for (const t of templates) {
          const cell = xlCellAt(target, t.col, t.style);
          if (!entry) { cell.value = 'Empty'; continue; }
          xlFillCell(cell, t.text, (name) => (name.startsWith(list + '.') ? xlResolve(entry, name.slice(list.length + 1)) : lookup(name)), missing);
        }
        if (entry) written++;
      }
      if (written < entries.length) overflow[list] = entries.length - written;
    }
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        const t = xlText(cell);
        if (t !== null && t.includes('{{')) xlFillCell(cell, t, lookup, missing);
      }
    }
  }
  return { missing: [...missing], overflow };
}

/**
 * Fill an .xlsx template and return the filled workbook's bytes, with every
 * formula recalculated so totals follow the new values. `api` is the
 * initialised Oxi wasm module (suite build).
 */
export function fillXlsx(api, templateBytes, data) {
  const workbook = api.parse_spreadsheet(templateBytes);
  const { missing, overflow } = fillWorkbook(workbook, data);
  const recalculated = api.recalculate_spreadsheet(workbook, Date.now());
  const xlsx = api.edit_xlsx_from_workbook(templateBytes, recalculated);
  return { xlsx, missing, overflow };
}
