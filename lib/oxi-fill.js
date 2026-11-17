// SPDX-License-Identifier: MIT OR Apache-2.0
// Fill a Word template with data, working on the Oxi document model (the JSON
// that docx_to_model returns and docx_from_model writes back).
//
// Placeholders are written in the template as {{name}}. A placeholder may be
// split across runs (Word does this when part of it is formatted differently
// or was typed in two goes); the replacement keeps the formatting of the run
// where the placeholder starts.
//
// A table row that mentions {{list.field}}, where `list` is an array in the
// data, is repeated once per entry of that array; inside the copies,
// {{list.field}} reads the entry's field. Blank rows right under it (the
// empty lines of a printed form) are used up first, so the table keeps its
// height until there are more entries than lines; with no entries the row
// stays, emptied.
//
// {{name[3]}} is the third character of the value and {{name[-1]}} the last
// one, for forms that print one character per box (amounts, registration
// numbers); {{name:和暦}} writes a date as 令和8年11月17日. Placeholders inside text boxes are filled too; headers and
// footers are not part of the model and are left as they are.

const PLACEHOLDER = /\{\{\s*([^{}]+?)\s*\}\}/g;

/** Every {"Text": ...} item held directly by the runs of a paragraph. */
function textItems(paragraph) {
  const items = [];
  for (const inline of paragraph.content || []) {
    const run = inline.Run;
    if (!run) continue;
    for (const item of run.content || []) {
      if (typeof item.Text === 'string') items.push(item);
    }
  }
  return items;
}

/**
 * The w:t elements of a paragraph kept as raw XML (text boxes and other
 * drawings), wrapped so they read and write like {"Text": ...} items.
 */
function rawTextItems(p) {
  const items = [];
  for (const r of p.children || []) {
    if (!r || r.name !== 'w:r') continue;
    for (const t of r.children || []) {
      if (!t || t.name !== 't' && t.name !== 'w:t') continue;
      items.push({
        get Text() { return (t.children || []).filter((c) => typeof c === 'string').join(''); },
        set Text(v) {
          t.children = v === '' ? [] : [v];
          const attrs = t.attrs || (t.attrs = []);
          if (!attrs.some((a) => a[0] === 'xml:space')) attrs.push(['xml:space', 'preserve']);
        },
      });
    }
  }
  return items;
}

/** Replace the placeholders spread over `items`, in place. */
function fillItems(items, lookup) {
  if (items.length === 0) return;
  const starts = [];
  let full = '';
  for (const item of items) { starts.push(full.length); full += item.Text; }
  if (!full.includes('{{')) return;
  const matches = [...full.matchAll(PLACEHOLDER)];
  // Right to left, so the offsets of earlier placeholders stay valid.
  for (const m of matches.reverse()) {
    const value = lookup(m[1]);
    if (value === undefined) continue; // unknown name: leave it visible
    const start = m.index, end = m.index + m[0].length;
    const segAt = (pos) => {
      let i = 0;
      while (i + 1 < items.length && starts[i + 1] <= pos) i++;
      return i;
    };
    const s0 = segAt(start), s1 = segAt(end - 1);
    const a = items[s0].Text, from = start - starts[s0];
    if (s0 === s1) {
      items[s0].Text = a.slice(0, from) + value + a.slice(end - starts[s0]);
    } else {
      items[s0].Text = a.slice(0, from) + value;
      for (let i = s0 + 1; i < s1; i++) items[i].Text = '';
      items[s1].Text = items[s1].Text.slice(end - starts[s1]);
    }
  }
}

/** All text of a subtree, to see which placeholders a row mentions. */
function textOf(node) {
  if (node == null) return '';
  if (typeof node === 'string') return '';
  if (Array.isArray(node)) return node.map(textOf).join('');
  if (typeof node.Text === 'string') return node.Text;
  if (node.name === 'w:t' || node.name === 't') {
    return (node.children || []).filter((c) => typeof c === 'string').join('');
  }
  return Object.values(node).map(textOf).join('');
}

/** The list a row repeats over: the first {{list.field}} naming an array. */
function rowList(row, data) {
  for (const m of textOf(row).matchAll(PLACEHOLDER)) {
    const dot = m[1].indexOf('.');
    if (dot > 0 && Array.isArray(data[m[1].slice(0, dot)])) return m[1].slice(0, dot);
  }
  return null;
}

/** A row with no text at all: a blank line of a printed form. */
function blankRow(item) {
  return !!item.Row && textOf(item.Row).trim() === '';
}

/** A copy of a repeating row with every placeholder emptied. */
function clearedRow(item) {
  const copy = structuredClone(item);
  walk(copy, () => '', {});
  return copy;
}

function walk(node, lookup, data) {
  if (node == null || typeof node !== 'object') return;
  if (Array.isArray(node)) { node.forEach((n) => walk(n, lookup, data)); return; }
  if (node.Table) {
    const src = node.Table.rows;
    const rows = [];
    for (let i = 0; i < src.length; i++) {
      const item = src[i];
      const list = item.Row ? rowList(item.Row, data) : null;
      if (!list) { rows.push(item); continue; }
      // A printed form leaves blank rows under the first line of a list. The
      // entries go into those rows first, so the table keeps its height; only
      // more entries than lines make it grow.
      let blanks = 0;
      while (i + 1 + blanks < src.length && blankRow(src[i + 1 + blanks])) blanks++;
      const entries = data[list];
      for (const entry of entries) {
        const copy = structuredClone(item);
        const inner = (name) => {
          if (name.startsWith(list + '.')) {
            const v = entryValue(entry, name.slice(list.length + 1));
            return v === undefined ? undefined : v;
          }
          return lookup(name);
        };
        walk(copy, inner, data);
        rows.push(copy);
      }
      const free = blanks + 1 - entries.length;
      if (free > 0) {
        if (entries.length === 0) rows.push(clearedRow(item));
        for (let k = blanks - free + (entries.length === 0 ? 1 : 0); k < blanks; k++) rows.push(src[i + 1 + k]);
      }
      i += blanks;
    }
    node.Table.rows = rows;
    for (const item of rows) walk(item, lookup, data);
    return;
  }
  if (node.Paragraph) fillItems(textItems(node.Paragraph), lookup);
  else if (node.name === 'w:p' && Array.isArray(node.children)) fillItems(rawTextItems(node), lookup);
  for (const value of Object.values(node)) walk(value, lookup, data);
}

/**
 * A name may carry a format and an index: {{name:和暦}}, {{name[-1]}}.
 *
 * `:和暦` writes a date (2026-11-17) as 令和8年11月17日 and `:年月日` as
 * 2026年11月17日; values that are not dates are left as they are.
 *
 * `[i]` is the i-th character of the value (from 1) and `[-i]` the i-th from
 * the end: one character per box, as printed forms ask for amounts and
 * registration numbers. Separators (",") are left out of the count.
 */
const NAME = /^(.+?)(?::(和暦|年月日))?(?:\[(-?\d+)\])?$/;
const DATE = /^(\d{4})-(\d{2})-(\d{2})/;
function formatValue(value, format) {
  const d = String(value).match(DATE);
  if (!format || !d) return String(value);
  const [y, m, day] = [Number(d[1]), Number(d[2]), Number(d[3])];
  if (format === '年月日') return `${y}年${m}月${day}日`;
  const ymd = y * 10000 + m * 100 + day;
  const [era, first] = ymd >= 20190501 ? ['令和', 2019] : ymd >= 19890108 ? ['平成', 1989] : ['昭和', 1926];
  const n = y - first + 1;
  return `${era}${n === 1 ? '元' : n}年${m}月${day}日`;
}
function pick(value, index) {
  const chars = [...value.replace(/,/g, '')];
  const at = index > 0 ? index - 1 : chars.length + index;
  return index !== 0 && at >= 0 && at < chars.length ? chars[at] : '';
}
/** The base field name of a placeholder, without format or index. */
function baseName(name) {
  return name.match(NAME)[1];
}
function entryValue(entry, name) {
  const [, field, format, index] = name.match(NAME);
  const v = entry[field];
  if (v === undefined || Array.isArray(v)) return undefined;
  const text = formatValue(v, format);
  return index === undefined ? text : pick(text, Number(index));
}

/**
 * Fill `model` (the parsed JSON of docx_to_model) with `data`, in place.
 * `data` maps names to strings or numbers, and list names to arrays of such
 * maps. Returns the names used in the template that `data` does not have.
 */
export function fillModel(model, data) {
  const missing = new Set();
  const lookup = (name) => {
    const v = entryValue(data, name);
    if (v === undefined) missing.add(baseName(name));
    return v;
  };
  walk(model, lookup, data);
  return [...missing];
}

/** Fill a .docx template; `api` is the initialised Oxi wasm module. */
export function fillDocx(api, templateBytes, data) {
  const model = JSON.parse(api.docx_to_model(templateBytes));
  const missing = fillModel(model, data);
  const docx = api.docx_from_model(JSON.stringify(model), templateBytes);
  return { docx, missing };
}

/**
 * Turn a kintone record (event.record or REST record) into fill data:
 * numbers are formatted with thousands separators, subtables become arrays.
 */
export function kintoneRecordData(record) {
  const fmt = (field) => {
    if (field == null) return '';
    const { type, value } = field;
    if (value == null || value === '') return '';
    if (type === 'NUMBER' || type === 'CALC') {
      const n = Number(value);
      return Number.isFinite(n) ? n.toLocaleString('ja-JP') : String(value);
    }
    if (Array.isArray(value)) {
      return value.map((v) => (typeof v === 'object' ? v.name ?? v.code ?? '' : v)).join('、');
    }
    if (typeof value === 'object') return value.name ?? value.code ?? '';
    return String(value);
  };
  const data = {};
  for (const [code, field] of Object.entries(record)) {
    if (field.type === 'SUBTABLE') {
      data[code] = field.value.map((row) => {
        const entry = {};
        for (const [c, f] of Object.entries(row.value)) entry[c] = fmt(f);
        return entry;
      });
    } else if (field.type !== 'FILE') {
      data[code] = fmt(field);
    }
  }
  return data;
}
