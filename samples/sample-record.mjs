// A record shaped like the sample app's, for trying the templates outside kintone.
const T = (type, value) => ({ type, value });
const row = (item, qty, unit, price) => ({ value: {
  品名: T('SINGLE_LINE_TEXT', item), 数量: T('NUMBER', qty), 単位: T('SINGLE_LINE_TEXT', unit),
  単価: T('NUMBER', price), 金額: T('CALC', String(Number(qty) * Number(price))) } });
const rows = [row('ノートPC 14インチ', '10', '台', '120000'), row('初期設定作業', '10', '台', '8000'), row('保守サポート（1年）', '1', '式', '106000')];
const subtotal = rows.reduce((s, r) => s + Number(r.value.金額.value), 0), tax = Math.floor(subtotal * 0.1);
export const record = {
  $id: T('__ID__', '1'), $revision: T('__REVISION__', '1'),
  文書番号: T('SINGLE_LINE_TEXT', 'Q-2026-0001'), 顧客名: T('SINGLE_LINE_TEXT', '株式会社テスト商事'),
  顧客担当者: T('SINGLE_LINE_TEXT', '山田 花子'), 件名: T('SINGLE_LINE_TEXT', '社内研修用ノートPCの導入'),
  発行日: T('DATE', '2026-12-02'), 支払期限: T('DATE', '2026-12-31'), 納品日: T('DATE', '2026-12-10'),
  明細: T('SUBTABLE', rows), 小計: T('CALC', String(subtotal)), 消費税: T('CALC', String(tax)), 合計: T('CALC', String(subtotal + tax)),
  備考: T('MULTI_LINE_TEXT', '納期はご発注後2週間です。'), 帳票: T('FILE', []),
};
