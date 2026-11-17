# Build the sample templates for the 5-minute kit, matching the fields of the
# sample app (samples/app-fields.json):
#   見積書.docx, 納品書.docx (Word) and 請求書.xlsx (Excel).
# The issuer block (株式会社サンプル …) is meant to be edited in Word / Excel.
# openpyxl stores no formula results, so 請求書.xlsx is saved once more through
# LibreOffice afterwards (see README in this folder).
import sys, os
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

OUT = sys.argv[1] if len(sys.argv) > 1 else '.'
ISSUER = ['株式会社サンプル', '〒000-0000 ○○県○○市○○町1-2-3', 'TEL 00-0000-0000', '登録番号 T0000000000000']
LINES = 10  # detail lines printed on each form

def word(title, date_label, date_field, lead, amount_label, extra_rows, name):
    d = Document()
    sec = d.sections[0]; sec.page_width, sec.page_height = Mm(210), Mm(297)
    for m in ('left_margin', 'right_margin'): setattr(sec, m, Mm(18))
    sec.top_margin = Mm(15); sec.bottom_margin = Mm(15)
    st = d.styles['Normal']; st.font.name = 'MS Mincho'; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'ＭＳ 明朝')
    st.paragraph_format.space_after = Pt(0)
    def para(text='', size=None, bold=False, align=None, gothic=False, underline=False):
        p = d.add_paragraph()
        if text:
            r = p.add_run(text); r.bold = bold; r.underline = underline
            if size: r.font.size = Pt(size)
            if gothic: r.font.name = 'MS Gothic'; r._element.rPr.rFonts.set(qn('w:eastAsia'), 'ＭＳ ゴシック')
        if align is not None: p.alignment = align
        return p
    para(title, 20, True, AL.CENTER, gothic=True)
    para()
    para('No. {{文書番号}}', align=AL.RIGHT)
    para(f'{date_label}：{{{{{date_field}:和暦}}}}', align=AL.RIGHT)
    para('{{顧客名}}　御中', 14, underline=True)
    para('{{顧客担当者}}　様')
    para()
    para('件名：{{件名}}')
    para(lead)
    para(f'{amount_label}　¥{{{{合計}}}}-（税込）', 14, True, underline=True)
    para()
    t = d.add_table(rows=2 + (LINES - 1) + 3, cols=5); t.style = 'Table Grid'
    widths = [Mm(70), Mm(18), Mm(14), Mm(28), Mm(32)]
    # Fixed layout with the widths in the table grid as well as on each cell,
    # so every renderer (not only Word) gives 品名 the wide column.
    t.autofit = False
    for i, col in enumerate(t._tbl.tblGrid.findall(qn('w:gridCol'))): col.set(qn('w:w'), str(int(widths[i].twips)))
    for row in t.rows:
        for i, c in enumerate(row.cells): c.width = widths[i]
    for i, h in enumerate(['品名', '数量', '単位', '単価', '金額']):
        t.cell(0, i).text = h; t.cell(0, i).paragraphs[0].alignment = AL.CENTER
    for i, h in enumerate(['{{明細.品名}}', '{{明細.数量}}', '{{明細.単位}}', '{{明細.単価}}', '{{明細.金額}}']):
        t.cell(1, i).text = h
    base = 1 + LINES
    for k, (label, field) in enumerate([('小計', '小計'), ('消費税（10%）', '消費税'), ('合計', '合計')]):
        r = base + k
        t.cell(r, 0).merge(t.cell(r, 3)).text = label
        t.cell(r, 0).paragraphs[0].alignment = AL.RIGHT
        t.cell(r, 4).text = '{{%s}}' % field
    for row in t.rows:
        for i in (1, 3, 4):
            for pp in row.cells[i].paragraphs: pp.alignment = AL.RIGHT
    para()
    for line in extra_rows: para(line)
    para('備考：{{備考}}')
    para()
    for line in ISSUER: para(line, align=AL.RIGHT)
    d.save(os.path.join(OUT, name))

word('御　見　積　書', '発行日', '発行日', '下記の通りお見積り申し上げます。', '御見積金額', ['有効期限：発行日より30日'], '見積書.docx')
word('納　品　書', '納品日', '納品日', '下記の通り納品いたしました。', '合計金額', [], '納品書.docx')

# 請求書.xlsx
F = 'ＭＳ Ｐゴシック'
thin = Side(style='thin'); box = Border(left=thin, right=thin, top=thin, bottom=thin)
wb = Workbook(); ws = wb.active; ws.title = '請求書'
for col, w in zip('ABCDEF', [30, 9, 7, 13, 15, 2]): ws.column_dimensions[col].width = w
def put(ref, value, size=11, bold=False, underline=None, align=None, fmt=None, border=False):
    c = ws[ref]; c.value = value; c.font = Font(name=F, size=size, bold=bold, underline=underline)
    if align: c.alignment = Alignment(horizontal=align)
    if fmt: c.number_format = fmt
    if border: c.border = box
    return c
ws.merge_cells('A1:E1'); put('A1', '請　求　書', 20, True, align='center')
put('D3', 'No.'); put('E3', '{{文書番号}}', align='right')
put('D4', '発行日'); put('E4', '{{発行日:和暦}}', align='right')
put('A3', '{{顧客名}}　御中', 14, underline='single')
put('A4', '{{顧客担当者}}　様')
put('A6', '件名：{{件名}}')
put('A7', '下記の通りご請求申し上げます。')
put('A9', 'ご請求金額', 12, True)
ws.merge_cells('B9:D9'); put('B9', '=E25', 16, True, 'single', 'center', '"¥"#,##0"-（税込）"')
put('A10', 'お支払期限：{{支払期限:和暦}}')
for i, h in enumerate(['品名', '数量', '単位', '単価', '金額']):
    c = ws.cell(row=12, column=i + 1, value=h); c.border = box; c.alignment = Alignment(horizontal='center'); c.font = Font(name=F, bold=True)
for r in range(13, 13 + LINES):
    for i in range(5):
        c = ws.cell(row=r, column=i + 1); c.border = box; c.font = Font(name=F)
    ws.cell(row=r, column=5, value=f'=IF(B{r}="","",B{r}*D{r})')
    for col in (2, 4, 5):
        ws.cell(row=r, column=col).number_format = '#,##0'
        ws.cell(row=r, column=col).alignment = Alignment(horizontal='right')
for col, field in zip('ABCD', ['品名', '数量', '単位', '単価']): ws[f'{col}13'] = '{{明細.%s}}' % field
last = 12 + LINES
for k, (label, formula) in enumerate([('小計', f'=SUM(E13:E{last})'), ('消費税（10%）', f'=ROUNDDOWN(E{last+1}*0.1,0)'), ('合計', f'=E{last+1}+E{last+2}')]):
    r = last + 1 + k
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    c = ws.cell(row=r, column=3, value=label); c.border = box; c.font = Font(name=F)
    c = ws.cell(row=r, column=5, value=formula); c.border = box; c.number_format = '#,##0'; c.font = Font(name=F, bold=(label == '合計'))
    c.alignment = Alignment(horizontal='right')
assert last + 3 == 25
put('A27', 'お振込先：○○銀行 ○○支店 普通 0000000 カ）サンプル')
put('A28', '備考：{{備考}}')
for i, line in enumerate(ISSUER): put(f'E{30 + i}', line, align='right')
ws.print_area = 'A1:E33'
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(os.path.join(OUT, '請求書.xlsx'))
print('ok')
