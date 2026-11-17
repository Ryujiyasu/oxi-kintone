import sys
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
d = Document()
sec = d.sections[0]; sec.page_width, sec.page_height = Mm(210), Mm(297)
for m in ('left_margin', 'right_margin'): setattr(sec, m, Mm(20))
st = d.styles['Normal']; st.font.name = 'MS Mincho'; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'ＭＳ 明朝')
def para(text='', size=None, bold=False, align=None, font=None):
    p = d.add_paragraph()
    if text:
        r = p.add_run(text); r.bold = bold
        if size: r.font.size = Pt(size)
        if font: r.font.name = font; r._element.rPr.rFonts.set(qn('w:eastAsia'), font)
    if align is not None: p.alignment = align
    return p
para('御 見 積 書', 20, True, WD_ALIGN_PARAGRAPH.CENTER, 'ＭＳ ゴシック')
para('見積日：{{見積日}}', align=WD_ALIGN_PARAGRAPH.RIGHT)
p = para(); r = p.add_run('{{顧客'); r.font.size = Pt(14); r = p.add_run('名}}'); r.font.size = Pt(14); r.bold = True
r = p.add_run('　御中'); r.font.size = Pt(14)
para('件名：{{件名}}')
para('下記の通りお見積り申し上げます。')
p = para(); r = p.add_run('御見積金額　{{合計}} 円（税込）'); r.bold = True; r.font.size = Pt(14); r.underline = True
t = d.add_table(rows=3, cols=4); t.style = 'Table Grid'
for i, h in enumerate(['品名', '数量', '単価', '金額']): t.cell(0, i).text = h
for i, h in enumerate(['{{明細.品名}}', '{{明細.数量}}', '{{明細.単価}}', '{{明細.金額}}']): t.cell(1, i).text = h
t.cell(2, 2).text = '合計'; t.cell(2, 3).text = '{{合計}}'
for row in t.rows:
    for i, c in enumerate(row.cells):
        if i > 0:
            for pp in c.paragraphs: pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
para()
para('備考：{{備考}}')
para('株式会社サンプル　〒100-0001 東京都千代田区千代田1-1', align=WD_ALIGN_PARAGRAPH.RIGHT)
d.save(sys.argv[1])
