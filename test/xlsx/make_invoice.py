# Build 請求書ひな形.xlsx: a typical Japanese invoice sheet with {{placeholders}},
# fixed detail rows, number formats and SUM formulas.
# openpyxl stores no formula results; open and save the file once in Excel or
# LibreOffice (soffice --headless --convert-to xlsx) before using it, as a
# template saved by Excel always carries them.
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
out = sys.argv[1]
wb = Workbook(); ws = wb.active; ws.title = '請求書'
F = 'ＭＳ Ｐゴシック'
thin = Side(style='thin'); box = Border(left=thin, right=thin, top=thin, bottom=thin)
for col, w in zip('ABCDE', [6, 34, 9, 13, 15]): ws.column_dimensions[col].width = w
ws.merge_cells('A1:E1'); ws['A1'] = '請　求　書'
ws['A1'].font = Font(name=F, size=20, bold=True); ws['A1'].alignment = Alignment(horizontal='center')
ws['D3'] = '請求日'; ws['E3'] = '{{請求日:和暦}}'
ws['D4'] = '請求番号'; ws['E4'] = '{{請求番号}}'
ws.merge_cells('A3:C3'); ws['A3'] = '{{顧客名}}　御中'; ws['A3'].font = Font(name=F, size=14, underline='single')
ws['A5'] = '件名：{{件名}}'
ws['A7'] = 'ご請求金額'; ws['A7'].font = Font(name=F, size=12, bold=True)
ws.merge_cells('C7:D7'); ws['C7'] = '=E25'; ws['C7'].number_format = '"¥"#,##0"-"'
ws['C7'].font = Font(name=F, size=16, bold=True, underline='single')
hdr = ['No.', '品名', '数量', '単価', '金額']
for i, h in enumerate(hdr):
    c = ws.cell(row=9, column=i + 1, value=h); c.border = box; c.alignment = Alignment(horizontal='center')
    c.font = Font(name=F, bold=True)
for r in range(10, 22):
    for i in range(5):
        c = ws.cell(row=r, column=i + 1); c.border = box; c.font = Font(name=F)
    ws.cell(row=r, column=5, value=f'=IF(C{r}="","",C{r}*D{r})')
    for col in (3, 4, 5): ws.cell(row=r, column=col).number_format = '#,##0'
ws['A10'] = '{{明細.No}}'; ws['B10'] = '{{明細.品名}}'; ws['C10'] = '{{明細.数量}}'; ws['D10'] = '{{明細.単価}}'
for r, (label, formula) in enumerate([('小計', '=SUM(E10:E21)'), ('消費税（10%）', '=ROUNDDOWN(E23*0.1,0)'), ('合計', '=E23+E24')], start=23):
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    ws.cell(row=r, column=3, value=label).border = box
    c = ws.cell(row=r, column=5, value=formula); c.border = box; c.number_format = '#,##0'
ws['A27'] = '備考：{{備考}}'
ws['A29'] = '株式会社サンプル　〒100-0001 東京都千代田区千代田1-1'
for row in ws.iter_rows():
    for c in row:
        if c.font.name != F: c.font = Font(name=F, size=c.font.size, bold=c.font.b, underline=c.font.u)
ws.print_area = 'A1:E29'
ws.page_setup.paperSize = ws.PAPERSIZE_A4
wb.save(out)
print(out)
