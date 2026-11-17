# python -I dump.py <docx>: list body paragraphs with their table/row/cell position
import sys, zipfile, xml.etree.ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
root=ET.fromstring(zipfile.ZipFile(sys.argv[1]).read('word/document.xml'))
def text(p): return ''.join(t.text or '' for t in p.iter(W+'t'))
def walk(el, path):
    ti=0
    for ch in el:
        if ch.tag==W+'p': print(path, repr(text(ch))[:70], 'txbx' if ch.find('.//'+W+'txbxContent') is not None else '')
        elif ch.tag==W+'tbl':
            for ri,tr in enumerate(ch.findall(W+'tr')):
                for ci,tc in enumerate(tr.findall(W+'tc')):
                    walk(tc, f'{path}T{ti}R{ri}C{ci} ')
            ti+=1
        elif ch.tag in (W+'body', W+'sdt', W+'sdtContent'): walk(ch, path)
walk(root.find(W+'body'), '')
