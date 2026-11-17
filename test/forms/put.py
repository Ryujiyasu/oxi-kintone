# python -I put.py <in.docx> <out.docx> <edits.json>
# Types placeholders into paragraphs the way a user editing in Word would.
# edits: {"<path>#<k>": ["append"|"replace", text]}, path as printed by dump.py
import sys, json, zipfile, copy, xml.etree.ElementTree as ET
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; w='{%s}'%W
src, dst, spec = sys.argv[1:4]
edits = json.load(open(spec, encoding='utf-8'))
z = zipfile.ZipFile(src)
raw = z.read('word/document.xml')
# keep every namespace prefix as written
for ev, (pfx, uri) in ET.iterparse(__import__('io').BytesIO(raw), events=['start-ns']):
    ET.register_namespace(pfx, uri)
root = ET.fromstring(raw)
paras = []
def walk(el, path):
    ti = 0
    for ch in el:
        if ch.tag == w+'p': paras.append((path, ch))
        elif ch.tag == w+'tbl':
            for ri, tr in enumerate(ch.findall(w+'tr')):
                for ci, tc in enumerate(tr.findall(w+'tc')):
                    walk(tc, f'{path}T{ti}R{ri}C{ci} ')
            ti += 1
        elif ch.tag in (w+'body', w+'sdt', w+'sdtContent'): walk(ch, path)
walk(root.find(w+'body'), '')
seen = {}
keyed = {}
for path, p in paras:
    k = seen.get(path, 0); seen[path] = k + 1
    keyed[f'{path.strip()}#{k}'] = p
done = 0
for key, (mode, text) in edits.items():
    p = keyed[key]
    runs = p.findall(w+'r')
    if mode == 'replace' and runs:
        ts = [t for r in runs for t in r.findall(w+'t')]
        ts[0].text = text
        for t in ts[1:]: t.text = ''
    else:
        r = ET.SubElement(p, w+'r')
        base = runs[-1].find(w+'rPr') if runs else p.find(f'{w}pPr/{w}rPr')
        if base is not None: r.append(copy.deepcopy(base))
        t = ET.SubElement(r, w+'t'); t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    done += 1
out = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for item in z.infolist():
    data = ET.tostring(root, xml_declaration=True, encoding='UTF-8') if item.filename == 'word/document.xml' else z.read(item)
    out.writestr(item, data)
out.close()
print('edits', done)
