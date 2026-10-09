from zipfile import ZipFile
from lxml import etree as E
from pathlib import Path
with ZipFile('ICS v9 - Siap Publikasi.docx') as z:d=E.fromstring(z.read('word/document.xml'))
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W}
for p in d.xpath('//w:p',namespaces=N):
 t=''.join(p.xpath('.//w:t/text()',namespaces=N))
 if 'rujukan dari FKTP ke dokter spesialis bedah' in t:
  Path('ics_qa/shape-inspection.xml').write_bytes(E.tostring(p,pretty_print=True))
  print(t[:100],[(E.QName(e).localname,len(E.tostring(e))) for e in p if E.QName(e).localname=='r'])
ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','w':W}
els=d.xpath('//a:srgbClr[@val="FF0000"]',namespaces=ns)
print('Red shapes:',len(els))
for i,e in enumerate(els):
 p=next((a for a in e.iterancestors() if E.QName(a).localname=='p'),None)
 if p is not None:Path(f'ics_qa/red-shape-{i}.xml').write_bytes(E.tostring(p,pretty_print=True))
