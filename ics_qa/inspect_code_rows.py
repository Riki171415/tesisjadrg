from zipfile import ZipFile
from lxml import etree as E
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W}
d=E.fromstring(ZipFile('ICS v9 - Siap Publikasi.docx').read('word/document.xml'))
for p in d.xpath('//w:p',namespaces=N):
 t=''.join(p.xpath('./w:r/w:t/text()',namespaces=N)).strip()
 if t.startswith(('93.62','74.91','O80.0','O83.0')):
  print(repr(t[:180]),'parent',E.QName(p.getparent()).localname)
  print(E.tostring(p.find('w:pPr',N)).decode()[-900:] if p.find('w:pPr',N) is not None else 'no pPr')
  print([(E.QName(e).localname,e.text) for r in p.findall('{'+W+'}r') for e in r if E.QName(e).localname in ['t','tab','br']][:12])
