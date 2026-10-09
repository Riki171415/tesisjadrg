from zipfile import ZipFile
from lxml import etree as E
from pathlib import Path
import re
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda x:'{'+W+'}'+x
d=E.fromstring(ZipFile(Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx').read('word/document.xml'))
count=0
for p in d.xpath('//w:body/w:p',namespaces=N):
 ind=p.find('w:pPr/w:ind',N)
 if ind is None or ind.get(q('hanging'))!='850':continue
 target=ind.get(q('left'));tabs=p.xpath('./w:pPr/w:tabs/w:tab/@w:pos',namespaces=N)
 assert tabs==[target],tabs
 before=''
 for el in p.xpath('./w:r/*',namespaces=N):
  if E.QName(el).localname=='tab':break
  if E.QName(el).localname=='t':before+=el.text or ''
 assert re.fullmatch(r'\s*(?:[A-Z]\d{2}(?:\d{2}/\d|\.(?:\d{1,4}|-))?|\d{2}\.\d{1,3})(?:[\u2020*])?',before),repr(before)
 count+=1
assert count>=559,count
print('Code hanging-indent and separator audit passed:',count,'rows.')
