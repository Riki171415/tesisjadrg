from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
path=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx';W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda x:'{'+W+'}'+x
with ZipFile(path) as z:parts={n:z.read(n) for n in z.namelist()}
d=E.fromstring(parts['word/document.xml']);nr=E.fromstring(parts['word/numbering.xml']);removed=0;active=False
for p in d.xpath('//w:body/w:p',namespaces=N):
 text=''.join(p.xpath('./w:r/w:t/text()',namespaces=N)).strip()
 if text.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
 if not active:continue
 pp=p.find(q('pPr'))
 if pp is None:continue
 np=pp.find(q('numPr'));ind=pp.find(q('ind'))
 if np is not None:
  found_text=False
  for r in p.findall(q('r')):
   for el in list(r):
    if E.QName(el).localname=='t' and (el.text or '').strip():found_text=True
    if not found_text and E.QName(el).localname=='tab':r.remove(el);removed+=1
 if ind is not None:
  left=int(ind.get(q('left'),'0'))
  if left>=425:ind.set(q('left'),str(left-85))
  if ind.get(q('hanging')) is not None:ind.set(q('hanging'),'340')
 for tab in pp.xpath('./w:tabs/w:tab',namespaces=N):
  pos=int(tab.get(q('pos'),'0'))
  if pos>=425:tab.set(q('pos'),str(pos-85))
for ab in nr.findall(q('abstractNum')):
 formats=ab.xpath('./w:lvl/w:numFmt/@w:val',namespaces=N)
 if formats[:5]!=['upperLetter','decimal','lowerLetter','decimal','lowerLetter']:continue
 for lv in ab.findall(q('lvl'))[:5]:
  level=int(lv.get(q('ilvl')));pp=lv.find(q('pPr'))
  if pp is None:pp=E.SubElement(lv,q('pPr'))
  ind=pp.find(q('ind'))
  if ind is None:ind=E.SubElement(pp,q('ind'))
  ind.attrib.clear();ind.set(q('left'),str(level*425+340));ind.set(q('hanging'),'340')
  tabs=pp.find(q('tabs'))
  if tabs is None:tabs=E.Element(q('tabs'));pp.insert(0,tabs)
  tabs.clear();tab=E.SubElement(tabs,q('tab'));tab.set(q('val'),'num');tab.set(q('pos'),str(level*425+340))
for n,x in [('word/document.xml',d),('word/numbering.xml',nr)]:parts[n]=E.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(path,'w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
print('Extra leading tabs removed:',removed,'; numbering gap standardized to 0.6 cm.')
