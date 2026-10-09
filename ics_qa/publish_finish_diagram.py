from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy
path=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx'
with ZipFile(path) as z:parts={n:z.read(n) for n in z.namelist()}
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda x:'{'+W+'}'+x
d=E.fromstring(parts['word/document.xml']);nr=E.fromstring(parts['word/numbering.xml'])
for tx in d.xpath('//w:txbxContent',namespaces=N):
 if ''.join(tx.xpath('.//w:t/text()',namespaces=N)) not in ['Merujuk pada sesi','Kuantitas berdasarkan lokasi']:continue
 for p in tx.findall(q('p')):
  pp=p.find(q('pPr'))
  for el in list(pp):
   if E.QName(el).localname in ['ind','textDirection','spacing']:pp.remove(el)
  sp=E.SubElement(pp,q('spacing'));sp.set(q('after'),'0');sp.set(q('line'),'160');sp.set(q('lineRule'),'auto')
  ind=E.SubElement(pp,q('ind'));ind.set(q('left'),'0');ind.set(q('right'),'0');ind.set(q('firstLine'),'0')
  direction=E.SubElement(pp,q('textDirection'));direction.set(q('val'),'lrTb')
  for r in p.findall(q('r')):
   rp=r.find(q('rPr'))
   if rp is None:rp=E.Element(q('rPr'));r.insert(0,rp)
   sz=rp.find(q('sz'))
   if sz is None:sz=E.SubElement(rp,q('sz'))
   sz.set(q('val'),'16')
numid=max(int(n.get(q('numId'))) for n in nr.findall(q('num')))+1
cases=[]
for p in d.xpath('//w:body/w:p',namespaces=N):
 t=''.join(p.xpath('./w:r/w:t/text()',namespaces=N)).strip()
 if t.startswith(('Pasien dengan kanker usus (penyakit utama)','Pasien dengan abses hati (penyakit utama)')):cases.append(p)
if cases:
 old=cases[0].find('w:pPr/w:numPr/w:numId',N).get(q('val'))
 source=next(n for n in nr.findall(q('num')) if n.get(q('numId'))==old)
 new=copy.deepcopy(source);new.set(q('numId'),str(numid))
 for ov in list(new.findall(q('lvlOverride'))):new.remove(ov)
 ov=E.SubElement(new,q('lvlOverride'));ov.set(q('ilvl'),'4');start=E.SubElement(ov,q('startOverride'));start.set(q('val'),'1');nr.append(new)
 for p in cases:p.find('w:pPr/w:numPr/w:numId',N).set(q('val'),str(numid))
for n,x in [('word/document.xml',d),('word/numbering.xml',nr)]:parts[n]=E.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(path,'w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
print('Diagram labels made readable; separate example list restarted:',len(cases))
