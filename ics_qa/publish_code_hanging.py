from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import re,copy
path=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx';W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda x:'{'+W+'}'+x
with ZipFile(path) as z:parts={n:z.read(n) for n in z.namelist()}
d=E.fromstring(parts['word/document.xml']);body=d.find(q('body'));columns={};pending=[]
for element in body:
 if E.QName(element).localname=='p':pending.append(element);section=element.find('w:pPr/w:sectPr',N)
 elif E.QName(element).localname=='sectPr':section=element
 else:continue
 if section is not None:
  cols=section.find(q('cols'));multi=cols is not None and int(cols.get(q('num'),'1'))>1
  for para in pending:columns[para]=multi
  pending=[]
pattern=re.compile(r'^\s*(?:[A-Z]\d{2}(?:\d{2}/\d|\.(?:\d{1,4}|-))?|\d{2}\.\d{1,3})(?:[\u2020*])?')
count=0;multi_count=0
for p in body.findall(q('p')):
 text=''.join(p.xpath('./w:r/w:t/text()',namespaces=N));m=pattern.match(text)
 if not m or not text[m.end():].strip() or p.find('w:pPr/w:numPr',N) is not None:continue
 pp=p.find(q('pPr'))
 if pp is None:pp=E.Element(q('pPr'));p.insert(0,pp)
 ind=pp.find(q('ind'));base=int(ind.get(q('left'),'0')) if ind is not None else 0
 if ind is not None and ind.get(q('hanging'))=='850':base-=850
 if columns.get(p):base=0;multi_count+=1
 if ind is None:ind=E.SubElement(pp,q('ind'))
 ind.attrib.clear();ind.set(q('left'),str(base+850));ind.set(q('hanging'),'850');ind.set(q('right'),'0')
 tabs=pp.find(q('tabs'))
 if tabs is None:tabs=E.Element(q('tabs'));pp.insert(list(pp).index(ind),tabs)
 tabs.clear();tab=E.SubElement(tabs,q('tab'));tab.set(q('val'),'left');tab.set(q('pos'),str(base+850))
 jc=pp.find(q('jc'))
 if jc is None:jc=E.SubElement(pp,q('jc'))
 jc.set(q('val'),'left')
 for tab in p.xpath('./w:r/w:tab',namespaces=N):tab.getparent().remove(tab)
 offset=0
 for r in list(p.findall(q('r'))):
  for t in list(r.findall(q('t'))):
   length=len(t.text or '')
   if offset<m.end()<=offset+length:
    cut=m.end()-offset;tail=(t.text or '')[cut:];t.text=(t.text or '')[:cut]
    marker=E.Element(q('r'));rp=r.find(q('rPr'))
    if rp is not None:marker.append(copy.deepcopy(rp))
    E.SubElement(marker,q('tab'));p.insert(p.index(r)+1,marker)
    if tail:
     rest=E.Element(q('r'))
     if rp is not None:rest.append(copy.deepcopy(rp))
     tt=E.SubElement(rest,q('t'));tt.set('{http://www.w3.org/XML/1998/namespace}space','preserve');tt.text=tail;p.insert(p.index(marker)+1,rest)
    break
   offset+=length
  else:continue
  break
 count+=1
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(path,'w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
print({'code_rows_aligned':count,'two_column_code_rows':multi_count})
