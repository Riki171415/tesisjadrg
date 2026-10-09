from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import re
root=Path(__file__).parent.parent;p=root/'ICS v9 - Siap Publikasi.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';n={'w':W,'v':'urn:schemas-microsoft-com:vml'};q=lambda x:'{'+W+'}'+x
with ZipFile(p) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
def child(e,t):
 x=e.find(q(t))
 if x is None:x=E.SubElement(e,q(t))
 return x
def setel(e,t,**a):
 x=child(e,t);x.attrib.clear()
 for k,v in a.items():x.set(q(k),str(v))
 return x
d=E.fromstring(parts['word/document.xml']);s=E.fromstring(parts['word/styles.xml']);settings=E.fromstring(parts['word/settings.xml'])
setel(child(settings,'compat'),'doNotExpandShiftReturn')
for st in s.findall(q('style')):
 sid=st.get(q('styleId'),'')
 if re.match(r'^TOC[1-9]$',sid):
  lev=int(sid[-1])-1;pp=child(st,'pPr');ind=lev*425 if lev<4 else 0
  setel(pp,'ind',left=ind,right=0,firstLine=0)
  tabs=child(pp,'tabs');tabs.clear()
  if lev<4:
   t=E.SubElement(tabs,q('tab'));t.set(q('val'),'left');t.set(q('pos'),str(ind+425))
  t=E.SubElement(tabs,q('tab'));t.set(q('val'),'right');t.set(q('leader'),'dot');t.set(q('pos'),'9026')
for pnode in d.xpath('//w:p[not(ancestor::w:txbxContent)]',namespaces=n):
 text=''.join(pnode.xpath('./w:r/w:t/text()',namespaces=n)).strip()
 if text in ['DAFTAR ISI','DAFTAR GAMBAR','DAFTAR TABEL','DAFTAR SINGKATAN','GLOSARIUM']:
  pp=child(pnode,'pPr');setel(pp,'ind',left=0,right=0,firstLine=0);setel(pp,'spacing',before=0,after=240,line=276,lineRule='auto')
  for r in pnode.findall(q('r')):
   rp=child(r,'rPr');setel(rp,'b',val=1);setel(rp,'sz',val=28);setel(rp,'szCs',val=28)
for tbl in d.xpath('//w:tbl',namespaces=n):
 rows=tbl.findall(q('tr'))
 for ri,row in enumerate(rows):
  for ci,cell in enumerate(row.findall(q('tc'))):
   cm=child(child(cell,'tcPr'),'tcMar')
   for side in ['top','left','bottom','right']:setel(cm,side,w=80,type='dxa')
   for pn in cell.findall(q('p')):
    text=''.join(pn.xpath('./w:r/w:t/text()',namespaces=n));rs=[r for r in pn.findall(q('r')) if r.find(q('t')) is not None]
    if len(text)>150 and rs and all(r.find('w:rPr/w:b',n) is not None for r in rs):
     for r in rs:
      rp=child(r,'rPr')
      for x in rp.findall(q('b'))+rp.findall(q('bCs')):rp.remove(x)
    if ri==0 and ci==0 and re.fullmatch(r'Contoh \d+',text.strip()):
     for r in rs:setel(child(r,'rPr'),'b',val=1)
for name in list(parts):
 if re.match(r'word/header\d+\.xml$',name):
  hdr=E.fromstring(parts[name])
  for shape in hdr.xpath('//v:shape[v:textpath[contains(@string,"DRAFT")]]',namespaces=n):shape.getparent().remove(shape)
  parts[name]=E.tostring(hdr,xml_declaration=True,encoding='UTF-8',standalone=True)
for name,el in [('word/document.xml',d),('word/styles.xml',s),('word/settings.xml',settings)]:parts[name]=E.tostring(el,xml_declaration=True,encoding='UTF-8',standalone=True)
clean=root/'ics_qa/cover-clean.png'
if clean.exists():parts['word/media/image1.png']=clean.read_bytes()
with ZipFile(p,'w',ZIP_DEFLATED) as z:
 for name,data in parts.items():z.writestr(name,data)
print('Refined TOC, titles, table padding, and publishing marks.')
