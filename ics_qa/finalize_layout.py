from pathlib import Path
import zipfile,re,copy
from lxml import etree as E
root=Path(__file__).parent.parent;path=root/'ICS v9 - Format Lampiran Permenkes.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';n={'w':W}
def q(s):return '{'+W+'}'+s
with zipfile.ZipFile(path) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']);body=d.find(q('body'))
for p in list(body.findall(q('p'))):
 t=''.join(p.xpath('./w:r/w:t/text()',namespaces=n)).strip()
 if t=='Pencatatan Sesi Tindakan (Setting) dan Jumlah Lokasi (Multiplicity)':
  p.find('w:pPr/w:pStyle',n).set(q('val'),'Heading3');p.find('w:pPr/w:numPr/w:ilvl',n).set(q('val'),'2')
  pr=p.find('w:pPr',n);ol=pr.find(q('outlineLvl'))
  if ol is None:ol=E.SubElement(pr,q('outlineLvl'))
  ol.set(q('val'),'2')
 if re.match(r'^[12]\) jika ',t):
  remaining=3
  for x in p.xpath('./w:r/w:t',namespaces=n):
   s=x.text or '';take=min(remaining,len(s));x.text=s[take:];remaining-=take
   if remaining==0:break
 if re.match(r'^(Tabel|Gambar) [ABC]\.\d+\.',t):
  drawings=p.xpath('./w:r/w:drawing|./w:r/w:pict',namespaces=n)
  if drawings:
   picture=E.Element(q('p'));pr=E.SubElement(picture,q('pPr'));E.SubElement(pr,q('jc')).set(q('val'),'center');E.SubElement(pr,q('keepNext')).set(q('val'),'1')
   for x in drawings:
    x.getparent().remove(x);r=E.SubElement(picture,q('r'));r.append(x)
   body.insert(body.index(p),picture)
 # Clear draft highlighting on titles without changing color-coded clinical text.
 if p.xpath('./w:pPr/w:pStyle[starts-with(@w:val,"Heading")]',namespaces=n):
  for x in p.xpath('./w:r/w:rPr/w:highlight|./w:r/w:rPr/w:shd',namespaces=n):x.getparent().remove(x)
# Place the contents title on its own page rather than behind the cover image.
first=body.find(q('p'))
if ''.join(first.xpath('./w:r/w:t/text()',namespaces=n)).strip()=='DAFTAR ISI':
 for x in first.xpath('./w:r/w:t',namespaces=n):
  if 'DAFTAR ISI' in (x.text or ''):x.text=(x.text or '').replace('DAFTAR ISI','')
 s=first.find('w:pPr/w:pStyle',n)
 if s is not None:s.set(q('val'),'Normal')
 title=E.Element(q('p'));pr=E.SubElement(title,q('pPr'))
 for key,val in [('keepNext','1'),('pageBreakBefore','1'),('outlineLvl','9')]:E.SubElement(pr,q(key)).set(q('val'),val)
 r=E.SubElement(title,q('r'));rp=E.SubElement(r,q('rPr'));E.SubElement(rp,q('b'));fonts=E.SubElement(rp,q('rFonts'));fonts.set(q('ascii'),'Times New Roman');fonts.set(q('hAnsi'),'Times New Roman');E.SubElement(rp,q('sz')).set(q('val'),'24');E.SubElement(r,q('t')).text='DAFTAR ISI'
 firstsdt=body.find(q('sdt'));body.insert(body.index(firstsdt),title)
 for x in firstsdt.xpath('.//w:p[1]/w:pPr/w:pageBreakBefore',namespaces=n):x.getparent().remove(x)
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
 for key,val in parts.items():z.writestr(key,val)
print('Contents title, figure placement, manual list prefixes, and one hierarchy correction applied.')
