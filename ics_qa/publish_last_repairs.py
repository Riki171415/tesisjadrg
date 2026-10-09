from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy
p=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx'
with ZipFile(p) as z:parts={n:z.read(n) for n in z.namelist()}
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';N={'w':W,'wp':WP};q=lambda x:'{'+W+'}'+x
d=E.fromstring(parts['word/document.xml']);body=d.find(q('body'));branch=False;fixed=0
for para in list(body.findall(q('p'))):
 text=''.join(para.xpath('./w:r/w:t/text()',namespaces=N)).strip()
 for color in para.xpath('./w:r/w:rPr/w:color',namespaces=N):color.set(q('val'),'000000')
 if text.startswith('Tahapan Memilih Kode Tindakan Utama'):branch=True;continue
 if branch and text.startswith('Catatan:'):branch=False
 if branch:
  il=para.find('w:pPr/w:numPr/w:ilvl',N)
  if il is not None:
   il.set(q('val'),'4');ind=para.find('w:pPr/w:ind',N);ind.set(q('left'),'2125');ind.set(q('hanging'),'425');fixed+=1
 if text.startswith('Gambar C.1.'):
  diagram=E.Element(q('p'));pp=E.SubElement(diagram,q('pPr'));jc=E.SubElement(pp,q('jc'));jc.set(q('val'),'center')
  for r in list(para.findall(q('r'))):
   anchors=r.xpath('.//wp:anchor',namespaces=N)
   if not anchors:continue
   for anchor in anchors:
    anchor.tag='{'+WP+'}inline';anchor.attrib.clear()
    for el in list(anchor):
     if E.QName(el).localname in ['simplePos','positionH','positionV','wrapNone','wrapSquare','wrapTight','wrapThrough','wrapTopAndBottom']:anchor.remove(el)
   para.remove(r);diagram.append(r)
  if len(diagram)>1:body.insert(body.index(para)+1,diagram)
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(p,'w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
print('Nested steps corrected:',fixed,'; diagram annotations placed inline.')
