from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
p=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx'
with ZipFile(p) as z:parts={n:z.read(n) for n in z.namelist()}
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda x:'{'+W+'}'+x
d=E.fromstring(parts['word/document.xml']);active=False;count=0
for para in d.xpath('//w:body/w:p',namespaces=N):
 text=''.join(para.xpath('./w:r/w:t/text()',namespaces=N)).strip()
 if 'Tahapan Memilih Kode Tindakan Utama' in text:active=True;continue
 if active and text.startswith('Catatan:'):active=False
 if active:
  il=para.find('w:pPr/w:numPr/w:ilvl',N)
  if il is not None:
   il.set(q('val'),'4');ind=para.find('w:pPr/w:ind',N);ind.set(q('left'),'2125');ind.set(q('hanging'),'425')
   for tab in para.xpath('./w:pPr/w:tabs/w:tab',namespaces=N):tab.set(q('pos'),'2125')
   count+=1
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(p,'w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
print('Nested substeps corrected:',count)
