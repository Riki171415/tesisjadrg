from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
p=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx';W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda x:'{'+W+'}'+x
with ZipFile(p) as z:parts={n:z.read(n) for n in z.namelist()}
d=E.fromstring(parts['word/document.xml'])
for para in d.xpath('//w:body/w:p',namespaces=N):
 texts=para.xpath('.//w:t/text()',namespaces=N)
 if texts and set(texts)=={'Merujuk pada sesi','Kuantitas berdasarkan lokasi'}:
  para.clear();pp=E.SubElement(para,q('pPr'));sp=E.SubElement(pp,q('spacing'));sp.set(q('after'),'120');sp.set(q('line'),'240');sp.set(q('lineRule'),'auto');ind=E.SubElement(pp,q('ind'));ind.set(q('firstLine'),'0');jc=E.SubElement(pp,q('jc'));jc.set(q('val'),'center')
  r=E.SubElement(para,q('r'));rp=E.SubElement(r,q('rPr'));font=E.SubElement(rp,q('rFonts'));font.set(q('ascii'),'Times New Roman');font.set(q('hAnsi'),'Times New Roman');sz=E.SubElement(rp,q('sz'));sz.set(q('val'),'20')
  E.SubElement(r,q('t')).text='Setting: Merujuk pada sesi; Multiplicity: Kuantitas berdasarkan lokasi.'
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(p,'w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
print('Figure annotations converted to a readable legend.')
