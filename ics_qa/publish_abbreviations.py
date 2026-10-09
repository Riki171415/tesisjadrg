from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy
p=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';n={'w':W};q=lambda s:'{'+W+'}'+s
with ZipFile(p) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']);tbl=d.xpath('//w:tbl',namespaces=n)[0];first=tbl.find(q('tr'))
if ''.join(first.xpath('./w:tc[1]//w:t/text()',namespaces=n))!='Singkatan':
 header=copy.deepcopy(first)
 for cell,label in zip(header.findall(q('tc')),['Singkatan','Kepanjangan']):
  ts=cell.xpath('.//w:t',namespaces=n)
  for t in ts:t.text=''
  ts[0].text=label
  for rp in cell.xpath('.//w:rPr',namespaces=n):
   for x in rp.findall(q('i'))+rp.findall(q('iCs')):rp.remove(x)
   b=rp.find(q('b'))
   if b is None:b=E.SubElement(rp,q('b'))
   b.set(q('val'),'1')
 for x in first.xpath('./w:trPr/w:tblHeader',namespaces=n):x.getparent().remove(x)
 for rp in first.xpath('.//w:rPr',namespaces=n):
  for x in rp.findall(q('b'))+rp.findall(q('bCs')):rp.remove(x)
 tbl.insert(list(tbl).index(first),header)
 parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
 with ZipFile(p,'w',ZIP_DEFLATED) as z:
  for name,data in parts.items():z.writestr(name,data)
print('Abbreviations table has a repeating column header; original entries preserved.')
