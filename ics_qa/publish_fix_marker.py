from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
p=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx'
with ZipFile(p) as z:parts={n:z.read(n) for n in z.namelist()}
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W}
d=E.fromstring(parts['word/document.xml']);count=0
for paragraph in d.xpath('//w:body/w:p',namespaces=N):
 ts=paragraph.xpath('./w:r/w:t',namespaces=N)
 if ts and ts[0].text=='\ufffd':ts[0].text='\u2022';count+=1
 if ts and ts[-1].text and ts[-1].text.endswith('\u2022') and not ''.join(t.text or '' for t in ts).startswith('\u2022'):
  ts[-1].text=ts[-1].text[:-1]
  r=E.Element('{'+W+'}r');E.SubElement(r,'{'+W+'}t').text='\u2022';E.SubElement(r,'{'+W+'}tab')
  paragraph.insert(1,r);count+=1
 for r in list(paragraph.findall('{'+W+'}r')):
  text=r.find('{'+W+'}t')
  if text is not None and text.text=='\u2022':
   paragraph.remove(r)
   paragraph.insert(1 if paragraph.find('{'+W+'}pPr') is not None else 0,r)
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(p,'w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
print('Markers corrected:',count)
