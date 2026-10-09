from pathlib import Path
import zipfile,re
from lxml import etree as E
root=Path(__file__).parent.parent;path=root/'ICS v9 - Format Lampiran Permenkes.docx';W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';n={'w':W}
with zipfile.ZipFile(path) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']);counter=0;count=0;changes=0;active=False
for p in d.xpath('//w:body/w:p',namespaces=n):
 nodes=p.xpath('./w:r/w:t|./w:hyperlink/w:r/w:t',namespaces=n);t=''.join(x.text or '' for x in nodes);sty=p.xpath('./w:pPr/w:pStyle/@w:val',namespaces=n)
 if t.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
 if not active:continue
 if sty and sty[0] in ['Heading1','Heading2','Heading3','Heading4']:counter=0
 replacements=[]
 for m in re.finditer(r'\bContoh\s+(\d+):',t,re.I):
  counter+=1;count+=1
  if int(m[1])!=counter:replacements.append((m.start(1),m.end(1),str(counter)));changes+=1
 for start,end,new in reversed(replacements):
  pos=0;insert=False
  for x in nodes:
   s=x.text or '';lo=max(start-pos,0);hi=min(end-pos,len(s))
   if hi>lo:x.text=s[:lo]+(new if not insert else '')+s[hi:];insert=True
   pos+=len(s)
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
 for key,val in parts.items():z.writestr(key,val)
print('All explicit prose example labels, including inline labels:',count,'; corrections:',changes)
