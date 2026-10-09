from zipfile import ZipFile,ZIP_DEFLATED
from pathlib import Path
from lxml import etree as E
import copy
p=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx';W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';n={'w':W};q=lambda s:'{'+W+'}'+s
def setel(e,t,**a):
 x=e.find(q(t))
 if x is None:x=E.SubElement(e,q(t))
 x.attrib.clear()
 for k,v in a.items():x.set(q(k),str(v))
 return x
with ZipFile(p) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']);body=d.find(q('body'));changed=0
for old in list(body):
 if old.tag!=q('p'):continue
 text=''.join(old.xpath('./w:r/w:t/text()',namespaces=n));mode=None;remaining=0
 if text.startswith('Abscess Of Lung (J85)') and len(text)>100:mode='break'
 elif text.startswith('Perhatian khusus pada kode eksklusi dalam kategori O99.-.') and 'Sebagai contoh' in text:mode='offset';remaining=text.index('Sebagai contoh')
 if mode is None:continue
 head=E.Element(q('p'),dict(old.attrib));pp=old.find(q('pPr'));head.append(copy.deepcopy(pp))
 narrative=E.Element(q('p'));np=E.SubElement(narrative,q('pPr'))
 setel(np,'pStyle',val='Normal');setel(np,'spacing',before=0,after=120,line=276,lineRule='auto');setel(np,'ind',left=0,right=0,firstLine=567);setel(np,'jc',val='both');setel(np,'widowControl',val=1)
 after=False
 for node in old:
  if node.tag==q('pPr'):continue
  if node.tag!=q('r'):
   (narrative if after else head).append(copy.deepcopy(node));continue
  pre=E.Element(q('r'),dict(node.attrib));post=E.Element(q('r'),dict(node.attrib));rp=node.find(q('rPr'))
  if rp is not None:pre.append(copy.deepcopy(rp));post.append(copy.deepcopy(rp))
  for c in node:
   if c.tag==q('rPr'):continue
   if mode=='break' and c.tag==q('br') and not after:after=True;continue
   if mode=='offset' and c.tag==q('t') and not after:
    value=c.text or '';take=min(remaining,len(value))
    if take:
     t=copy.deepcopy(c);t.text=value[:take];t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');pre.append(t)
    remaining-=take
    if remaining==0:after=True
    if take<len(value):t=copy.deepcopy(c);t.text=value[take:];t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');post.append(t)
   else:(post if after else pre).append(copy.deepcopy(c))
  if any(c.tag!=q('rPr') for c in pre):head.append(pre)
  if any(c.tag!=q('rPr') for c in post):narrative.append(post)
 assert after,'Heading boundary not found'
 for r in narrative.findall(q('r')):
  rp=r.find(q('rPr'))
  if rp is None:rp=E.SubElement(r,q('rPr'))
  for tag in ['b','bCs']:
   for x in rp.findall(q(tag)):rp.remove(x)
  setel(rp,'rFonts',ascii='Times New Roman',hAnsi='Times New Roman',eastAsia='Times New Roman',cs='Times New Roman');setel(rp,'sz',val=24);setel(rp,'szCs',val=24)
 index=list(body).index(old);body.remove(old);body.insert(index,head);body.insert(index+1,narrative);changed+=1
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(p,'w',ZIP_DEFLATED) as z:
 for name,data in parts.items():z.writestr(name,data)
print('Separated',changed,'explanations from their headings; text stream preserved.')
