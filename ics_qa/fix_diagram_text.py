from pathlib import Path
import zipfile,re
from lxml import etree as E
root=Path(__file__).parent.parent;path=root/'ICS v9 - Format Lampiran Permenkes.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';n={'w':W,'wps':'http://schemas.microsoft.com/office/word/2010/wordprocessingShape','v':'urn:schemas-microsoft-com:vml'}
def q(s):return '{'+W+'}'+s
def prop(parent,name,value):
 x=parent.find(q(name))
 if x is None:x=E.SubElement(parent,q(name))
 x.set(q('val'),str(value));return x
with zipfile.ZipFile(path) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']);count=0
for p in d.xpath('//w:body/w:p[.//w:txbxContent]',namespaces=n):
 alltext=''.join(p.xpath('.//w:t/text()',namespaces=n))
 if not alltext.startswith('2401210'):continue
 for box in p.xpath('.//w:txbxContent',namespaces=n):
  t=''.join(box.xpath('.//w:t/text()',namespaces=n)).strip()
  size=28 if t.isdigit() else 18
  for paragraph in box.findall(q('p')):
   pr=paragraph.find(q('pPr'))
   if pr is None:pr=E.Element(q('pPr'));paragraph.insert(0,pr)
   spacing=pr.find(q('spacing'))
   if spacing is None:spacing=E.SubElement(pr,q('spacing'))
   spacing.set(q('before'),'0');spacing.set(q('after'),'0');spacing.set(q('line'),'240');spacing.set(q('lineRule'),'auto');prop(pr,'jc','center')
   for run in paragraph.findall(q('r')):
    rp=run.find(q('rPr'))
    if rp is None:rp=E.Element(q('rPr'));run.insert(0,rp)
    prop(rp,'sz',size);prop(rp,'szCs',size);prop(rp,'color','000000')
    fonts=rp.find(q('rFonts'))
    if fonts is None:fonts=E.Element(q('rFonts'));rp.insert(0,fonts)
    for attr in ['ascii','hAnsi','cs','eastAsia']:fonts.set(q(attr),'Times New Roman')
    order={s:i for i,s in enumerate('rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang'.split())};rp[:]=sorted(rp,key=lambda x:order.get(E.QName(x).localname,999))
   order={s:i for i,s in enumerate('pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs spacing ind contextualSpacing jc textDirection textAlignment outlineLvl rPr sectPr'.split())};pr[:]=sorted(pr,key=lambda x:order.get(E.QName(x).localname,999))
  count+=1
 for bp in p.xpath('.//wps:bodyPr',namespaces=n):
  for attr in ['lIns','tIns','rIns','bIns']:bp.set(attr,'0')
 for tb in p.xpath('.//v:textbox',namespaces=n):tb.set('inset','0,0,0,0')
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
 for key,val in parts.items():z.writestr(key,val)
print('Diagram text boxes fitted without changing their text:',count)
