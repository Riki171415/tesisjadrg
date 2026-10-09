from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
path=Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx'
with ZipFile(path) as z:parts={n:z.read(n) for n in z.namelist()}
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda x:'{'+W+'}'+x
d=E.fromstring(parts['word/document.xml']);active=False;left=425;count=0
for p in d.xpath('//w:body/w:p',namespaces=N):
 t=''.join(p.xpath('./w:r/w:t/text()',namespaces=N)).strip()
 if t.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
 if not active:continue
 pp=p.find(q('pPr'))
 if pp is None:pp=E.Element(q('pPr'));p.insert(0,pp)
 elif p.index(pp)!=0:p.remove(pp);p.insert(0,pp)
 sid=p.xpath('./w:pPr/w:pStyle/@w:val',namespaces=N)
 if sid and sid[0] in ['Heading1','Heading2','Heading3','Heading4']:left=int(sid[0][-1])*425;continue
 np=pp.find(q('numPr'));ind=pp.find(q('ind'))
 if np is not None or t.startswith('\u2022'):
  left=int(ind.get(q('left'),'425')) if ind is not None else 425;continue
 if not t or t.startswith(('Tabel ','Gambar ')) or p.xpath('.//w:drawing|.//w:pict|.//w:object',namespaces=N):continue
 if ind is None:ind=E.SubElement(pp,q('ind'))
 ind.attrib.clear();ind.set(q('left'),str(left));ind.set(q('right'),'0');ind.set(q('firstLine'),'0' if len(t)<140 and t.endswith(':') else '567')
 count+=1
 # Keep paragraph properties in Word's schema sequence.
 order=['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange']
 children=list(pp)
 for c in children:pp.remove(c)
 for c in sorted(children,key=lambda c:order.index(E.QName(c).localname) if E.QName(c).localname in order else 99):pp.append(c)
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(path,'w',ZIP_DEFLATED) as z:
 for n,v in parts.items():z.writestr(n,v)
print('Stable narrative indentation:',count)
