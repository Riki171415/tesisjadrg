import zipfile,copy,random
from pathlib import Path
from lxml import etree as E
root=Path(__file__).parent.parent;path=root/'ICS v9 - Format Lampiran Permenkes.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';n={'w':W}
def q(s):return '{'+W+'}'+s
with zipfile.ZipFile(path) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']);nr=E.fromstring(parts['word/numbering.xml']);styles=E.fromstring(parts['word/styles.xml'])
head=d.xpath('//w:body/w:p[w:pPr/w:pStyle[@w:val="Heading1"]][w:pPr/w:numPr]',namespaces=n)[0]
mainid=head.find('w:pPr/w:numPr/w:numId',n).get(q('val'))
used={x.get(q('val')) for x in d.xpath('//w:body/w:p/w:pPr/w:numPr/w:numId',namespaces=n)}
abmap={x.get(q('abstractNumId')):x for x in nr.findall(q('abstractNum'))};aid=max(map(int,abmap))+1;cloned=0
for num in nr.findall(q('num')):
    nid=num.get(q('numId'));ab=abmap[num.find(q('abstractNumId')).get(q('val'))]
    if nid not in used or nid==mainid:continue
    fmts=ab.xpath('./w:lvl/w:numFmt/@w:val',namespaces=n)
    if fmts[:4]!=['upperLetter','decimal','lowerLetter','decimal']:continue
    clone=copy.deepcopy(ab);clone.set(q('abstractNumId'),str(aid))
    for tag in ['nsid','tmpl']:
        x=clone.find(q(tag))
        if x is None:x=E.Element(q(tag));clone.insert(0,x)
        x.set(q('val'),f'{random.getrandbits(32):08X}')
    clone[:]=sorted(clone,key=lambda x:{'nsid':0,'multiLevelType':1,'tmpl':2,'lvl':3}.get(E.QName(x).localname,3))
    nr.insert(len(nr.findall(q('abstractNum'))),clone);num.find(q('abstractNumId')).set(q('val'),str(aid));aid+=1;cloned+=1
    levels={x.get(q('val')) for x in d.xpath('//w:body/w:p[w:pPr/w:numPr/w:numId[@w:val="'+nid+'"]]/w:pPr/w:numPr/w:ilvl',namespaces=n)}
    for old in num.findall(q('lvlOverride')):num.remove(old)
    for lev in sorted(levels):
        override=E.SubElement(num,q('lvlOverride'));override.set(q('ilvl'),lev);start=E.SubElement(override,q('startOverride'));start.set(q('val'),'1')
# Remove draft formatting, including formatting inherited by automatic number marks.
for r in [d,nr,styles]:
    for x in r.xpath('//w:highlight|//w:strike|//w:dstrike|//w:rPr/w:shd',namespaces=n):x.getparent().remove(x)
for key,val in [('word/document.xml',d),('word/numbering.xml',nr),('word/styles.xml',styles)]:parts[key]=E.tostring(val,xml_declaration=True,encoding='UTF-8',standalone=True)
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
    for key,val in parts.items():z.writestr(key,val)
print('Independent body list definitions:',cloned,'; main heading list:',mainid)
