import zipfile,copy,random,re
from pathlib import Path
from lxml import etree as E
root=Path(__file__).parent.parent;path=root/'ICS v9 - Format Lampiran Permenkes.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';n={'w':W}
def q(s):return '{'+W+'}'+s
with zipfile.ZipFile(path) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']);nr=E.fromstring(parts['word/numbering.xml'])
abmap={x.get(q('abstractNumId')):x for x in nr.findall(q('abstractNum'))}
nums={x.get(q('numId')):x for x in nr.findall(q('num'))}
base=next(x for x in abmap.values() if x.xpath('./w:lvl/w:numFmt/@w:val',namespaces=n)[:4]==['upperLetter','decimal','lowerLetter','decimal'])
aid=max(map(int,abmap))+1;nid=max(map(int,nums))+1;context=0;groups={};count=0;active=False
for p in d.xpath('//w:body/w:p',namespaces=n):
    text=''.join(p.xpath('./w:r/w:t/text()',namespaces=n)).strip();sty=p.xpath('./w:pPr/w:pStyle/@w:val',namespaces=n)
    if text.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
    if not active:continue
    if sty and sty[0] in ['Heading1','Heading2','Heading3','Heading4']:
        context+=1;continue
    if re.match(r'^Contoh \d+:',text):continue
    pr=p.find('w:pPr/w:numPr',n)
    if pr is None:continue
    lev=pr.find(q('ilvl'));num=pr.find(q('numId'))
    if lev is None or num is None:continue
    old=nums.get(num.get(q('val')))
    if old is None:continue
    ab=abmap[old.find(q('abstractNumId')).get(q('val'))]
    if ab.xpath('./w:lvl/w:numFmt/@w:val',namespaces=n)[:4]!=['upperLetter','decimal','lowerLetter','decimal']:continue
    level=lev.get(q('val'));key=(context,level)
    if key not in groups:
        newab=copy.deepcopy(base);newab.set(q('abstractNumId'),str(aid))
        for tag in ['nsid','tmpl']:
            x=newab.find(q(tag))
            if x is None:x=E.Element(q(tag));newab.insert(0,x)
            x.set(q('val'),f'{random.getrandbits(32):08X}')
        newab[:]=sorted(newab,key=lambda x:{'nsid':0,'multiLevelType':1,'tmpl':2,'lvl':3}.get(E.QName(x).localname,3))
        nr.append(newab);newnum=E.SubElement(nr,q('num'));newnum.set(q('numId'),str(nid));a=E.SubElement(newnum,q('abstractNumId'));a.set(q('val'),str(aid));ov=E.SubElement(newnum,q('lvlOverride'));ov.set(q('ilvl'),level);so=E.SubElement(ov,q('startOverride'));so.set(q('val'),'1');groups[key]=str(nid);aid+=1;nid+=1
    num.set(q('val'),groups[key]);count+=1
nr[:]=sorted(nr,key=lambda x:{'numPicBullet':0,'abstractNum':1,'num':2,'numIdMacAtCleanup':3}.get(E.QName(x).localname,4))
for key,val in [('word/document.xml',d),('word/numbering.xml',nr)]:parts[key]=E.tostring(val,xml_declaration=True,encoding='UTF-8',standalone=True)
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
    for key,val in parts.items():z.writestr(key,val)
print('Sequential ordinary lists:',count,'items in',len(groups),'parent/case groups.')
