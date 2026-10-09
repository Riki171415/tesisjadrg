from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,re,json
root=Path(__file__).parent.parent;target=root/'ICS v9 - Siap Publikasi.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda x:'{'+W+'}'+x
def child(e,t):
 x=e.find(q(t))
 if x is None:x=E.SubElement(e,q(t))
 return x
def setel(e,t,**attrs):
 x=child(e,t);x.attrib.clear()
 for k,v in attrs.items():x.set(q(k),str(v))
 return x
def text(p):return ''.join(p.xpath('./w:r/w:t/text()',namespaces=N)).strip()
def normalize(t):return re.sub(r'\s+',' ',t).strip()
with ZipFile(root/'ICS v9.docx') as z:orig=E.fromstring(z.read('word/document.xml'));on=E.fromstring(z.read('word/numbering.xml'))
oa={x.get(q('abstractNumId')):x for x in on.findall(q('abstractNum'))};om={x.get(q('numId')):oa[x.find(q('abstractNumId')).get(q('val'))] for x in on.findall(q('num'))}
source={}
for p in orig.xpath('//w:body/w:p',namespaces=N):
 t=normalize(text(p))
 if not t:continue
 np=p.find('w:pPr/w:numPr',N);ind=p.find('w:pPr/w:ind',N);fmt=None;left=None
 if np is not None:
  ni=np.find(q('numId'));il=np.find(q('ilvl'));il=il.get(q('val')) if il is not None else '0'
  if ni is not None and ni.get(q('val')) in om:
   lvl=om[ni.get(q('val'))].find('w:lvl[@w:ilvl="'+il+'"]',N)
   if lvl is not None:
    f=lvl.find(q('numFmt'));fmt=f.get(q('val')) if f is not None else None
    li=lvl.find('w:pPr/w:ind',N)
    if li is not None:left=float(li.get(q('left'),'0'))
 if ind is not None and ind.get(q('left')) is not None:left=float(ind.get(q('left')))
 source.setdefault(t,[]).append({'numbered':np is not None,'format':fmt,'left':left or 0})
with ZipFile(root/'ICS v9 - Format Lampiran Permenkes.docx') as z:
 previous=E.fromstring(z.read('word/document.xml'));pn=E.fromstring(z.read('word/numbering.xml'))
pa={x.get(q('abstractNumId')):x for x in pn.findall(q('abstractNum'))};pm={x.get(q('numId')):pa[x.find(q('abstractNumId')).get(q('val'))] for x in pn.findall(q('num'))}
previous_kinds={}
for p in previous.xpath('//w:body/w:p[w:pPr/w:numPr]',namespaces=N):
 np=p.find('w:pPr/w:numPr',N);ni=np.find(q('numId'));il=np.find(q('ilvl'));il=il.get(q('val')) if il is not None else '0'
 if ni is not None and ni.get(q('val')) in pm:
  lv=pm[ni.get(q('val'))].find('w:lvl[@w:ilvl="'+il+'"]',N)
  if lv is not None:previous_kinds[normalize(text(p))]=lv.find(q('numFmt')).get(q('val'))
with ZipFile(target) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']);nr=E.fromstring(parts['word/numbering.xml']);active=False;context=0;headlevel=0;items=[]
for p in d.xpath('//w:body/w:p',namespaces=N):
 t=text(p);sid=p.xpath('./w:pPr/w:pStyle/@w:val',namespaces=N)
 if t.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
 if not active:continue
 if sid and sid[0] in ['Heading1','Heading2','Heading3','Heading4']:
  context+=1;headlevel=int(sid[0][-1])-1;continue
 if not t:continue
 candidates=source.get(normalize(t),[]);info=candidates[0] if candidates else None
 oldkind=previous_kinds.get(normalize(t))
 if oldkind and info and not info['numbered']:
  pp=p.find(q('pPr'));np=pp.find(q('numPr'))
  if np is not None:pp.remove(np)
  setel(pp,'ind',left=0,right=0,firstLine=0);setel(pp,'keepNext',val=1)
  for r in p.findall(q('r')):setel(child(r,'rPr'),'b',val=1)
  context+=1;continue
 if oldkind is None:
  continue
 child(child(p,'pPr'),'numPr')
 items.append({'p':p,'context':context,'head':headlevel,'info':info,'kind':oldkind})
groups={}
for item in items:groups.setdefault(item['context'],[]).append(item)
abids=[int(x.get(q('abstractNumId'))) for x in nr.findall(q('abstractNum'))];numids=[int(x.get(q('numId'))) for x in nr.findall(q('num'))]
aid=max(abids)+1;nid=max(numids)+1
base=next(x for x in nr.findall(q('abstractNum')) if x.xpath('./w:lvl/w:numFmt/@w:val',namespaces=N)[:5]==['upperLetter','decimal','lowerLetter','decimal','lowerLetter'])
stats={'restored_plain_labels':0,'restored_bullets':0,'nested_numbered_items':0,'unmatched':0}
for ctx,group in groups.items():
 vals=sorted({round(x['info']['left']) for x in group if x['info'] and x['info']['numbered'] and x['info']['format']!='bullet'})
 clusters=[]
 for v in vals:
  if not clusters or v-clusters[-1]>200:clusters.append(v)
 minimum=clusters[0] if clusters else min((x['info']['left'] for x in group if x['info']),default=0)
 ab=copy.deepcopy(base);ab.set(q('abstractNumId'),str(aid))
 for tag in ['nsid','tmpl']:setel(ab,tag,val=f'{aid:08X}')
 nr.insert(len(nr.findall(q('abstractNum'))),ab)
 nu=E.SubElement(nr,q('num'));nu.set(q('numId'),str(nid));setel(nu,'abstractNumId',val=aid)
 regularid=str(nid);aid+=1;nid+=1
 bulletid=None
 for item in group:
  p=item['p'];pp=p.find(q('pPr'));np=pp.find(q('numPr'));info=item['info'];baselevel=min(item['head']+1,4)
  if info and not info['numbered']:
   pp.remove(np);setel(pp,'ind',left=0,right=0,firstLine=0)
   if len(text(p))<140:
    setel(pp,'keepNext',val=1)
    for r in p.findall(q('r')):setel(child(r,'rPr'),'b',val=1)
   stats['restored_plain_labels']+=1;continue
  if info and (info['format']!='bullet' or item['kind']=='bullet' or clusters):
   delta=max(0,info['left']-minimum)
   depth=round(delta/425) if delta>200 else 0
  else:
   depth=0
   if info is None:stats['unmatched']+=1
  virtual=baselevel+depth
  isbullet=item['kind']=='bullet' or virtual>4
  if isbullet:
   virtual=min(virtual,5)
   if bulletid is None:
    ba=E.Element(q('abstractNum'));ba.set(q('abstractNumId'),str(aid));setel(ba,'nsid',val=f'{aid:08X}');setel(ba,'multiLevelType',val='singleLevel')
    lv=E.SubElement(ba,q('lvl'));lv.set(q('ilvl'),'0');setel(lv,'start',val=1);setel(lv,'numFmt',val='bullet');setel(lv,'lvlText',val='\u2022');setel(lv,'lvlJc',val='left')
    rp=child(lv,'rPr');setel(rp,'rFonts',ascii='Times New Roman',hAnsi='Times New Roman');setel(rp,'sz',val=24)
    nr.insert(len(nr.findall(q('abstractNum'))),ba)
    bn=E.SubElement(nr,q('num'));bn.set(q('numId'),str(nid));setel(bn,'abstractNumId',val=aid)
    bulletid=str(nid);aid+=1;nid+=1
   setel(np,'numId',val=bulletid);setel(np,'ilvl',val=0);stats['restored_bullets']+=1
  else:
   setel(np,'numId',val=regularid);setel(np,'ilvl',val=virtual)
   if depth:stats['nested_numbered_items']+=1
  setel(pp,'ind',left=(virtual+1)*425,right=0,hanging=425)
for name,el in [('word/document.xml',d),('word/numbering.xml',nr)]:parts[name]=E.tostring(el,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(target,'w',ZIP_DEFLATED) as z:
 for name,data in parts.items():z.writestr(name,data)
print(json.dumps(stats))
