from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
import copy
path=Path(__file__).with_name('publish_list_hierarchy.py')
code=path.read_text(encoding='utf-8')
s={'__file__':str(path),'__name__':'dry_run'}
exec(code[:code.index("for name,el in [('word/document.xml'")],s)
d,nr,parts=s['d'],s['nr'],s['parts']
q,N,child,setel=s['q'],s['N'],s['child'],s['setel']
aid=max(int(x.get(q('abstractNumId'))) for x in nr.findall(q('abstractNum')))+1
nid=max(int(x.get(q('numId'))) for x in nr.findall(q('num')))+1
bullets=0
for group in s['groups'].values():
 levels=sorted({int(x['p'].find('w:pPr/w:ind',N).get(q('left')))//425-1 for x in group})
 base=min(group[0]['head']+1,4)
 mapping={v:base+i for i,v in enumerate(levels)}
 ab=copy.deepcopy(s['base']);ab.set(q('abstractNumId'),str(aid))
 for tag in ['nsid','tmpl']:setel(ab,tag,val=f'{aid:08X}')
 nr.insert(len(nr.findall(q('abstractNum'))),ab)
 nu=E.SubElement(nr,q('num'));nu.set(q('numId'),str(nid));setel(nu,'abstractNumId',val=aid)
 for item in group:
  p=item['p'];pp=child(p,'pPr');ind=pp.find(q('ind'));v=mapping[int(ind.get(q('left')))//425-1]
  np=pp.find(q('numPr'))
  if item['kind']=='bullet' or v>4:
   v=min(v,5)
   if np is not None:pp.remove(np)
   setel(pp,'pStyle',val='Normal')
   r=E.Element(q('r'));rp=child(r,'rPr');setel(rp,'rFonts',ascii='Times New Roman',hAnsi='Times New Roman');setel(rp,'sz',val=24);setel(rp,'b',val=0)
   E.SubElement(r,q('t')).text='•';E.SubElement(r,q('tab'));p.insert(1,r);bullets+=1
  else:
   setel(np,'numId',val=nid);setel(np,'ilvl',val=v)
  setel(pp,'ind',left=(v+1)*425,right=0,hanging=425)
  tabs=child(pp,'tabs');tabs.clear();tab=E.SubElement(tabs,q('tab'));tab.set(q('val'),'left');tab.set(q('pos'),str((v+1)*425))
 aid+=1;nid+=1
active=False;current=425;count=0
for p in d.xpath('//w:body/w:p',namespaces=N):
 t=''.join(p.xpath('./w:r/w:t/text()',namespaces=N)).strip()
 sid=p.xpath('./w:pPr/w:pStyle/@w:val',namespaces=N)
 if t.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
 if not active:continue
 pp=child(p,'pPr');np=pp.find(q('numPr'))
 if sid and sid[0] in ['Heading1','Heading2','Heading3','Heading4']:
  current=int(sid[0][-1])*425;continue
 if np is not None or t.startswith('•'):
  current=int(pp.find(q('ind')).get(q('left')));continue
 if not t or p.xpath('.//w:drawing|.//w:pict|.//w:object',namespaces=N):continue
 if t.startswith(('Tabel ','Gambar ')):continue
 short=(len(t)<140 and (t.endswith(':') or pp.find(q('keepNext')) is not None))
 setel(pp,'ind',left=current,right=0,firstLine=0 if short else 567)
 count+=1
for name,el in [('word/document.xml',d),('word/numbering.xml',nr)]:parts[name]=E.tostring(el,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(s['target'],'w',ZIP_DEFLATED) as z:
 for name,data in parts.items():z.writestr(name,data)
print({'contextual_paragraphs':count,'literal_bullets':bullets})
