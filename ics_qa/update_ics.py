import zipfile, re, collections, json
from pathlib import Path
from lxml import etree as E

ROOT=Path(__file__).parent
SRC=ROOT/'ICS v9.docx'
OUT=ROOT/'ICS v9 - Format Lampiran Permenkes.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS={'w':W}
def q(s):return '{'+W+'}'+s
def el(s,**attrs):
    x=E.Element(q(s))
    for k,v in attrs.items():x.set(q(k),str(v))
    return x
def ensure(x,s):
    y=x.find(q(s))
    if y is None:y=el(s);x.append(y)
    return y
def setprop(x,s,**attrs):
    y=ensure(x,s);y.attrib.clear()
    for k,v in attrs.items():y.set(q(k),str(v))
    return y
def rm(x,*names):
    for name in names:
        for y in x.findall(q(name)):x.remove(y)
def pp(p):
    x=p.find(q('pPr'))
    if x is None:x=el('pPr');p.insert(0,x)
    return x
def texts(p):return p.xpath('./w:r/w:t|./w:hyperlink/w:r/w:t',namespaces=NS)
def text(p):return ''.join(t.text or '' for t in texts(p))
def replace_prefix(p,length,new=''):
    ts=texts(p);remaining=length
    for t in ts:
        s=t.text or '';take=min(remaining,len(s));t.text=s[take:];remaining-=take
        if remaining==0:break
    if ts:ts[0].text=new+(ts[0].text or '')
def formatp(p,heading=False,level=0):
    pr=pp(p);rm(pr,'tabs','pBdr')
    setprop(pr,'jc',val='left' if heading else 'both')
    setprop(pr,'spacing',before=120 if heading else 0,after=80 if heading else 100,line=276,lineRule='auto')
    setprop(pr,'ind',left=level*360,firstLine=0)
    setprop(pr,'widowControl',val=1)
    if heading:setprop(pr,'keepNext',val=1)
    else:rm(pr,'keepNext')
    for r in p.findall(q('r')):
        rp=ensure(r,'rPr')
        setprop(rp,'rFonts',ascii='Times New Roman',hAnsi='Times New Roman',eastAsia='Times New Roman',cs='Times New Roman')
        setprop(rp,'sz',val=24);setprop(rp,'szCs',val=24)
        if heading:setprop(rp,'b',val=1);setprop(rp,'color',val='000000')

with zipfile.ZipFile(SRC) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
doc=E.fromstring(parts['word/document.xml']); numbering=E.fromstring(parts['word/numbering.xml'])
aid=max(int(a.get(q('abstractNumId'))) for a in numbering.findall(q('abstractNum')))+1
nid=max(int(a.get(q('numId'))) for a in numbering.findall(q('num')))+1
ab=el('abstractNum',abstractNumId=aid);ab.append(el('multiLevelType',val='multilevel'))
for i,(fmt,mark) in enumerate([('upperLetter','%1.'),('decimal','%2.'),('lowerLetter','%3.'),('decimal','%4)')]):
    lev=el('lvl',ilvl=i);lev.extend([el('start',val=1),el('numFmt',val=fmt),el('lvlText',val=mark),el('lvlJc',val='left'),el('suff',val='tab')])
    if i:lev.append(el('lvlRestart',val=i))
    pr=el('pPr');tabs=el('tabs');tabs.append(el('tab',val='num',pos=(i+1)*360));pr.append(tabs);pr.append(el('ind',left=(i+1)*360,hanging=360));lev.append(pr)
    rp=el('rPr');rp.append(el('rFonts',ascii='Times New Roman',hAnsi='Times New Roman'));lev.append(rp);ab.append(lev)
numbering.append(ab)
def new_num():
    global nid
    cur=nid;nid+=1;x=el('num',numId=cur);x.append(el('abstractNumId',val=aid));numbering.append(x);return cur
mainnum=new_num()
def assign(p,level,num=mainnum):
    pr=pp(p);rm(pr,'numPr');np=el('numPr');np.extend([el('ilvl',val=level),el('numId',val=num)]);pr.append(np)
    setprop(pr,'ind',left=(level+1)*360,hanging=360)
    tabs=ensure(pr,'tabs');tabs.append(el('tab',val='num',pos=(level+1)*360))

body=doc.find(q('body'));ps=body.findall(q('p'))
chapter=0;current_level=0;mdc=False;wrapper=False;tablecounts=collections.Counter();figcounts=collections.Counter();changes=[];headingcount=collections.Counter();listgroup=None;groupnum=None
# Explicit semantic corrections discovered in the source.
special={34:1,35:2,43:2,47:2,51:2,613:2,623:3,626:3}
for idx,p in enumerate(ps):
    t=text(p);st=p.xpath('./w:pPr/w:pStyle/@w:val',namespaces=NS);style=st[0] if st else ''
    m=re.match(r'^\s*CHAPTER\s+(I{1,3})\s+',t)
    if m:
        chapter=len(m[1]);replace_prefix(p,m.end());formatp(p,True,0);assign(p,0);setprop(pp(p),'pStyle',val='Heading1');setprop(pp(p),'pageBreakBefore',val=1)
        current_level=0;mdc=False;wrapper=False;listgroup=None;headingcount[0]+=1;changes.append([idx,'chapter',t,text(p)]);continue
    if not chapter:continue
    if not t.strip():
        if style.startswith('Heading'):rm(pp(p),'pStyle','numPr','keepNext');setprop(pp(p),'spacing',before=0,after=0,line=120,lineRule='auto')
        continue
    cap=re.match(r'^\s*(Tabel|Gambar)\s+(\d+(?:\.\d+)+)\.?\s*',t,re.I)
    if cap:
        kind=cap[1].capitalize();counter=tablecounts if kind=='Tabel' else figcounts;counter[chapter]+=1
        new=f'{kind} {"ABC"[chapter-1]}.{counter[chapter]}. '
        replace_prefix(p,cap.end(),new);pr=pp(p);rm(pr,'numPr');setprop(pr,'pStyle',val='Heading5' if kind=='Tabel' else 'Heading6');formatp(p,False);setprop(pr,'jc',val='center');setprop(pr,'keepNext',val=1 if kind=='Tabel' else 0);setprop(pr,'outlineLvl',val=9)
        changes.append([idx,'caption',t,text(p)]);continue
    prefix=re.match(r'^\s*\d+(?:\.\d+){1,4}\.?\s*',t)
    wrapper_heading=bool(re.match(r'^\s*(PENJELASAN(?: KHUSUS)? KODE ICD|PERHATIAN KHUSUS)',t,re.I))
    level=None
    if idx in special:level=special[idx]
    elif style in ['Heading2','Heading3','Heading4'] or (prefix and prefix[0].strip().startswith(str(chapter)+'.')):
        if wrapper_heading:level=2;wrapper=True
        elif style=='Heading2':level=1;mdc=bool(re.search(r'MDC\s+\d',t));wrapper=False
        elif prefix:
            depth=len(re.findall(r'\d+',prefix[0]))-1
            level=min(depth,3)
            if mdc:level=3 if wrapper else 2
        elif style in ['Heading3','Heading4']:level=min(int(style[-1])-1,3)
    elif wrapper_heading:level=2;wrapper=True
    if level is not None:
        if prefix:replace_prefix(p,prefix.end())
        elif idx in (623,626):
            mm=re.match(r'^\s*\d+\.\s*',t)
            if mm:replace_prefix(p,mm.end())
        formatp(p,True,level);assign(p,level);setprop(pp(p),'pStyle',val='Heading'+str(level+1));setprop(pp(p),'outlineLvl',val=level)
        current_level=level;listgroup=None;headingcount[level]+=1;changes.append([idx,'heading',t,text(p)]);continue
    oldnum=p.xpath('./w:pPr/w:numPr/w:numId/@w:val',namespaces=NS)
    oldlevel=p.xpath('./w:pPr/w:numPr/w:ilvl/@w:val',namespaces=NS)
    drawing=bool(p.xpath('.//w:drawing|.//w:pict',namespaces=NS))
    if drawing:continue
    formatp(p,False,current_level)
    if oldnum:
        target=min(current_level+1+(int(oldlevel[0]) if oldlevel else 0),4)
        key=(oldnum[0],target)
        if target<=3:
            if listgroup!=key:groupnum=new_num();listgroup=key
            assign(p,target,groupnum)
        else:
            # Detail beyond the requested four levels becomes prose; preserve its text.
            rm(pp(p),'numPr');setprop(pp(p),'ind',left=1440,firstLine=0)
    else:listgroup=None

# Preserve table data and geometry; eliminate exact heights that can clip text.
for tr in doc.xpath('//w:tbl/w:tr',namespaces=NS):
    for h in tr.xpath('./w:trPr/w:trHeight',namespaces=NS):
        if h.get(q('hRule'))=='exact':h.set(q('hRule'),'atLeast')
# Uniform original A4 page baseline: 2.54 cm margins.
for sect in doc.xpath('//w:sectPr',namespaces=NS):
    mar=ensure(sect,'pgMar')
    for k in ['left','right','top','bottom']:mar.set(q(k),'1440')
styles=E.fromstring(parts['word/styles.xml'])
for s in styles.findall(q('style')):
    sid=s.get(q('styleId'),'')
    if sid in ['Normal','Heading1','Heading2','Heading3','Heading4','Heading5','Heading6']:
        pr=ensure(s,'pPr');rm(pr,'pBdr','tabs','ind');setprop(pr,'spacing',before=0,after=100,line=276,lineRule='auto')
        if sid=='Normal':setprop(pr,'jc',val='both')
        else:setprop(pr,'jc',val='left');setprop(pr,'outlineLvl',val=int(sid[-1])-1 if int(sid[-1])<=4 else 9)
        rp=ensure(s,'rPr');setprop(rp,'rFonts',ascii='Times New Roman',hAnsi='Times New Roman',cs='Times New Roman',eastAsia='Times New Roman');setprop(rp,'sz',val=24)
        if sid!='Normal':setprop(rp,'color',val='000000')
settings=E.fromstring(parts['word/settings.xml']);setprop(settings,'updateFields',val='true')
for fld in doc.xpath('//w:fldChar[@w:fldCharType="begin"]',namespaces=NS):fld.set(q('dirty'),'true')
# Word requires canonical OOXML child order; otherwise list levels may be silently lost.
orders={
 'pPr':'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
 'rPr':'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange',
 'lvl':'start numFmt lvlRestart pStyle isLgl suff lvlText lvlPicBulletId legacy lvlJc pPr rPr',
 'abstractNum':'nsid multiLevelType tmpl name styleLink numStyleLink lvl',
 'numbering':'numPicBullet abstractNum num numIdMacAtCleanup'}
for root in [doc,numbering,styles,settings]:
    for x in root.iter():
        name=E.QName(x).localname
        if name in orders:
            rank={v:i for i,v in enumerate(orders[name].split())}
            x[:]=sorted(x,key=lambda y:rank.get(E.QName(y).localname,999))
# Headings and captions are still the original Word TOC sources.
for key,x in [('word/document.xml',doc),('word/numbering.xml',numbering),('word/styles.xml',styles),('word/settings.xml',settings)]:parts[key]=E.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True)
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
    for name,data in parts.items():z.writestr(name,data)
qa=ROOT/'ics_qa';qa.mkdir(exist_ok=True)
(qa/'changes.json').write_text(json.dumps(changes,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'output':str(OUT),'headings':dict(headingcount),'tables':dict(tablecounts),'figures':dict(figcounts),'preserved_table_objects':len(doc.xpath('//w:tbl',namespaces=NS))}))
