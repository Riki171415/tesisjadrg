from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
import copy, json, re

ROOT=Path(__file__).parent.parent
SOURCE=ROOT/'ICS v9 - Format Lampiran Permenkes.docx'
TARGET=ROOT/'ICS v9 - Siap Publikasi.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
N={'w':W}; q=lambda x:'{'+W+'}'+x
with ZipFile(SOURCE) as z: parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']); styles=E.fromstring(parts['word/styles.xml']); nr=E.fromstring(parts['word/numbering.xml'])
def child(el,tag):
    x=el.find(q(tag))
    if x is None: x=E.SubElement(el,q(tag))
    return x
def setel(el,tag,**attrs):
    x=child(el,tag);x.attrib.clear()
    for key,val in attrs.items():x.set(q(key),str(val))
    return x
def remove(el,*tags):
    for tag in tags:
        for x in el.findall(q(tag)):el.remove(x)
def font(rp,size):
    setel(rp,'rFonts',ascii='Times New Roman',hAnsi='Times New Roman',eastAsia='Times New Roman',cs='Times New Roman')
    setel(rp,'sz',val=size*2);setel(rp,'szCs',val=size*2)
    remove(rp,'spacing','position','w','kern','highlight','strike','dstrike')
def paragraph(pp,first=0,left=0,hanging=0,after=120,before=0,line=276,align='both'):
    remove(pp,'ind','spacing','jc','tabs','contextualSpacing','snapToGrid')
    attrs={'left':left,'right':0}
    if hanging:attrs['hanging']=hanging
    else:attrs['firstLine']=first
    setel(pp,'ind',**attrs);setel(pp,'spacing',before=before,after=after,line=line,lineRule='auto')
    setel(pp,'jc',val=align);setel(pp,'widowControl',val=1)
def runs(p,size):
    for r in p.xpath('./w:r|./w:hyperlink/w:r|./w:smartTag/w:r',namespaces=N):
        font(child(r,'rPr'),size)
    font(child(child(p,'pPr'),'rPr'),size)

# A single default replaces conflicting inherited paragraph settings.
defs=child(styles,'docDefaults');font(child(child(defs,'rPrDefault'),'rPr'),12)
paragraph(child(child(defs,'pPrDefault'),'pPr'),first=567)
for st in styles.findall(q('style')):
    sid=st.get(q('styleId'),'');rp=child(st,'rPr')
    if st.get(q('type')) in ['paragraph','character']:font(rp,12)
    if st.get(q('type'))!='paragraph':continue
    pp=child(st,'pPr')
    if sid=='Normal':paragraph(pp,first=567)
    elif sid in ['Heading1','Heading2','Heading3','Heading4']:
        lev=int(sid[-1])-1;paragraph(pp,left=lev*425,align='left',before=240 if lev==0 else 180,after=120)
        setel(pp,'keepNext',val=1);setel(pp,'keepLines',val=1);setel(rp,'b',val=1);font(rp,14 if lev==0 else 12)
    elif sid in ['Heading5','Heading6']:
        paragraph(pp,align='left',line=240,before=120,after=80);font(rp,10)
    elif sid.lower().startswith('toc'):
        paragraph(pp,align='left',line=240,after=40);remove(pp,'keepNext','keepLines')
    elif sid=='ListParagraph':paragraph(pp,first=0)

# Normalize every used hierarchy definition to five levels.
for ab in nr.findall(q('abstractNum')):
    formats=ab.xpath('./w:lvl/w:numFmt/@w:val',namespaces=N)
    if formats[:3]!=['upperLetter','decimal','lowerLetter']:continue
    for lev in ab.findall(q('lvl')):
        i=int(lev.get(q('ilvl')))
        if i>4:continue
        setel(lev,'numFmt',val=['upperLetter','decimal','lowerLetter','decimal','lowerLetter'][i])
        setel(lev,'lvlText',val=f'%{i+1}'+('.' if i<3 else ')'))
        setel(lev,'lvlRestart',val=0 if i==0 else i)
        setel(lev,'suff',val='tab');setel(lev,'lvlJc',val='left')
        pp=child(lev,'pPr');paragraph(pp,left=i*425+425,hanging=425,align='left')
        tabs=child(pp,'tabs');setel(tabs,'tab',val='num',pos=i*425+425)
        font(child(lev,'rPr'),14 if i==0 else 12)
        remove(child(lev,'rPr'),'b','bCs')

body=d.find(q('body')); active=False;current_level=0;parent_serial=0;grouped={};new_lists=0;removed_blanks=0
main=d.xpath('//w:body/w:p[w:pPr/w:pStyle[@w:val="Heading1"]][w:pPr/w:numPr]',namespaces=N)[0]
mainid=main.find('w:pPr/w:numPr/w:numId',N).get(q('val'))
base_num=next(x for x in nr.findall(q('num')) if x.get(q('numId'))==mainid)
base_ab=next(x for x in nr.findall(q('abstractNum')) if x.get(q('abstractNumId'))==base_num.find(q('abstractNumId')).get(q('val')))
next_a=max(int(x.get(q('abstractNumId'))) for x in nr.findall(q('abstractNum')))+1
next_n=max(int(x.get(q('numId'))) for x in nr.findall(q('num')))+1
stats={}
for p in list(body):
    if p.tag!=q('p'):continue
    text=''.join(p.xpath('./w:r/w:t/text()|./w:hyperlink/w:r/w:t/text()',namespaces=N)).strip()
    pp=child(p,'pPr');sid=pp.find(q('pStyle'));sid=sid.get(q('val')) if sid is not None else ''
    num=pp.find(q('numPr'));drawing=bool(p.xpath('.//w:drawing|.//w:pict|.//w:object',namespaces=N))
    if p is main:active=True
    if not active:
        if text:
            runs(p,14 if text in ['DAFTAR ISI','DAFTAR GAMBAR','DAFTAR TABEL','DAFTAR SINGKATAN','GLOSARIUM'] else 12)
            if text in ['DAFTAR ISI','DAFTAR GAMBAR','DAFTAR TABEL','DAFTAR SINGKATAN','GLOSARIUM']:
                paragraph(pp,align='left',before=0,after=240);setel(pp,'keepNext',val=1)
        continue
    # Clear inherited/manual pagination but retain structural section boundaries.
    remove(pp,'keepNext','keepLines','pageBreakBefore','rPr')
    for br in p.xpath('./w:r/w:br[@w:type="page"]',namespaces=N):br.getparent().remove(br)
    if not text and not drawing:
        if not p.xpath('.//w:sectPr|.//w:bookmarkStart|.//w:bookmarkEnd|.//w:fldChar|.//w:object',namespaces=N):
            body.remove(p);removed_blanks+=1;continue
        paragraph(pp,after=0,line=240,align='left');runs(p,1);continue
    if sid in ['Heading1','Heading2','Heading3','Heading4']:
        lev=int(sid[-1])-1;current_level=lev;parent_serial+=1
        paragraph(pp,left=lev*425+425,hanging=425,before=0 if lev==0 else 180,after=120,align='left')
        setel(pp,'keepNext',val=1);setel(pp,'keepLines',val=1)
        if lev==0:setel(pp,'pageBreakBefore',val=1)
        runs(p,14 if lev==0 else 12)
        for rp in p.xpath('./w:r/w:rPr|./w:pPr/w:rPr',namespaces=N):setel(rp,'b',val=1)
        stats['headings']=stats.get('headings',0)+1
    elif sid in ['Heading5','Heading6'] or re.match(r'^(Tabel|Gambar) [ABC]\.\d+\.',text):
        paragraph(pp,before=120,after=80,line=240,align='left');runs(p,10)
        setel(pp,'keepLines',val=1)
        if text.startswith('Tabel '):setel(pp,'keepNext',val=1)
        stats['captions']=stats.get('captions',0)+1
    elif drawing:
        paragraph(pp,after=80,line=240,align='center');runs(p,12);setel(pp,'keepNext',val=1)
    elif num is not None:
        # All numbered/bulleted prose lists become children of their actual heading.
        lev=min(current_level+1,4);key=(parent_serial,lev)
        if key not in grouped:
            ab=copy.deepcopy(base_ab);ab.set(q('abstractNumId'),str(next_a))
            for tag in ['nsid','tmpl']:setel(ab,tag,val=f'{next_a:08X}')
            nr.insert(len(nr.findall(q('abstractNum'))),ab)
            nu=E.SubElement(nr,q('num'));nu.set(q('numId'),str(next_n));setel(nu,'abstractNumId',val=next_a)
            ov=E.SubElement(nu,q('lvlOverride'));ov.set(q('ilvl'),str(lev));setel(ov,'startOverride',val=1)
            grouped[key]=str(next_n);next_a+=1;next_n+=1;new_lists+=1
        setel(num,'ilvl',val=lev);setel(num,'numId',val=grouped[key])
        paragraph(pp,left=lev*425+425,hanging=425);runs(p,12)
        # ListParagraph's old style may carry unrelated numbering/indent.
        remove(pp,'pStyle')
        stats['lists']=stats.get('lists',0)+1
    else:
        paragraph(pp,first=567);runs(p,12)
        if text.endswith(':') and len(text)<100:
            paragraph(pp,first=0);setel(pp,'keepNext',val=1)
        stats['narrative']=stats.get('narrative',0)+1

# Apply the same readable typesetting to every table, including examples.
tables=d.xpath('//w:tbl',namespaces=N)
for tbl in tables:
    tp=child(tbl,'tblPr');setel(tp,'tblW',w=9026,type='dxa');setel(tp,'tblLayout',type='fixed');setel(tp,'jc',val='center')
    setel(tp,'tblInd',w=0,type='dxa')
    cm=child(tp,'tblCellMar')
    for side in ['top','left','bottom','right']:setel(cm,side,w=80,type='dxa')
    grid=tbl.find(q('tblGrid'));cols=grid.findall(q('gridCol')) if grid is not None else []
    widths=[int(x.get(q('w'),'1')) for x in cols];total=sum(widths)
    if total:
        scaled=[round(w*9026/total) for w in widths];scaled[-1]+=9026-sum(scaled)
        for c,w in zip(cols,scaled):c.set(q('w'),str(w))
    rows=tbl.findall(q('tr'))
    for ri,row in enumerate(rows):
        trp=child(row,'trPr');remove(trp,'trHeight','tblHeader','cantSplit')
        if ri==0:setel(trp,'tblHeader',val=1)
        # Short rows stay together; long narrative rows may flow over a page.
        rowtext=''.join(row.xpath('.//w:t/text()',namespaces=N))
        if len(rowtext)<1200:setel(trp,'cantSplit',val=1)
        if ri==1 and rows and 'Contoh ' in ''.join(rows[0].xpath('./w:tc[1]//w:t/text()',namespaces=N)) and any(v in rowtext for v in ['Diagnosis Dokter','Penulisan Diagnosis','Pengode Menetapkan']):setel(trp,'tblHeader',val=1)
        col=0
        for cell in row.findall(q('tc')):
            cp=child(cell,'tcPr');span=cp.find(q('gridSpan'));span=int(span.get(q('val'))) if span is not None else 1
            if total and col+span<=len(scaled):setel(cp,'tcW',w=sum(scaled[col:col+span]),type='dxa')
            col+=span;setel(cp,'vAlign',val='top');remove(cp,'noWrap')
            for p in cell.xpath('./w:p',namespaces=N):
                text=''.join(p.xpath('./w:r/w:t/text()',namespaces=N));pp=child(p,'pPr')
                remove(pp,'keepNext','keepLines','pageBreakBefore','outlineLvl')
                paragraph(pp,after=40,line=240,align='both' if len(text)>160 else 'left');runs(p,10)
                if ri==0 and not text.startswith('Contoh') and len(rowtext)<200:
                    for rp in p.xpath('./w:r/w:rPr',namespaces=N):setel(rp,'b',val=1)

# TOC fields live in content controls; do not apply narrative indents to them.
for p in d.xpath('//w:sdtContent//w:p[not(ancestor::w:tc)]',namespaces=N):
    pp=child(p,'pPr');sid=pp.find(q('pStyle'));sid=sid.get(q('val'),'') if sid is not None else ''
    level=int(sid[-1])-1 if re.match(r'^TOC[1-9]$',sid,re.I) else 0
    oldtabs=copy.deepcopy(pp.find(q('tabs')))
    paragraph(pp,left=level*425,after=40,line=240,align='left');runs(p,12)
    if oldtabs is not None:pp.append(oldtabs)
    remove(pp,'keepNext','keepLines')

for name in list(parts):
    if re.match(r'word/(header|footer)\d+\.xml$',name):
        x=E.fromstring(parts[name])
        for p in x.xpath('//w:p[not(ancestor::w:txbxContent)]',namespaces=N):
            runs(p,10);pp=child(p,'pPr');setel(pp,'spacing',before=0,after=0,line=240,lineRule='auto');setel(pp,'ind',left=0,right=0,firstLine=0)
        parts[name]=E.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True)
for name,x in [('word/document.xml',d),('word/styles.xml',styles),('word/numbering.xml',nr)]:parts[name]=E.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(TARGET,'w',ZIP_DEFLATED) as z:
    for name,data in parts.items():z.writestr(name,data)
print(json.dumps({'output':str(TARGET),'formatted':stats,'tables':len(tables),'removed_empty_paragraphs':removed_blanks,'child_list_groups':new_lists}))
