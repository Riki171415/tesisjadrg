import zipfile,lxml.etree as E,re,json,hashlib,io
from PIL import Image
from pathlib import Path
root=Path(__file__).parent.parent
n={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
def read(name):
    with zipfile.ZipFile(root/name) as z:
        meaningful=[]
        for x in z.namelist():
            if x.startswith('word/media/'):
                data=z.read(x)
                try:
                    im=Image.open(io.BytesIO(data))
                    if im.size==(1,1) and im.getbbox() is None:continue
                except Exception:pass
                meaningful.append(hashlib.sha256(data).hexdigest())
        return E.fromstring(z.read('word/document.xml')),E.fromstring(z.read('word/numbering.xml')),set(meaningful)
a,_,am=read('ICS v9.docx');b,nr,bm=read('ICS v9 - Format Lampiran Permenkes.docx')
def tabletext(tbl):
    cells=tbl.xpath('./w:tr[1]/w:tc[1]',namespaces=n)
    clone=E.fromstring(E.tostring(tbl))
    first=clone.xpath('./w:tr[1]/w:tc[1]',namespaces=n)
    if first:
        t=''.join(first[0].xpath('.//w:t/text()',namespaces=n))
        if re.match(r'^\s*contoh\s*\d+(?:\s*-\s*\d+)?\s*$',t,re.I):
            nodes=first[0].xpath('.//w:t',namespaces=n)
            for x in nodes:x.text=''
            if nodes:nodes[0].text='Contoh'
    return ''.join(clone.xpath('.//w:t/text()',namespaces=n))
ta=[tabletext(x) for x in a.xpath('//w:tbl',namespaces=n)]
tb=[tabletext(x) for x in b.xpath('//w:tbl',namespaces=n)]
assert ta==tb,'Table text changed'
assert am==bm,'Media count changed'
captions={}
for p in b.xpath('//w:body/w:p',namespaces=n):
    t=''.join(p.xpath('./w:r/w:t/text()',namespaces=n));m=re.match(r'^(Tabel|Gambar) ([ABC])\.(\d+)\.',t)
    if m:captions.setdefault((m[1],m[2]),[]).append(int(m[3]))
for key,vals in captions.items():assert vals==list(range(1,len(vals)+1)),key
good=[]
for ab in nr.findall('w:abstractNum',n):
    fmts=ab.xpath('./w:lvl/w:numFmt/@w:val',namespaces=n)
    if fmts[:4]==['upperLetter','decimal','lowerLetter','decimal']:good.append(ab.get('{'+n['w']+'}abstractNumId'))
assert good,'Four-level numbering not preserved by Word'
assert not any('No table of contents entries found' in t for t in b.xpath('//w:t/text()',namespaces=n)),'Unpopulated table of contents'
bodycount=0;tablecount=0;bodyexamples=0;tableexamples=0;active=False
for x in b.find('w:body',n):
    if E.QName(x).localname=='p':
        text=''.join(x.xpath('./w:r/w:t/text()',namespaces=n)).strip();sty=x.xpath('./w:pPr/w:pStyle/@w:val',namespaces=n)
        if text.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
        if not active:continue
        if sty and sty[0] in ['Heading1','Heading2','Heading3','Heading4']:
            bodycount=0
            if sty[0] in ['Heading1','Heading2']:tablecount=0
        matches=list(re.finditer(r'\bContoh\s+(\d+):',text,re.I))
        for m in matches:
            bodycount+=1;bodyexamples+=1;assert int(m[1])==bodycount,'Body example gap'
            if m.start()==0:assert not x.xpath('./w:pPr/w:numPr',namespaces=n),'Double-numbered example'
    elif E.QName(x).localname=='tbl' and active:
        first=x.xpath('./w:tr[1]/w:tc[1]',namespaces=n)
        t=''.join(first[0].xpath('.//w:t/text()',namespaces=n)).strip() if first else ''
        m=re.fullmatch(r'Contoh (\d+)',t)
        if m:tablecount+=1;tableexamples+=1;assert int(m[1])==tablecount,'Table example gap'
print('Example numbering verified:',bodyexamples,'prose labels;',tableexamples,'table labels.')
print(json.dumps({'table_text_unchanged':len(ta),'media_preserved':len(bm),'caption_sequences':{str(k):len(v) for k,v in captions.items()},'four_levels_in_Word':good}))
