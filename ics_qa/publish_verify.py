from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
from collections import Counter
import hashlib,re,json
ROOT=Path(__file__).parent.parent;N={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
def read(name):
    with ZipFile(ROOT/name) as z:
        return E.fromstring(z.read('word/document.xml')),E.fromstring(z.read('word/numbering.xml')),{hashlib.sha256(z.read(x)).hexdigest() for x in z.namelist() if x.startswith('word/media/')}
a,_,am=read('ICS v9 - Format Lampiran Permenkes.docx');b,num,bm=read('ICS v9 - Siap Publikasi.docx')
def prose(d):
    values=[''.join(p.xpath('.//w:t/text()',namespaces=N)) for p in d.xpath('//w:body/w:p',namespaces=N) if ''.join(p.xpath('.//w:t/text()',namespaces=N)).strip()]
    values=[v.replace('\u2022','') for v in values]
    values=[v.replace('Merujuk pada sesiMerujuk pada sesi','Merujuk pada sesi').replace('Kuantitas berdasarkan lokasiKuantitas berdasarkan lokasi','Kuantitas berdasarkan lokasi').replace('Setting: Merujuk pada sesi; Multiplicity: Kuantitas berdasarkan lokasi.','Merujuk pada sesiKuantitas berdasarkan lokasi') for v in values]
    return values
assert ''.join(prose(a))==''.join(prose(b)),'Prose text changed'
def tables(d):
    output=[]
    for i,t in enumerate(d.xpath('//w:tbl',namespaces=N)):
        clone=E.fromstring(E.tostring(t))
        if i==0:
            row=clone.find('w:tr',N)
            cells=[''.join(c.xpath('.//w:t/text()',namespaces=N)) for c in row.findall('w:tc',N)]
            if cells==['Singkatan','Kepanjangan']:clone.remove(row)
        value=''.join(clone.xpath('.//w:t/text()',namespaces=N))
        value=value.replace('Pasien dalam contoh 8 (dengan diagnosis utama deep vein thrombosis)','Pasien dalam contoh 5 (dengan diagnosis utama deep vein thrombosis)')
        output.append(value)
    return output
assert tables(a)==tables(b),'Table content changed'
with ZipFile(ROOT/'ICS v9 - Format Lampiran Permenkes.docx') as z:
    original_cover=hashlib.sha256(z.read('word/media/image1.png')).hexdigest()
clean_cover=hashlib.sha256((ROOT/'ics_qa/cover-clean.png').read_bytes()).hexdigest()
assert am-{original_cover}==bm-{clean_cover},'Technical images changed'
cap={}
for p in b.xpath('//w:body/w:p',namespaces=N):
    text=''.join(p.xpath('./w:r/w:t/text()',namespaces=N));m=re.match(r'^(Tabel|Gambar) ([ABC])\.(\d+)\.',text)
    if m:cap.setdefault((m[1],m[2]),[]).append(int(m[3]))
for key,vals in cap.items():assert vals==list(range(1,len(vals)+1)),key
assert len(cap['Tabel','C'])==135
defs={x.get('{'+N['w']+'}abstractNumId'):x for x in num.findall('w:abstractNum',N)}
numdefs={x.get('{'+N['w']+'}numId'):defs[x.find('w:abstractNumId',N).get('{'+N['w']+'}val')] for x in num.findall('w:num',N)}
used=b.xpath('//w:body/w:p/w:pPr/w:numPr/w:numId/@w:val',namespaces=N)
for nid in set(used):
    fmts=numdefs[nid].xpath('./w:lvl/w:numFmt/@w:val',namespaces=N)
    if fmts[0]=='bullet':continue
    assert fmts[:5]==['upperLetter','decimal','lowerLetter','decimal','lowerLetter'],nid
    assert numdefs[nid].xpath('./w:lvl/w:lvlText/@w:val',namespaces=N)[:5]==['%1.','%2.','%3.','%4)','%5)'],nid
assert not any('No table of contents entries found' in t for t in b.xpath('//w:t/text()',namespaces=N))
prose_examples=table_examples=0;pc=tc=0;active=False
for item in b.find('w:body',N):
    if E.QName(item).localname=='p':
        text=''.join(item.xpath('./w:r/w:t/text()',namespaces=N)).strip()
        sid=item.xpath('./w:pPr/w:pStyle/@w:val',namespaces=N)
        if text.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
        if not active:continue
        if sid and sid[0] in ['Heading1','Heading2','Heading3','Heading4']:
            pc=0
            if sid[0] in ['Heading1','Heading2']:tc=0
        for match in re.finditer(r'\bContoh\s+(\d+):',text,re.I):
            pc+=1;prose_examples+=1;assert int(match[1])==pc,'Prose example numbering gap'
    elif E.QName(item).localname=='tbl' and active:
        first=item.xpath('./w:tr[1]/w:tc[1]',namespaces=N)
        text=''.join(first[0].xpath('.//w:t/text()',namespaces=N)).strip() if first else ''
        match=re.fullmatch(r'Contoh (\d+)',text)
        if match:tc+=1;table_examples+=1;assert int(match[1])==tc,'Table example numbering gap'
assert (prose_examples,table_examples)==(102,133)
print('Example sequences verified:',prose_examples,'in prose;',table_examples,'in tables.')
print(json.dumps({'prose_paragraphs_preserved':len(prose(b)),'tables_preserved':len(tables(b)),'technical_image_files_preserved':len(am-{original_cover}),'cover_draft_stamp_removed':clean_cover in bm,'captions':{str(k):len(v) for k,v in cap.items()},'used_five_level_lists':len(set(used)),'levels':dict(Counter(b.xpath('//w:body/w:p/w:pPr/w:numPr/w:ilvl/@w:val',namespaces=N)))}))
