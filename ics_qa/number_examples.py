import zipfile,re,json
from pathlib import Path
from lxml import etree as E
root=Path(__file__).parent.parent;path=root/'ICS v9 - Format Lampiran Permenkes.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';n={'w':W}
def q(s):return '{'+W+'}'+s
def rewrite(nodes,start,end,replacement):
    pos=0;inserted=False
    for x in nodes:
        s=x.text or '';lo=max(start-pos,0);hi=min(end-pos,len(s))
        if hi>lo:
            x.text=s[:lo]+(replacement if not inserted else '')+s[hi:];inserted=True
        pos+=len(s)
def tx(p):return ''.join(p.xpath('.//w:t/text()',namespaces=n))
with zipfile.ZipFile(path) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
d=E.fromstring(parts['word/document.xml']);body=d.find(q('body'));topic='';section='';bodycount=0;tablecount=0;changes=[];active=False
for x in body:
    if x.tag==q('p'):
        text=tx(x).strip();sty=x.xpath('./w:pPr/w:pStyle/@w:val',namespaces=n)
        if text.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
        if not active:continue
        if sty and sty[0] in ['Heading1','Heading2','Heading3','Heading4']:
            topic=text;bodycount=0
            if sty[0] in ['Heading1','Heading2']:section=text;tablecount=0
        nodes=x.xpath('./w:r/w:t|./w:hyperlink/w:r/w:t',namespaces=n);outer=''.join(a.text or '' for a in nodes)
        m=re.match(r'^\s*contoh\s*(?:\d+(?:-\d+)?\s*[:.)]?|:|kasus\s*$)\s*',outer,re.I)
        if m:
            bodycount+=1;new=f'Contoh {bodycount}: '
            rewrite(nodes,0,m.end(),new)
            pr=x.find(q('pPr'))
            if pr is not None:
                for num in pr.findall(q('numPr')):pr.remove(num)
            changes.append({'type':'body','topic':topic,'old':m[0],'new':new})
        # Update the one numbered table cross-reference still present in the prose.
        nodes=x.xpath('./w:r/w:t|./w:hyperlink/w:r/w:t',namespaces=n);outer=''.join(a.text or '' for a in nodes)
        for m in reversed(list(re.finditer(r'\bTabel\s+1\.4\b',outer,re.I))):rewrite(nodes,m.start(),m.end(),'Tabel A.4')
    elif x.tag==q('tbl') and active:
        cells=x.xpath('./w:tr[1]/w:tc[1]',namespaces=n)
        if not cells:continue
        cell=cells[0];nodes=cell.xpath('.//w:t',namespaces=n);text=''.join(a.text or '' for a in nodes)
        m=re.match(r'^\s*contoh(?:\s*\d+(?:\s*-\s*\d+)?)?\s*$',text,re.I)
        if m:
            tablecount+=1;new=f'Contoh {tablecount}';rewrite(nodes,0,len(text),new)
            changes.append({'type':'table','section':section,'old':text,'new':new})
parts['word/document.xml']=E.tostring(d,xml_declaration=True,encoding='UTF-8',standalone=True)
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
    for key,val in parts.items():z.writestr(key,val)
(root/'ics_qa/example_changes.json').write_text(json.dumps(changes,ensure_ascii=False,indent=2),encoding='utf-8')
print('Example labels normalized:',sum(x['type']=='body' for x in changes),'in prose;',sum(x['type']=='table' for x in changes),'in tables.')
