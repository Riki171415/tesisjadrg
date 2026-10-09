from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';N={'w':W};q=lambda x:'{'+W+'}'+x
with ZipFile(Path(__file__).parent.parent/'ICS v9 - Siap Publikasi.docx') as z:d=E.fromstring(z.read('word/document.xml'))
active=False;left=340;bad=[];numbered=0;bullets=0
for p in d.xpath('//w:body/w:p',namespaces=N):
 t=''.join(p.xpath('./w:r/w:t/text()',namespaces=N)).strip();pp=p.find(q('pPr'));sid=p.xpath('./w:pPr/w:pStyle/@w:val',namespaces=N)
 if t.startswith('INDONESIAN DIAGNOSIS RELATED GROUP'):active=True
 if not active:continue
 if pp is None:
  if t and not t.startswith(('Tabel ','Gambar ')):bad.append((t[:80],'missing paragraph properties',left))
  continue
 ind=pp.find(q('ind'));actual=int(ind.get(q('left'),'0')) if ind is not None else 0
 if 'Tahapan Memilih Kode Tindakan Utama' in t:steps_parent=actual
 if t=='Catatan:' and 'steps_parent' in globals():left=steps_parent;del steps_parent
 if sid and sid[0] in ['Heading1','Heading2','Heading3','Heading4']:left=int(sid[0][-1])*425-85;continue
 np=pp.find(q('numPr'))
 if np is not None:
  il=np.find(q('ilvl'));v=int(il.get(q('val'),'0')) if il is not None else 0
  assert 1<=v<=4,(v,t[:100]);left=actual;numbered+=1;continue
 if t.startswith(('\u2022','\ufffd')):left=actual;bullets+=1;continue
 if not t or t.startswith(('Tabel ','Gambar ','Setting: Merujuk pada sesi;')) or p.xpath('.//w:drawing|.//w:pict|.//w:object',namespaces=N):continue
 if actual<left:bad.append((t[:80],actual,left))
assert not bad,bad[:10]
print({'paragraph_indent_audit':'passed','ordinary_numbered_points':numbered,'bullet_points':bullets})
