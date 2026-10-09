from pathlib import Path
import pypdfium2 as pdf
import json,re
root=Path(__file__).parent;doc=pdf.PdfDocument(str(root/'final-preview.pdf'));report=[];alltext=[];outside=[]
for i in range(len(doc)):
 p=doc[i];tp=p.get_textpage();t=tp.get_text_range();alltext.append(t);w,h=p.get_size();bad=[]
 for j in range(tp.count_chars()):
  char=tp.get_text_range(j,1)
  if not char.strip():continue
  x0,y0,x1,y1=tp.get_charbox(j)
  if x0<18 or x1>w-18 or y0<18 or y1>h-18:bad.append((char,tuple(round(v,1) for v in (x0,y0,x1,y1))))
 if bad:outside.append({'page':i+1,'count':len(bad),'samples':bad[:5]})
 payload=re.sub(r'Copyright@.*?bentuk apapun','',t,flags=re.S).strip()
 report.append({'page':i+1,'payload_chars':len(payload),'first':payload[:130].replace('\r\n',' ')})
 tp.close();p.close()
captions=set(int(x) for x in re.findall(r'Tabel\s+C\.(\d+)','\n'.join(alltext)))
assert captions==set(range(1,136)),('Missing captions in PDF',set(range(1,136))-captions)
assert not outside,('Text outside page',outside)
result={'pages':len(doc),'outside_page_text':outside,'C_table_captions':len(captions),'low_content_pages':[x for x in report if x['payload_chars']<70]}
(root/'publish_render/pdf-audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=True))
