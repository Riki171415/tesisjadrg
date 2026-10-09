from pathlib import Path
import pypdfium2 as pdf
from PIL import Image,ImageDraw
root=Path(__file__).parent;out=root/'publish_render';out.mkdir(exist_ok=True)
d=pdf.PdfDocument(str(root/'final-preview.pdf'));texts=[];summary=[]
print('Pages:',len(d),flush=True)
for start in range(0,len(d),4):
    sheet=Image.new('RGB',(1810,2600),'#d8d8d8');draw=ImageDraw.Draw(sheet)
    for j in range(start,min(start+4,len(d))):
        p=d[j];im=p.render(scale=1.5).to_pil();im.save(out/f'page-{j+1:03}.png')
        x=(j-start)%2*905;y=(j-start)//2*1300
        draw.text((x+10,y+4),f'Page {j+1}',fill='black');sheet.paste(im,(x,y+25))
        t=p.get_textpage().get_text_range();texts.append(t)
        summary.append({'page':j+1,'characters':len(t),'first':t[:140].replace('\r\n',' ')})
        p.close()
    sheet.save(out/f'proof-{start//4+1:02}.png')
(out/'text.txt').write_text('\n\f\n'.join(texts),encoding='utf-8')
(out/'pages.json').write_text(__import__('json').dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
print('Low text pages:',[(x['page'],x['first']) for x in summary if x['characters']<60],flush=True)
