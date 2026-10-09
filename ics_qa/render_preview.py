import pypdfium2 as pdf,pathlib
from PIL import Image,ImageDraw
root=pathlib.Path(__file__).parent
d=pdf.PdfDocument(str(root/'ICS-preview.pdf'))
print('PAGES',len(d),flush=True)
for start in range(0,len(d),16):
    sheet=Image.new('RGB',(2480,3600),'#d4d4d4');draw=ImageDraw.Draw(sheet)
    for j in range(start,min(start+16,len(d))):
        page=d[j];im=page.render(scale=1).to_pil();x=(j-start)%4*620;y=(j-start)//4*900;sheet.paste(im,(x,y+24));draw.text((x+6,y+5),str(j+1),fill='black');page.close()
    sheet.save(root/f'sheet-{start//16+1:02}.png')
alltext=[]
for i in range(len(d)):
    p=d[i];t=p.get_textpage().get_text_range();alltext.append(t)
    if i<4 or 'KETENTUAN & PERHATIAN' in t:print(i+1,t[:450].replace('\r\n',' '),flush=True)
    p.close()
(root/'pdf-text.txt').write_text('\n\f\n'.join(alltext),encoding='utf-8')
