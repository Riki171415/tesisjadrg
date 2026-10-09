from pathlib import Path
from pypdf import PdfReader,PdfWriter
import json
root=Path(__file__).parent;meta=json.loads((root/'reference-page.json').read_text(encoding='utf-8-sig'))
base=PdfReader(root/'publish-preview.pdf');patch=PdfReader(root/'reference-page.pdf')
assert len(base.pages)==meta['pages'],'Unexpected pagination change'
assert len(patch.pages)==1,'Replacement must contain one page'
writer=PdfWriter()
for i,p in enumerate(base.pages):writer.add_page(patch.pages[0] if i==meta['page']-1 else p)
out=root/'publish-preview-final.pdf'
with out.open('wb') as f:writer.write(f)
print('Final preview:',len(base.pages),'pages; verified reference on page',meta['page'])
