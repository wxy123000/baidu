from PIL import Image,ImageOps,ImageDraw
from pathlib import Path
folder=Path(r'D:\PycharmProjects\PythonProject6\tmp\competitor_review_pages')
files=sorted(folder.glob('page-*.png'),key=lambda p:int(p.stem.split('-')[1]))
w=330; m=16; cols=3; ts=[]
for f in files:
 im=Image.open(f).convert('RGB'); h=int(im.height*w/im.width); im=im.resize((w,h))
 c=Image.new('RGB',(w+2,h+30),'white'); c.paste(im,(1,28)); ImageDraw.Draw(c).text((8,7),f.stem,fill='#222'); ts.append(ImageOps.expand(c,border=1,fill='#BBC3CC'))
rows=(len(ts)+cols-1)//cols; ch=max(x.height for x in ts); sheet=Image.new('RGB',(cols*(w+20)+m,rows*(ch+m)+m),'#E9EDF2')
for i,t in enumerate(ts): sheet.paste(t,(m+(i%cols)*(w+20),m+(i//cols)*(ch+m)))
out=folder.parent/'competitor_review_sheet.png'; sheet.save(out); print(out)
