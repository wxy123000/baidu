from PIL import Image, ImageOps, ImageDraw
from pathlib import Path

folder = Path(r"D:\PycharmProjects\PythonProject6\tmp\competitor_report\pdf_pages2")
files = sorted(folder.glob("page-*.png"))
thumb_w = 330
margin = 16
cols = 3
thumbs = []
for f in files:
    im = Image.open(f).convert("RGB")
    h = int(im.height * thumb_w / im.width)
    im = im.resize((thumb_w, h))
    canvas = Image.new("RGB", (thumb_w + 2, h + 30), "white")
    canvas.paste(im, (1, 28))
    d = ImageDraw.Draw(canvas)
    d.text((8, 7), f.stem, fill="#222222")
    canvas = ImageOps.expand(canvas, border=1, fill="#BFC6CE")
    thumbs.append(canvas)
rows = (len(thumbs) + cols - 1) // cols
cell_h = max(t.height for t in thumbs)
sheet = Image.new("RGB", (cols*(thumb_w+4+margin)+margin, rows*(cell_h+margin)+margin), "#E9EDF2")
for i, t in enumerate(thumbs):
    x = margin + (i % cols)*(thumb_w+4+margin)
    y = margin + (i // cols)*(cell_h+margin)
    sheet.paste(t, (x, y))
sheet.save(folder.parent / "contact_sheet2.png")
print(folder.parent / "contact_sheet2.png")
