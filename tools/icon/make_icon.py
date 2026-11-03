# Draws the app icon (src/SpreadsheetEditor.App/Assets/AppIcon.ico and AppIcon.png). Development only, never shipped; needs
# Python 3 with Pillow. Run: python tools/icon/make_icon.py src/SpreadsheetEditor.App/Assets

import sys
from PIL import Image, ImageDraw, ImageFilter

S = 1024  # drawing size; downsampled for every icon size

def vgrad(size, top, bottom):
    w, h = size
    g = Image.new("RGBA", size)
    d = ImageDraw.Draw(g)
    for y in range(h):
        t = y / max(1, h - 1)
        c = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(4))
        d.line([(0, y), (w, y)], fill=c)
    return g

def rounded_mask(size, box, r):
    m = Image.new("L", size, 0)
    ImageDraw.Draw(m).rounded_rectangle(box, r, fill=255)
    return m

def shadow(size, box, r, blur, offset, alpha):
    sh = Image.new("RGBA", size, (0, 0, 0, 0))
    x0, y0, x1, y1 = box
    ImageDraw.Draw(sh).rounded_rectangle((x0, y0 + offset, x1, y1 + offset), r, fill=(0, 0, 0, alpha))
    return sh.filter(ImageFilter.GaussianBlur(blur))

img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

# Back: the spreadsheet, a green tile with a header band and grid lines.
sheet = (112, 96, 784, 824)
img.alpha_composite(shadow((S, S), sheet, 72, 24, 16, 70))
tile = vgrad((S, S), (33, 168, 112, 255), (11, 112, 74, 255))
img.paste(tile, (0, 0), rounded_mask((S, S), sheet, 72))
d = ImageDraw.Draw(img)
d.rounded_rectangle((112, 96, 784, 236), 72, fill=(16, 124, 82, 255))
d.rectangle((112, 180, 784, 236), fill=(16, 124, 82, 255))
grid = (255, 255, 255, 70)
for x in (336, 560):
    d.line([(x, 236), (x, 824)], fill=grid, width=10)
for y in (236, 384, 532, 680):
    d.line([(112, y), (784, y)], fill=grid, width=10)

# Front: a card, the row as a form, with a picture and two entries.
card = (352, 352, 928, 936)
img.alpha_composite(shadow((S, S), card, 64, 30, 22, 110))
cardfill = vgrad((S, S), (255, 255, 255, 255), (240, 243, 246, 255))
img.paste(cardfill, (0, 0), rounded_mask((S, S), card, 64))
d = ImageDraw.Draw(img)
d.rounded_rectangle(card, 64, outline=(0, 0, 0, 28), width=4)

# Picture: sky, sun, two hills, clipped to a rounded rectangle.
pic = (408, 408, 872, 676)
picimg = vgrad((S, S), (96, 190, 255, 255), (40, 130, 230, 255))
pd = ImageDraw.Draw(picimg)
pd.ellipse((744, 440, 816, 512), fill=(255, 214, 92, 255))
pd.polygon([(408, 676), (560, 520), (700, 676)], fill=(28, 150, 96, 255))
pd.polygon([(560, 676), (716, 548), (872, 676)], fill=(18, 118, 78, 255))
img.paste(picimg, (0, 0), rounded_mask((S, S), pic, 32))

# Entries: a label and a field, twice.
for y in (724, 828):
    d.rounded_rectangle((408, y, 540, y + 22), 11, fill=(120, 130, 140, 255))
    d.rounded_rectangle((408, y + 38, 872, y + 76), 14, fill=(255, 255, 255, 255), outline=(150, 160, 170, 255), width=5)

out = sys.argv[1]
img.resize((256, 256), Image.LANCZOS).save(f"{out}/AppIcon.png")
img.save(f"{out}/AppIcon.ico", sizes=[(16, 16), (20, 20), (24, 24), (32, 32), (40, 40), (48, 48), (64, 64), (256, 256)])
