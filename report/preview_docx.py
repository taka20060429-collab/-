# -*- coding: utf-8 -*-
"""docx を読み、A4ページのPNGプレビューを生成（近似レンダラ）。"""
import sys, os, io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image, ImageDraw, ImageFont

FONT_R = "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"   # gothic
FONT_M = "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"  # fallback
DPI = 110
A4 = (int(8.27 * DPI), int(11.69 * DPI))
MARGIN = int(0.98 * DPI)
BLACK = (0, 0, 0)

docx_path = sys.argv[1]
outdir = sys.argv[2] if len(sys.argv) > 2 else "/tmp/docxprev"
os.makedirs(outdir, exist_ok=True)
doc = Document(docx_path)

_fc = {}
def font(sz, bold):
    k = (int(sz), bold)
    if k not in _fc:
        px = max(10, int(sz * DPI / 72))
        _fc[k] = ImageFont.truetype(FONT_R, px)
    return _fc[k]

pages = []
def new_page():
    img = Image.new("RGB", A4, (255, 255, 255))
    return img, ImageDraw.Draw(img)

img, d = new_page()
y = MARGIN
maxw = A4[0] - 2 * MARGIN

def flush():
    global img, d, y
    pages.append(img)
    img, d = new_page(); y = MARGIN
    return y

def wrap(text, fnt):
    lines, cur = [], ""
    for ch in text:
        t = cur + ch
        if d.textlength(t, font=fnt) > maxw and cur:
            lines.append(cur); cur = ch
        else:
            cur = t
    lines.append(cur)
    return lines

# iterate body elements in order (paragraphs)
from docx.oxml.ns import qn
body = doc.element.body
for child in body.iterchildren():
    if child.tag == qn('w:p'):
        from docx.text.paragraph import Paragraph
        p = Paragraph(child, doc)
        # page break?
        if 'w:br' in child.xml and 'page' in child.xml:
            y = flush()
        # gather runs
        runs = p.runs
        # image?
        has_img = False
        for r in runs:
            blips = r._element.findall('.//' + qn('a:blip'))
            for b in blips:
                rEmbed = b.get(qn('r:embed'))
                if rEmbed:
                    part = doc.part.related_parts.get(rEmbed)
                    if part is not None:
                        im = Image.open(io.BytesIO(part.blob)).convert('RGB')
                        w = int(maxw * 0.62)
                        h = int(im.height * w / im.width)
                        if y + h > A4[1] - MARGIN:
                            y = flush()
                        x = MARGIN + (maxw - w) // 2
                        img.paste(im.resize((w, h)), (x, y))
                        y += h + 6
                        has_img = True
        if has_img:
            continue
        text = p.text
        if not text.strip():
            y += int(11 * DPI / 72 * 0.6)
            continue
        r0 = runs[0] if runs else None
        sz = (r0.font.size.pt if r0 and r0.font.size else 10.5)
        bold = bool(r0.font.bold) if r0 else False
        fnt = font(sz, bold)
        align = p.alignment
        for line in wrap(text, fnt):
            lh = int(sz * DPI / 72 * 1.4)
            if y + lh > A4[1] - MARGIN:
                y = flush()
            tw = d.textlength(line, font=fnt)
            if align == WD_ALIGN_PARAGRAPH.CENTER:
                x = MARGIN + (maxw - tw) / 2
            elif align == WD_ALIGN_PARAGRAPH.RIGHT:
                x = MARGIN + maxw - tw
            else:
                x = MARGIN
            d.text((x, y), line, fill=BLACK, font=fnt)
            y += lh
        y += int(4 * DPI / 72)

pages.append(img)
for i, pg in enumerate(pages, 1):
    pg.save(os.path.join(outdir, f"page_{i:02d}.png"))
print("pages:", len(pages))
