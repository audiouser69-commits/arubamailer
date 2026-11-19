"""Verifica rapida su Windows: conversione docx, logo in vari formati, struttura MIME."""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docx
from docx.shared import RGBColor
from PIL import Image, ImageDraw
import mailer_engine as me

d = tempfile.mkdtemp()
im = Image.new("RGBA", (600, 200), (0, 0, 0, 0))
ImageDraw.Draw(im).ellipse((5, 5, 195, 195), fill=(200, 30, 30, 255))
logos = []
for ext in ("png", "jpg", "gif", "bmp", "webp", "tiff", "ico"):
    p = os.path.join(d, "logo." + ext)
    (im.convert("RGB") if ext in ("jpg", "bmp", "gif") else im).save(p)
    logos.append(p)
doc = docx.Document()
doc.add_heading("Titolo", 1)
p = doc.add_paragraph("Testo ")
r = p.add_run("rosso grassetto"); r.bold = True; r.font.color.rgb = RGBColor(200, 0, 0)
doc.add_paragraph("punto", style="List Bullet")
dp = os.path.join(d, "t.docx"); doc.save(dp)
for lp in logos:
    c = me.build_content("Oggetto", docx_path=dp, logo_path=lp)
    m = me.Account("io@example.com", "").build_message("a@example.com", c)
    types = [x.get_content_type() for x in m.walk()]
    assert "text/html" in types and "text/plain" in types and any(t.startswith("image/") for t in types), types
    print(os.path.basename(lp), "OK", types)
print(me.estimate(2000, 250, 6, 35))
print("SMOKE OK")
