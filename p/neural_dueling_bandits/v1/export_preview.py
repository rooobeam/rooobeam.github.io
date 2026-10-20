# -*- coding: utf-8 -*-
"""Render this deck using installed PowerPoint in a separate hidden instance."""
from pathlib import Path
import json
import sys
import tempfile

ROOT = Path(__file__).resolve().parent
PREVIEW = ROOT / "preview"
PREVIEW.mkdir(exist_ok=True)
sys.path.insert(0, str(Path(tempfile.gettempdir()) / "codex-dueling-ppt-deps"))
from PIL import Image, ImageDraw, ImageFont
import pythoncom
import win32com.client

PPTX = ROOT / "Dueling_Bandits_Sections_1-4_v1.pptx"
PDF = ROOT / "Dueling_Bandits_Sections_1-4_v1.pdf"

pythoncom.CoInitialize()
app = None
pres = None
overflow = []
count = 0
try:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.DisplayAlerts = 1
    pres = app.Presentations.Open(str(PPTX), ReadOnly=True, Untitled=False, WithWindow=False)
    count = pres.Slides.Count
    for i in range(1, count+1):
        slide = pres.Slides(i)
        slide.Export(str(PREVIEW / f"slide-{i:02d}.png"), "PNG", 1600, 900)
        for j in range(1, slide.Shapes.Count+1):
            shp = slide.Shapes(j)
            if not shp.HasTextFrame or not shp.TextFrame.HasText:
                continue
            rng = shp.TextFrame2.TextRange
            if rng.BoundHeight > shp.Height + 2 or rng.BoundWidth > shp.Width + 2:
                overflow.append({"slide": i, "shape": shp.Name,
                                 "text": shp.TextFrame.TextRange.Text,
                                 "shape_pt": [round(shp.Width, 1), round(shp.Height, 1)],
                                 "text_pt": [round(rng.BoundWidth, 1), round(rng.BoundHeight, 1)]})
    pres.SaveAs(str(PDF), 32)
finally:
    if pres is not None:
        pres.Close()
    if app is not None:
        app.Quit()
    pythoncom.CoUninitialize()

deck = json.loads((ROOT / "slide_content.json").read_text(encoding="utf-8"))
thumb_w, thumb_h, gap, cap_h = 576, 324, 18, 46
cols, rows = 3, (count+2)//3
canvas = Image.new("RGB", (cols*(thumb_w+gap)+gap, rows*(thumb_h+cap_h+gap)+gap), "#E6EAF1")
draw = ImageDraw.Draw(canvas)
font = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 18)
for i in range(count):
    x = gap+(i%cols)*(thumb_w+gap)
    y = gap+(i//cols)*(thumb_h+cap_h+gap)
    im = Image.open(PREVIEW / f"slide-{i+1:02d}.png").convert("RGB")
    im = im.resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    canvas.paste(im, (x, y))
    label = f'{i+1:02d}  {deck[i]["title"]}'
    draw.text((x+4, y+thumb_h+10), label, fill="#16263B", font=font)
canvas.save(PREVIEW / "contact_sheet.png")
audit = {"renderer": "Microsoft PowerPoint", "slides_rendered": count,
         "pdf": str(PDF), "text_overflow": overflow}
(PREVIEW / "render_audit.json").write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(audit, ensure_ascii=False, indent=2))
