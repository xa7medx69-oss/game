from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

qa = Path(r"C:\Users\Ahmed\Documents\ChatGPT\games\physical-social-games-handbook\8-game-expansion-pack\qa-final")
pages = sorted(qa.glob("page-*.png"))
font = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 28)

for start in range(0, len(pages), 4):
    group = pages[start:start + 4]
    images = [Image.open(p).convert("RGB") for p in group]
    w = max(im.width for im in images)
    h = max(im.height for im in images)
    sheet = Image.new("RGB", (w * 2 + 60, h * 2 + 100), "#B8C1CA")
    draw = ImageDraw.Draw(sheet)
    for i, (path, im) in enumerate(zip(group, images)):
        x = 20 + (i % 2) * (w + 20)
        y = 50 + (i // 2) * (h + 20)
        sheet.paste(im, (x, y))
        draw.text((x + 10, y - 38), path.stem.upper(), font=font, fill="#102A43")
    end = start + len(group)
    sheet.save(qa / f"contact-{start+1:02d}-{end:02d}.jpg", quality=90)

print(f"Created {(len(pages)+3)//4} contact sheets for {len(pages)} pages")
