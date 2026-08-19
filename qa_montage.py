from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

qa = Path(r"C:\Users\Ahmed\Documents\ChatGPT\games\physical-social-games-handbook\qa-final2")
pages = sorted(qa.glob("page-*.png"))
font = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 28)

for group_index in range(0, len(pages), 4):
    group = pages[group_index:group_index + 4]
    opened = [Image.open(p).convert("RGB") for p in group]
    w = max(im.width for im in opened)
    h = max(im.height for im in opened)
    sheet = Image.new("RGB", (w * 2 + 60, h * 2 + 100), "#B8C1CA")
    draw = ImageDraw.Draw(sheet)
    for j, (path, im) in enumerate(zip(group, opened)):
        x = 20 + (j % 2) * (w + 20)
        y = 50 + (j // 2) * (h + 20)
        sheet.paste(im, (x, y))
        draw.text((x + 10, y - 38), path.stem.upper(), font=font, fill="#102A43")
    start = group_index + 1
    end = group_index + len(group)
    sheet.save(qa / f"contact-{start:02d}-{end:02d}.jpg", quality=90)

print(f"Created {(len(pages)+3)//4} contact sheets for {len(pages)} pages")
