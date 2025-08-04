from PIL import Image

SRC = "imageCleaned.png"          # your cleaned outline
OUT = "comport.ico"

base = Image.open(SRC).convert("RGBA")

# sizes Windows actually asks for in the tray
sizes = [64, 48, 40, 32, 24, 16]  # common sizes for icons
# build sharp versions with **nearest-neighbour** resampling
icons = [base.resize((s, s), Image.Resampling.NEAREST) for s in sizes]

# save them as *BMP* bitmaps inside the .ico – no PNG, no extra gamma
icons[0].save(
    OUT,
    format="ICO",
    bitmap_format="bmp"      # <- key line
)
print("Wrote", OUT)
