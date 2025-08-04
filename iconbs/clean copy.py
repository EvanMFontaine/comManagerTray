from PIL import Image

# Open the ICO file (first frame)
ico = Image.open("image.png")
ico = ico.convert("RGBA")  # Ensure it has alpha

# Get pixel data
pixels = ico.load()
width, height = ico.size
print(f"Image size: {width}x{height}")
assert pixels is not None, "Failed to load pixel data"
# Process pixels
for y in range(height):
    for x in range(width):
        r, g, b, a = pixels[x, y]
        if b > 50 and r < 175 and g < 175:  # Check for blue pixels
            # pixels[x, y] = (255, 255, 255, 255)  # Pure white
            pixels[x, y] = (245, 245, 245, 245)  # Pure white (whiter than white?? stupid windows lol )
        else:
            pixels[x, y] = (0, 0, 0, 0)  # Fully transparent

# Save as ICO for transparency support
ico.save("comport_245.ico", format="ICO")
print("Saved cleaned image as comport.ico")


base = Image.new("RGBA", (66, 66), (0, 0, 0, 0))
base.paste(ico, (1, 1))
base.save("comport_padded.ico", format="ICO", sizes=[(64, 64)])


sizes = [16, 24, 32, 48, 64]
icons = [ico.resize((s, s), Image.Resampling.NEAREST) for s in sizes]
icons[0].save("comportManualResize.ico", format="ICO", sizes=[(s, s) for s in sizes])