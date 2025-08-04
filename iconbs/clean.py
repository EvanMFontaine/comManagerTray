from PIL import Image

# Open the ICO file (first frame)
ico = Image.open("comport.ico")
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
        if a < 200:
            pixels[x, y] = (0, 0, 0, 0)  # Fully transparent
        else:
            pixels[x, y] = (255, 255, 255, 255)  # Pure white

# # Save as ICO for transparency support
# ico.save("comport_clean.ico", format="ICO")
# print("Saved cleaned image as comport_clean.ico")

# Save only 1 size to prevent resizing
ico.save("comport_clean.ico", format="ICO", sizes=[(width, height)])
print("Saved cleaned image as comport_clean.ico with no downsampling")