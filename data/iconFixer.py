from PIL import Image
import os
# Load the uploaded image
input_path = "image.png"
image = Image.open(input_path)

# Convert to ICO with multiple sizes for Windows tray icon compatibility
ico_path = "data/serial_port_icon.ico"
sizes = [(64, 64)]
image.save(ico_path, format='ICO', sizes=sizes)

# Convert image to RGBA to support transparency
image = image.convert("RGBA")

# Get data and replace white background with transparency
datas = image.getdata()
assert datas is not None, "Image data could not be retrieved."
new_data = []
for item in datas:
    # Replace white (or near-white) with transparency
    if item[0] > 240 and item[1] > 240 and item[2] > 240:
        new_data.append((255, 255, 255, 0))
    else:
        new_data.append(item)

# Apply new data
image.putdata(new_data)

# Save the transparent icon
transparent_ico_path = "comport.ico"
image.save(transparent_ico_path, format='ICO', sizes=sizes)

