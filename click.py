from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os
import re

# --- CONFIG ---
TEMPLATE_PATH = "certificate_template.jpg"
EXCEL_PATH = "names.xlsx"
OUTPUT_DIR = "certificates"
FONT_PATH = "times.ttf"  # Replace with your TTF font
BASE_FONT_SIZE = 72      # Starting font size
MAX_WIDTH_RATIO = 0.8    # Max width of text relative to certificate width
VERTICAL_POS_RATIO = 0.50  # Adjust vertical position

# Create output directory if it doesn't exist
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Load names from Excel
df = pd.read_excel(EXCEL_PATH)
names = df['Name']  # Ensure your Excel has a column named 'Name'

# Load certificate template
template = Image.open(TEMPLATE_PATH)
W, H = template.size

for name in names:
    # Create a copy of the template
    img = template.copy()
    draw = ImageDraw.Draw(img)
    
    # Prepare font
    font_size = BASE_FONT_SIZE
    font = ImageFont.truetype(FONT_PATH, font_size)
    
    # Adjust font size if name is too long
    bbox = draw.textbbox((0, 0), name, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    
    while text_w > W * MAX_WIDTH_RATIO:
        font_size -= 2
        font = ImageFont.truetype(FONT_PATH, font_size)
        bbox = draw.textbbox((0, 0), name, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
    
    # Calculate centered position
    x = (W - text_w) / 2
    y = H * VERTICAL_POS_RATIO
    
    # Draw the name
    draw.text((x, y), name, fill="black", font=font)
    
    # Clean name for filename
    safe_name = re.sub(r'[\\/*?:"<>|]', "", name)
    
    # Save certificate as PDF instead of PNG
    output_path = os.path.join(OUTPUT_DIR, f"{safe_name}.pdf")
    rgb_img = img.convert("RGB")  # Ensure RGB mode for PDF
    rgb_img.save(output_path, "PDF", resolution=100.0)
    
    print(f"Saved: {output_path}")

print("✅ All certificates generated as PDF!")
