import os
import re
import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont, ImageTk

# Globals
template_path = "certificate_template.jpg"
excel_path = "names.xlsx"
signature_path = None
font_path = "times.ttf"

BASE_FONT_SIZE = 72
MAX_WIDTH_RATIO = 0.8
EVENT_POS = (0.5, 0.40)
NAME_POS = (0.5, 0.50)
DATE_POS = (0.5, 0.58)
SIGN_POS = (0.75, 0.75)

output_dir = "certificates_test"
os.makedirs(output_dir, exist_ok=True)

def select_template():
    global template_path
    path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg;*.png;*.jpeg")])
    if path:
        template_path = path
        template_label.config(text=f"Selected: {os.path.basename(path)}")

def select_excel():
    global excel_path
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        excel_path = path
        excel_label.config(text=f"Selected: {os.path.basename(path)}")

def select_signature():
    global signature_path
    path = filedialog.askopenfilename(filetypes=[("Images", "*.png;*.jpg")])
    if path:
        signature_path = path
        signature_label.config(text=f"Selected: {os.path.basename(path)}")

def generate_certificates():
    if not template_path or not excel_path:
        messagebox.showerror("Error", "Select template and Excel file first.")
        return

    df = pd.read_excel(excel_path)

    required_cols = ["Name", "Event", "Date"]
    if not all(col in df.columns for col in required_cols):
        messagebox.showerror("Error", "Excel must contain: Name, Event, Date")
        return

    template = Image.open(template_path)
    W, H = template.size

    for _, row in df.iterrows():
        name = str(row["Name"])
        event = str(row["Event"])
        date = str(row["Date"])

        img = template.copy()
        draw = ImageDraw.Draw(img)

        # NAME – auto-fit
        font_size = BASE_FONT_SIZE
        font = ImageFont.truetype(font_path, font_size)
        bbox = draw.textbbox((0, 0), name, font=font)
        text_w = bbox[2] - bbox[0]

        while text_w > W * MAX_WIDTH_RATIO:
            font_size -= 2
            font = ImageFont.truetype(font_path, font_size)
            bbox = draw.textbbox((0, 0), name, font=font)
            text_w = bbox[2] - bbox[0]

        # Positions
        name_x = NAME_POS[0] * W - text_w / 2
        name_y = NAME_POS[1] * H
        draw.text((name_x, name_y), name, fill="black", font=font)

        # EVENT
        font_event = ImageFont.truetype(font_path, 55)
        event_w = draw.textbbox((0, 0), event, font=font_event)[2]
        draw.text((EVENT_POS[0] * W - event_w / 2, EVENT_POS[1] * H),
                  event, fill="black", font=font_event)

        # DATE
        font_date = ImageFont.truetype(font_path, 48)
        date_w = draw.textbbox((0, 0), date, font=font_date)[2]
        draw.text((DATE_POS[0] * W - date_w / 2, DATE_POS[1] * H),
                  date, fill="black", font=font_date)

        # SIGNATURE
        if signature_path:
            sign_img = Image.open(signature_path).convert("RGBA")
            sign_img = sign_img.resize((400, 150))
            sx = int(SIGN_POS[0] * W)
            sy = int(SIGN_POS[1] * H)
            img.paste(sign_img, (sx, sy), sign_img)

        safe_name = re.sub(r'[\\/*?:"<>|]', "", name)
        output_path = os.path.join(output_dir, f"{safe_name}.pdf")
        img.convert("RGB").save(output_path, "PDF", resolution=100)

    messagebox.showinfo("Success", "All certificates generated!")

# --- GUI Layout ---
root = Tk()
root.title("Certificate Generator")
root.geometry("650x450")

Label(root, text="Certificate Generator", font=("Arial", 20, "bold")).pack(pady=10)

Button(root, text="Select Template Image", command=select_template).pack()
template_label = Label(root, text="No file selected")
template_label.pack()

Button(root, text="Select Excel File", command=select_excel).pack()
excel_label = Label(root, text="No file selected")
excel_label.pack()

Button(root, text="Select Signature (Optional)", command=select_signature).pack()
signature_label = Label(root, text="No file selected")
signature_label.pack()

Button(root, text="Generate Certificates", command=generate_certificates,
       font=("Arial", 14), bg="green", fg="white").pack(pady=20)

root.mainloop()
