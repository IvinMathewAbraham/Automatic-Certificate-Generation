#  Automatic Certificate Generation

This Python project automates the process of generating personalized certificates from a template image and a list of names stored in an Excel file.

---

## 📘 Overview

The script reads names from an Excel sheet (`names.xlsx`), places each name onto a certificate template image (`certificate_template.jpg`), and exports the personalized certificates as image files inside a folder called `certificates/`.

It uses the **Pillow (PIL)** library for image processing and **pandas** for reading Excel data.

---

## ⚙️ Features

✅ Automatically centers names on the certificate  
✅ Dynamically adjusts font size based on name length  
✅ Customizable vertical position, font, and text color  
✅ Outputs high-quality PNG certificates  
✅ Safe filename generation (removes invalid characters)  

---

## 📁 Project Structure

```
Automatic-Certificate-Generation/
│
├── certificate_template.jpg       # Your certificate background image
├── names.xlsx                     # Excel file with a 'Name' column
├── times.ttf                      # Font file used for writing names
├── certificates/                  # Auto-generated certificates will appear here
├── generate_certificates.py       # Main Python script
└── README.md                      # Project documentation
```

---

## 🧩 Requirements

Install the required libraries before running the script:

```bash
pip install pillow pandas openpyxl
```

---

## 🧠 Usage

1. **Prepare your files:**
   - Place your certificate template image as `certificate_template.jpg`
   - Create an Excel file named `names.xlsx` with a column titled `Name`
   - Add all recipient names under that column
   - Ensure a font file (e.g., `times.ttf`) exists in your project folder

2. **Run the script:**
   ```bash
   python generate_certificates.py
   ```

3. **Find your generated certificates:**
   All certificates will be saved in the `certificates/` folder as:
   ```
   Certificate_<Name>.png
   ```

---

## 🧰 Configuration Options

You can adjust these values in the script to fine-tune the layout:

| Variable | Description | Default |
|-----------|--------------|----------|
| `TEMPLATE_PATH` | Path to the certificate template image | `"certificate_template.jpg"` |
| `EXCEL_PATH` | Path to Excel file containing names | `"names.xlsx"` |
| `OUTPUT_DIR` | Folder where certificates are saved | `"certificates"` |
| `FONT_PATH` | Path to `.ttf` font file | `"times.ttf"` |
| `BASE_FONT_SIZE` | Initial font size before scaling | `26` |
| `MAX_WIDTH_RATIO` | Maximum allowed width for name text relative to template width | `0.8` |
| `VERTICAL_POS_RATIO` | Vertical position of the name (lower = higher text) | `0.40` |

---

## 🖼 Example

| Input Excel | Output Certificate |
|--------------|--------------------|
| ![Excel Example](https://via.placeholder.com/300x100?text=names.xlsx) | ![Certificate Example](https://via.placeholder.com/400x250?text=Generated+Certificate) |

---

## 🚀 Future Improvements
- Add support for event name, date, and signature placement  
- Export as PDF instead of PNG  
- Add a simple web interface for uploading templates and Excel files  

---

## 🧑‍💻 Author
**Ivin Mathew Abraham**  
[GitHub Repository](https://github.com/IvinMathewAbraham/Automatic-Certificate-Generation)

---

## 🪪 License
This project is licensed under the MIT License — feel free to modify and distribute it.
