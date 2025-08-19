# pre alpha stage!!! 

A made-it-simple **Windows desktop app** (Tkinter + Python) that converts **bank statement PDFs** into clean **Excel tables**.  

✨ Works with both **text-based PDFs** and **scanned statements**.  

---

## 🔥 Features  

✅ **Picks a PDF** \
⚡ **Fast text parsing** with PyMuPDF  
🔍 **OCR fallback** with EasyOCR CPU/(GPU WIP)  
📊 Exports a clean **Excel Table** (`.xlsx`) with columns:  
- 📅 `Date`  
- 📝 `Particulars`  
- 💸 `Debit`  
- 💰 `Credit`  
- 📈 `Balance`  


## ⚙️ Installation  

1. Install Python **3.9+** from [python.org](https://www.python.org/downloads/)  
   > ☑️ During install, tick **“Add Python to PATH”**  

2. Install C++ VS **2015+** from [microsoft.com](https://www.microsoft.com/en-in/download//details.aspx?id=48145&msockid=38f389d4ee3663ec07d39f99ef4962d2/)  

3. Install required packages:  

```bash
python -m pip install --upgrade pip
pip install pymupdf pillow numpy pandas openpyxl
pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cpu
pip install easyocr

```


▶️ Usage
Clone this repo or download the source:


```bash
Copy
Edit
git clone https://github.com/pragadisan/pdf-to-excel.git
cd bank-pdf-to-excel
```
Run the app:

```
bash
Copy
Edit
python pdf_to_excel_ocr_gui.py
```

Steps inside the app:

🖱️ Click Pick PDF and select one statement

🌀 Click Convert to Excel

🎉 Wait for conversion → choose what to do next

📂 Output
Each PDF generates a matching Excel file:

Copy
Edit
Axis_June.pdf   →   Axis_June.xlsx
HDFC_July.pdf   →   HDFC_July.xlsx

🛠️ Tech Stack 
🐍 Python
🖼️ Tkinter (GUI)
📖 PyMuPDF (PDF text parsing)
👁️ EasyOCR (OCR)
📊 Pandas + OpenPyXL (Excel export)

🚧 Roadmap
 📑 Combine multiple PDFs → one Excel with multiple sheets
 🐍 GPU support
 🔧 Smarter regex tuning for different banks
 📂 Export to CSV as well as Excel

