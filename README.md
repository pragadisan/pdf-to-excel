# pre alpha stage!!! 

A made-it-simple **Windows desktop app** (Tkinter + Python) that converts **bank statement PDFs** into clean **Excel tables**.  

✨ Works with both **text-based PDFs** and **scanned statements**.  

---

## 🔥 Features  

✅ **Pick one** 
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


2. Install required packages:  

```bash
python -m pip install --upgrade pip
pip install pymupdf pillow numpy pandas openpyxl
pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cpu
pip install easyocr


```
👉 If you have a supported GPU (CUDA), install GPU-enabled PyTorch from pytorch.org. 


▶️ Usage
Clone this repo or download the source:


```bash
Copy
Edit
git clone https://github.com/yourusername/bank-pdf-to-excel.git
cd bank-pdf-to-excel
```
Run the app:

```
bash
Copy
Edit
python bank_pdf_to_excel_ocr_gui.py
```

Steps inside the app:

🖱️ Click Pick PDF(s) and select one or more statements

⚡ (Optional) Tick Use GPU if available

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
 🔧 Smarter regex tuning for different banks
 📂 Export to CSV as well as Excel
 🐳 Dockerized version

