# 📄 DOC to PDF Converter (Python + COMTypes)

[![Python](https://img.shields.io/badge/Python-3.x-blue)](https://www.python.org/)
[![Windows](https://img.shields.io/badge/OS-Windows-lightgrey)](https://www.microsoft.com/windows)
[![Word](<https://img.shields.io/badge/Requires-Microsoft%20Word-orange>)](https://www.microsoft.com/microsoft-365/word)
[![License](https://img.shields.io/badge/License-MIT-green)](https://opensource.org/licenses/MIT)

A professional **Word document to PDF converter** using Python and Microsoft Word COM automation. Converts DOC and DOCX files to PDFs efficiently, including batch folder processing.

---

## 🚀 Features

| Feature            | Description                               |
| ------------------ | ----------------------------------------- |
| Convert DOC & DOCX | Automatically convert all files in folder |
| Batch processing   | Scan folders and subfolders (optional)    |
| Output directory   | PDFs saved in target folder               |
| Lightweight & Fast | Minimal code, fast conversion             |
| COM API            | Uses official Microsoft Word API          |

---

## 🛠️ Requirements

- Windows OS (COM automation)
- Microsoft Word installed
- Python 3.8+
- `comtypes` library

```bash
pip install comtypes
```

---

## 📦 Installation & Setup

### 1. Clone or Download Project

```bash
git clone https://github.com/kamrulislam27/DOC-to-PDF-Converter-Python-COMTypes.git
cd DOC-to-PDF-Converter-Python-COMTypes
```

### 2. Install Dependencies

```bash
pip install comtypes
```

### 3. Ensure Word is Installed

---

## ▶️ How to Use

1. Set source folder (DOC/DOCX) and target folder (PDF output) in script:

```python
source = r"D:\\sourceFolder"
target = r"D:\\sourceFolder\\targetFolder"
doc_to_pdf(source, target)
```

2. Run the script:

```bash
python docToPdf.py
```

3. Optional: Use command line args for file/folder:

```bash
python convert.py "C:/path/to/folder" --recursive
```

---

---

## 📁 Folder Structure

```
doc-to-pdf-converter/
│── docToPdf.py
│── README.md
└── docFolder/   # optional input folder
└── pdfFolder/  # optional output folder
```

---

## ⚠️ Notes & Limitations

- Word **must** be installed
- Only works on **Windows**
- PDF output quality depends on Word export settings

---

## 📄 License

This project is licensed under the **MIT License**. Modify and redistribute freely.

---

## 📬 Contact

📧 Email: kamrul@ahut.edu.cn
💻 GitHub: https://github.com/kamrulislam27
