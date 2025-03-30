# EfCon File Converter Tool

A terminal-based Python utility to bulk convert `.ppt/.pptx` and `.docx` files into PDF format using Microsoft Office (PowerPoint and Word). Clean, simple, and gets the job done.

---

## ✅ Features

- 📂 Converts all PowerPoint (`.ppt`, `.pptx`) files in a folder to PDF  
- 📝 Converts all Word (`.docx`) files in a folder to PDF  
- 🔢 Automatically names output files as `1.pdf`, `2.pdf`, etc.  
- 💡 Uses PowerPoint/Word COM interface via `pywin32`  
- 🛡 Handles fallback methods and errors

---

## 🛠 Requirements

- Windows OS  
- Microsoft Office (PowerPoint and Word installed)  
- Python 3.x  
- Python package: `pywin32`

Install dependencies:
```bash
pip install pywin32
