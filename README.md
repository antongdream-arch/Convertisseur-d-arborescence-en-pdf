# 📄 Universal PDF Converter

A modular, lightweight Python tool that batch-converts Word (`.docx`), Excel (`.xlsx`), and Text (`.txt`) files into PDFs. The application automatically creates a "mirror" folder containing all the generated PDFs while strictly preserving your original directory structure.

## ✨ Key Features

* **Native Conversion:** Uses COM interfaces (`win32com`, `docx2pdf`) to interact directly with Microsoft Office, ensuring 100% formatting accuracy for Word and Excel files.
* **Batch Processing:** Scans an entire directory tree and converts all supported files automatically.
* **Non-Destructive:** Original files are never modified. The script builds a separate `_PDF` directory alongside your source folder.
* **Modular Architecture:** Clean, maintainable code split across specialized configuration, generation, and execution files.

## ⚠️ Prerequisites

* **Operating System:** Windows (Strictly required for Microsoft COM integration).
* **Software:** Microsoft Office (Word and Excel) must be physically installed on the machine.
* **Python:** Python 3.x.

## 🛠️ Installation

1. Clone this repository to your local machine:
   ```bash
   git clone [https://github.com/your-username/universal-pdf-converter.git](https://github.com/your-username/universal-pdf-converter.git)
   cd universal-pdf-converter
TODO UPDATE Readme with library to install