# 📄 Universal PDF Converter

A powerful, modular, and lightweight Python tool that batch-converts Word (`.docx`), Excel (`.xlsx`), and Text (`.txt`) files into PDFs. The application processes files in alphabetical order and gives you full command-line control over the output destination, file filtering, and logging.

## ✨ Key Features

* **Native Conversion:** Uses COM interfaces (`win32com`, `docx2pdf`) to interact directly with Microsoft Office, ensuring 100% formatting accuracy for Word and Excel files.
* **Custom Output Routing:** Automatically creates a mirror `_PDF` folder, or lets you define an exact destination path using the `--output` argument.
* **Selective Processing:** Use command-line flags to easily ignore specific file types during batch processing.
* **Live Progress Tracking:** Features a clean, real-time console progress counter (e.g., `Progress: 15 / 30 files (50%)`) without cluttering the terminal.
* **Smart Sorting:** Automatically processes and converts your files in strict alphabetical order.
* **Advanced Logging System:** Dual-layer logging keeps the terminal clean while recording a detailed debug trace. You can define a custom log file location using the `--logfile` argument.
* **Non-Destructive:** Original files are never modified or moved.

## ⚠️ Prerequisites

* **Operating System:** Windows (Strictly required for Microsoft COM integration).
* **Software:** Microsoft Office (Word and Excel) must be physically installed on the machine.
* **Python:** Python 3.x.

## 🛠️ Installation

1. Clone this repository to your local machine:
   ```bash
   git clone [https://github.com/your-username/universal-pdf-converter.git](https://github.com/your-username/universal-pdf-converter.git)
   cd universal-pdf-converter
   2.Install the required Python dependencies:
   pip install reportlab docx2pdf pywin32

##🚀 Usage & Commands
Run the main script from your terminal. The only mandatory argument is the path to your source folder.

1. Basic Conversion (Automatic Output):
python main.py "C:\Path\To\Your\Documents"

3. Custom Output Directory:
Converts files and saves the PDFs to a specific folder.
python main.py "C:\Path\To\Your\Documents" --output "D:\My_PDF_Backups"

3. Selective Conversion (Ignore formats):
Skips all Text and Excel files, processing only Word documents.
python main.py "C:\Path\To\Your\Documents" --ignore-txt --ignore-xlsx

4. Custom Log File & Debug Mode:
Forces detailed logging and saves the log file to a specific location.
python main.py "C:\Path\To\Your\Documents" --debug --logfile "C:\Logs\conversion_trace.txt"

5. Combined Power Command:
python main.py "C:\Source" --output "E:\Destination" --ignore-docx --debug

6.Display Help Menu:
python main.py -h
