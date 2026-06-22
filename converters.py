import logging
import sys
from pathlib import Path
import shutil
import win32com.client
from docx2pdf import convert
from txt_to_pdf import create_pdf


def cant_convert(file):
    """
    Checks if the file extension is supported by the available converters.

    :param file: The file path to validate.
    :type file: Path
    :return: True if the file cannot be converted, False otherwise.
    :rtype: bool
    """
    allowed_extensions = [".xls", ".xlsx", ".txt", ".docx", ".doc"]

    if file.suffix.lower() not in allowed_extensions:
        logging.error(f"It's not a file I can convert: {file.name}")
        return True
    return False


def txt_converter(file, input_folder, output_folder):
    """
    Reads a text file and converts its content to a PDF document.

    :param file: The full path to the text file to be converted.
    :type file: Path
    :param input_folder: The root folder containing the original text files.
    :type input_folder: Path
    :param output_folder: The root folder where the generated PDF files will be saved.
    :type output_folder: Path
    :return: No value returned.
    :rtype: None
    """
    try:
        with open(file, "r", encoding="utf-8") as f:
            lines = f.readlines()
        create_pdf(lines, file, input_folder, output_folder)
    except Exception as e:
        logging.error(f"Failed to convert text file {file.name}: {e}")


def docx_converter(file, input_folder, output_folder):
    """
    Converts a Word document (.docx) to PDF while preserving the folder structure.

    If the conversion fails, the error is logged and the program exits.

    :param file: The full path to the Word file to be converted.
    :type file: Path
    :param input_folder: The root folder containing the original text files.
    :type input_folder: Path
    :param output_folder: The root folder where the generated PDF files will be saved.
    :type output_folder: Path
    :return: No value returned.
    :rtype: None
    """
    try:
        relative_path = file.relative_to(input_folder)
        pdf_file = (output_folder / relative_path).with_suffix(".pdf")
        pdf_file.parent.mkdir(parents=True, exist_ok=True)

        convert(str(file), str(pdf_file))
        logging.info(f"PDF generated successfully (Word .docx): {pdf_file}")
    except Exception as e:
        logging.error(f"Failed to convert Word file {file.name}: {e}")
        sys.exit(1)


def doc_converter(file, input_folder, output_folder):
    """
    Converts a legacy Word document (.doc) directly to PDF using COM automation.

    Requires Microsoft Word to be installed on a Windows machine. The application
    is launched in the background and safely closed after processing.

    :param file: The full path to the legacy Word file to be converted.
    :type file: Path
    :param input_folder: The root folder containing the original files.
    :type input_folder: Path
    :param output_folder: The root folder where the generated PDF files will be saved.
    :type output_folder: Path
    :return: No value returned.
    :rtype: None
    """
    word = None

    try:
        relative_path = file.relative_to(input_folder)
        pdf_file = (output_folder / relative_path).with_suffix(".pdf")
        pdf_file.parent.mkdir(parents=True, exist_ok=True)

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(str(file.resolve()))

        doc.SaveAs(str(pdf_file.resolve()), FileFormat=17)
        doc.Close(False)

        logging.info(f"PDF generated successfully (Legacy Word .doc): {pdf_file}")
    except Exception as e:
        logging.error(f"Failed to convert Legacy Word file {file.name}: {e}")
        sys.exit(1)

    finally:
        if word is not None:
            word.Quit()


def xlsx_converter(file, input_folder, output_folder):
    """
    Converts an Excel workbook (.xls, .xlsx) to PDF using COM automation.

    Requires Microsoft Excel to be installed on a Windows machine. The application
    is launched in the background and safely closed after processing.

    :param file: The full path to the Excel file to be converted.
    :type file: Path
    :param input_folder: The root folder containing the original Excel files.
    :type input_folder: Path
    :param output_folder: The root folder where the generated PDF files will be saved.
    :type output_folder: Path
    :return: No value returned.
    :rtype: None
    """
    excel = None
    try:
        relative_path = file.relative_to(input_folder)
        pdf_file = (output_folder / relative_path).with_suffix(".pdf")
        pdf_file.parent.mkdir(parents=True, exist_ok=True)

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(file.resolve()))
        for ws in wb.Worksheets:
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = 1

        wb.ExportAsFixedFormat(0, str(pdf_file.resolve()))
        wb.Close(False)

        logging.info(f"PDF generated successfully (Excel): {pdf_file}")
    except Exception as e:
        logging.error(f"Failed to convert Excel file {file.name}: {e}")
        sys.exit(1)

    finally:
        if excel is not None:
            excel.Quit()


def copy_file(file, input_folder, output_folder):
    """
    Copies an unconverted file (such as an existing PDF or an image)
    to the destination folder while preserving the directory structure.

    :param file: The full path to the file to be copied.
    :type file: Path
    :param input_folder: The root folder containing the original files.
    :type input_folder: Path
    :param output_folder: The root folder where the copied files will be saved.
    :type output_folder: Path
    :return: No value returned.
    :rtype: None
    """
    try:
        relative_path = file.relative_to(input_folder)
        dest_file = output_folder / relative_path
        dest_file.parent.mkdir(parents=True, exist_ok=True)

        shutil.copy2(file, dest_file)
        logging.info(f"Copied untouched file: {dest_file.name}")
    except Exception as e:
        logging.error(f"Failed to copy file {file.name}: {e}")
        sys.exit(1)
