import logging
import sys
from pathlib import Path
import shutil
import win32com.client
from docx2pdf import convert
from txt_to_pdf import create_pdf

logger = logging.getLogger("UniversalConverter")


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
        logger.error(f"It's not a file I can convert: {file.name}")
        return True
    return False


def ensure_utf8_and_read(file_path):
    """
    Reads a text file by testing multiple encodings.
    If the original encoding is not UTF-8, the file is converted and rewritten in UTF-8.

    :param file_path: The path to the source text file.
    :type file_path: Path
    :return: The lines of text from the file. Returns an empty list if decoding fails.
    :rtype: list[str]
    """
    encodings_to_try = ['utf-8', 'cp1252', 'iso-8859-1']
    content_lines = None
    used_encoding = None

    for enc in encodings_to_try:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                content_lines = f.readlines()
            used_encoding = enc
            break
        except UnicodeDecodeError:
            continue

    if content_lines is None:
        logger.error(f"Critical decoding failure for file: {file_path.name}")
        return []

    if used_encoding != 'utf-8':
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.writelines(content_lines)
            logger.info(f"File converted to UTF-8: {file_path.name} (Old encoding: {used_encoding})")
        except Exception as e:
            logger.error(f"Error rewriting {file_path.name} in UTF-8: {e}")

    return content_lines


def txt_converter(file, input_folder, output_folder):
    """
    Reads a text file (ensuring UTF-8 encoding) and converts its content to a PDF document.

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
        lines = ensure_utf8_and_read(file)

        if lines:
            create_pdf(lines, file, input_folder, output_folder)

    except Exception as e:
        logger.error(f"Failed to process text file {file.name}: {e}")


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
        logger.info(f"PDF generated successfully (Word .docx): {pdf_file}")
    except Exception as e:
        logger.error(f"Failed to convert Word file {file.name}: {e}")


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

        logger.info(f"PDF generated successfully (Legacy Word .doc): {pdf_file}")
    except Exception as e:
        logger.error(f"Failed to convert Legacy Word file {file.name}: {e}")

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
            ws.PageSetup.Orientation = 2

            ws.PageSetup.CenterHorizontally = True

            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = False

        wb.ExportAsFixedFormat(0, str(pdf_file.resolve()))
        wb.Close(False)

        logger.info(f"PDF generated successfully (Excel): {pdf_file}")
    except Exception as e:
        logger.error(f"Failed to convert Excel file {file.name}: {e}")

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
        logger.info(f"Copied untouched file: {dest_file.name}")
    except Exception as e:
        logger.error(f"Failed to copy file {file.name}: {e}")
