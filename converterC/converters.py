import logging

import sys
import win32com.client
from docx2pdf import convert
from txt_to_pdf import create_pdf


def txt_converter(file, input_folder, output_folder):
    """launch the convert of the txt file"""
    with open(file, "r", encoding="utf-8") as f:
        lines = f.readlines()
    create_pdf(lines, file, input_folder, output_folder)


def docx_converter(file, input_folder, output_folder):
    """convert docx file to PDF"""
    try:
        relative_path = file.relative_to(input_folder)
        pdf_file = (output_folder / relative_path).with_suffix(".pdf")
        pdf_file.parent.mkdir(parents=True, exist_ok=True)

        convert(str(file), str(pdf_file))
        logging.info(f"PDF generated successfully (Word): {pdf_file}")
    except Exception as e:
        logging.error(f"Failed to convert Word file {file}: {e}")
        sys.exit(1)


def xlsx_converter(file, input_folder, output_folder):
    """convert xlsx file to PDF"""
    excel = None
    try:
        relative_path = file.relative_to(input_folder)
        pdf_file = (output_folder / relative_path).with_suffix(".pdf")
        pdf_file.parent.mkdir(parents=True, exist_ok=True)

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(file.resolve()))
        wb.ExportAsFixedFormat(0, str(pdf_file.resolve()))
        wb.Close(False)

        logging.info(f"PDF generated successfully (Excel): {pdf_file}")
    except Exception as e:
        logging.error(f"Failed to convert Excel file {file}: {e}")
        sys.exit(1)
    finally:
        if excel is not None:
            excel.Quit()