import win32com.client
from docx2pdf import convert
from pdf_generator import create_pdf


def txt_converter(file, input_folder, output_folder, logger):
    try:
        with open(file, "r", encoding="utf-8") as f:
            lines = f.readlines()
        create_pdf(lines, file, input_folder, output_folder, logger)
    except Exception as e:
        logger(f"✖ TXT Error ({file.name}): {e}")


def docx_converter(file, input_folder, output_folder, logger):
    try:
        relative_path = file.relative_to(input_folder)
        pdf_file = (output_folder / relative_path).with_suffix(".pdf")
        pdf_file.parent.mkdir(parents=True, exist_ok=True)

        convert(str(file), str(pdf_file))
        logger(f"PDF generated successfully (Word): {pdf_file.name}")
    except Exception as e:
        logger(f"✖ Word Error ({file.name}): {e}")


def xlsx_converter(file, input_folder, output_folder, logger):
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

        logger(f"PDF generated successfully (Excel): {pdf_file.name}")
    except Exception as e:
        logger(f"✖ Excel Error ({file.name}): {e}")
    finally:
        if excel is not None:
            excel.Quit()