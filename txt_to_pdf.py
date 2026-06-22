import logging
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from config_log_level import PDF_FONT, PDF_FONT_SIZE, PDF_MARGIN_LEFT, PDF_MARGIN_BOTTOM, PDF_LINE_HEIGHT


def create_pdf(lines, source_file, input_folder, output_folder):
    """
    Converts a list of text lines into a PDF document using ReportLab.

    :param lines: Lines of text read from the source file.
    :type lines: list[str]
    :param source_file: The full path to the source text file.
    :type source_file: Path
    :param input_folder: The root folder containing the original text files.
    :type input_folder: Path
    :param output_folder: The root folder where the generated PDF files will be saved.
    :type output_folder: Path
    :return: No value returned.
    :rtype: None
    """
    try:
        relative_path = source_file.relative_to(input_folder)
        pdf_file = (output_folder / relative_path).with_suffix(".pdf")
        pdf_file.parent.mkdir(parents=True, exist_ok=True)

        c = canvas.Canvas(str(pdf_file), pagesize=A4)
        width, height = A4


        cursor_y = height - 72

        c.setFont(PDF_FONT, PDF_FONT_SIZE)

        for line in lines:

            if cursor_y < PDF_MARGIN_BOTTOM:
                c.showPage()
                cursor_y = height - 72
                c.setFont(PDF_FONT, PDF_FONT_SIZE)

            clean_line = line.strip().replace('\t', '    ')
            c.drawString(PDF_MARGIN_LEFT, cursor_y, clean_line)
            cursor_y -= PDF_LINE_HEIGHT

        c.save()
        logging.info(f"PDF generated successfully (TXT): {pdf_file.name}")

    except Exception as e:
        logging.error(f"Failed to generate PDF for text file {source_file.name}: {e}")