import logging

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from config import PDF_FONT, PDF_FONT_SIZE, PDF_MARGIN_LEFT, PDF_MARGIN_BOTTOM, PDF_LINE_HEIGHT

def create_pdf(lines, source_file, input_folder, output_folder):
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
    logging.info(f"PDF generated successfully: {pdf_file}")